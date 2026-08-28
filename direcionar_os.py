# -*- coding: utf-8 -*-
r"""
Direcionamento automatico de Ordens de Servico (O.S.) no sistema ILUX.

O que faz
---------
Le todos os PDFs da pasta NOVA AGENDA. O nome do arquivo (sem .pdf) e o nome do
tecnico. Em cada pagina extrai o numero da O.S. (campo "Numero", 5 digitos) e,
para cada O.S., executa no ILUX a sequencia:
    pesquisar O.S.  ->  botao direito na grade  ->  "Alterar Tecnico"
    ->  duplo clique no campo de tecnico  ->  digitar o nome  ->  selecionar
    ->  confirmar (Verificar)

Extracao do numero
------------------
Estes PDFs tem CAMADA DE TEXTO, entao o numero e lido direto (rapido e exato).
Se alguma pagina nao tiver texto, cai para OCR (EasyOCR) so nessa pagina.
O mesmo numero pode aparecer em varias paginas do mesmo PDF -> nao repetimos.

Modos de uso
------------
  python direcionar_os.py --simular
        So le os PDFs e mostra/loga os numeros de O.S. por tecnico.
        NAO mexe no ILUX. Use isto PRIMEIRO para conferir a leitura.

  python direcionar_os.py --calibrar
        (Re)define as coordenadas dos elementos do ILUX e salva em config.json.

  python direcionar_os.py
        Executa de verdade. Na 1a vez pede calibracao. Sempre confirma antes.

  python direcionar_os.py --limite 1
        Executa so a 1a O.S. (ideal para o primeiro teste real no ILUX).

Parametros uteis
----------------
  --pasta "CAMINHO"   pasta dos PDFs (padrao: Desktop\NOVA AGENDA)
  --limite N          processa no maximo N O.S. no total (teste seguro)
  --delay F           multiplicador dos tempos de espera (padrao 1.0)

Seguranca
---------
  - PyAutoGUI FAILSAFE: jogue o mouse para o CANTO SUPERIOR ESQUERDO para abortar.
  - Confirmacao obrigatoria antes de comecar a mexer no ILUX.
  - Log em CSV: log_direcionamento_AAAAMMDD_HHMM.csv
"""

import os
import re
import csv
import sys
import json
import time
import argparse
import unicodedata
from datetime import datetime

import fitz  # PyMuPDF

# pyautogui / pyperclip so sao necessarios na execucao real; importa preguicoso
pyautogui = None
pyperclip = None

# EasyOCR so entra em acao como reserva; importa preguicoso
_reader = None

PASTA_PADRAO = r"C:\Users\suporte06\Desktop\NOVA AGENDA"
CONFIG_JSON = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config_direcionar.json")
# imagem de referencia do item "Alterar Tecnico" (casamento de imagem, rapido)
ITEM_IMG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "item_alterar_tecnico.png")

# numero da O.S. = campo "Numero" com 5 digitos OU MAIS (o tamanho pode crescer)
RE_NUMERO = re.compile(r"N[uú]mero\s*[:.\-]?\s*(\d{5,})", re.IGNORECASE)
RE_5DIG = re.compile(r"\b(\d{5,})\b")

# elementos da tela do ILUX que precisam de coordenada (x, y)
CAMPOS_COORD = [
    ("campo_pesquisa",       "Campo de pesquisa 'Seq. O.S.' (onde digita o numero)"),
    ("btn_pesquisar",        "Botao 'Pesquisar' (a lupa)"),
    ("area_grid",            "Area da GRADE onde da o clique com botao DIREITO"),
    ("menu_alterar_tecnico", "Item 'Alterar Tecnico' do menu que abre no botao direito"),
    ("campo_tecnico",        "Campo de selecao do tecnico (onde da DUPLO clique)"),
    ("primeiro_resultado",   "Primeiro tecnico da lista depois de digitar o nome"),
    ("btn_verificar",        "Botao 'Verificar' / confirmar a alteracao"),
]


def log(msg):
    print(msg, flush=True)


# ---------------------------------------------------------------------------
# Extracao do numero da O.S.
# ---------------------------------------------------------------------------
def get_reader():
    """Carrega o EasyOCR so quando precisar (reserva para PDF sem texto)."""
    global _reader
    if _reader is None:
        import numpy as np  # noqa: F401  (usado indiretamente pelo easyocr)
        import easyocr
        log("  (carregando EasyOCR para pagina sem texto...)")
        _reader = easyocr.Reader(["pt"], gpu=False)
    return _reader


def numero_do_texto(texto):
    """Extrai o numero da O.S. do texto da pagina. Retorna str ou None."""
    m = RE_NUMERO.search(texto)
    if m:
        return m.group(1)
    return None


# ---------------------------------------------------------------------------
# Leitura do numero pela POSICAO do rotulo (cobre os 3 modelos de O.S.)
# ---------------------------------------------------------------------------
# Nestes formularios o numero fica na MESMA LINHA, logo a DIREITA de um rotulo,
# mas o rotulo e o lugar na folha mudam conforme o modelo. Ordem de preferencia
# (a mesma combinada com o usuario):
#   1) PADRAO (normal)  -> rotulo "Numero :"  (canto superior direito)
#   2) INSTALACAO       -> rotulo "Numero :"  (mesma marca, mais a esquerda)
#   3) LEITURA          -> rotulo "OS Nº :"   (a palavra-chave e "Nº")
# Se nenhum rotulo casar na pagina, ela e PULADA.
# O numero da O.S. tem 5 digitos OU MAIS (sem teto, pois o numero cresce com o
# tempo) e NAO comeca com zero. A trava do zero a esquerda nao corta nada (uma
# O.S. e sequencial e nunca comeca com 0) e e o que separa a O.S. da
# Insc.Estadual (017542200). As bordas \b evitam pegar um pedaco de um numero
# maior (ex.: nao casa "17542200" dentro de "017542200").
RE_SO_DIGITOS = re.compile(r"^[1-9]\d{4,}$")           # token isolado que E uma O.S.
RE_OS = re.compile(r"\b[1-9]\d{4,}\b")                 # O.S. dentro de um texto maior
# rotulo do modelo LEITURA: a palavra-chave e "Nº" (isolada ou grudada em
# "OS Nº : 968720", como costuma sair do OCR). Aceita º, ° e o degrau lido como o.
RE_ROTULO_LEITURA = re.compile(r"\bN[º°o]\b|\bN[º°]", re.IGNORECASE)


def _norm_palavra(s):
    """Minusculo e sem acento, para casar rotulos com seguranca."""
    nfkd = unicodedata.normalize("NFKD", s)
    return "".join(c for c in nfkd if not unicodedata.combining(c)).lower().strip()


def _numero_a_direita(itens, rotulo, ytol):
    """
    Menor numero (5+ digitos) que esteja na MESMA linha do rotulo e a DIREITA
    dele. 'itens' e uma lista de (texto, x0, y0, x1, y1); 'rotulo' e um item.
    """
    _, lx0, ly0, lx1, ly1 = rotulo
    ymeio = (ly0 + ly1) / 2.0
    cands = []
    for texto, x0, y0, x1, y1 in itens:
        if not RE_SO_DIGITOS.match(texto.strip()):
            continue
        if x0 < lx1 - 2:                          # tem de estar a direita do rotulo
            continue
        if abs((y0 + y1) / 2.0 - ymeio) > ytol:   # e na mesma linha
            continue
        cands.append((x0, texto.strip()))
    if cands:
        cands.sort()                              # o primeiro numero apos o rotulo
        return cands[0][1]
    return None


def _numero_do_rotulo(itens, rotulo, ytol):
    """
    Dado o item do rotulo, tenta o numero (a) DENTRO da propria caixa do rotulo
    (comum no OCR, ex.: 'Numero : 67742') e (b) na mesma linha, a DIREITA.
    """
    m = RE_OS.search(rotulo[0])
    if m:
        return m.group(0)
    return _numero_a_direita(itens, rotulo, ytol)


def _numero_por_rotulo(itens, ytol=6):
    """
    Aplica a regra dos 3 modelos sobre uma lista de (texto, x0,y0,x1,y1), venha
    ela da camada de texto do PDF ou do OCR. Retorna o numero (str) ou None.
    """
    # 1) e 2) modelos PADRAO e INSTALACAO: rotulo "Numero"
    for it in itens:
        if "numero" in _norm_palavra(it[0]):
            num = _numero_do_rotulo(itens, it, ytol)
            if num:
                return num
    # 3) modelo LEITURA: rotulo "Nº" (isolado ou grudado em "OS Nº")
    for it in itens:
        if RE_ROTULO_LEITURA.search(it[0]):
            num = _numero_do_rotulo(itens, it, ytol)
            if num:
                return num
    return None


def numero_por_geometria(page):
    """Le o numero pela posicao do rotulo, usando a camada de texto do PDF."""
    palavras = page.get_text("words")  # cada item: (x0, y0, x1, y1, texto, ...)
    if not palavras:
        return None
    itens = [(w[4], w[0], w[1], w[2], w[3]) for w in palavras]
    return _numero_por_rotulo(itens)


# Nos 3 modelos o rotulo do numero fica no TOPO da folha (ate ~29% da altura no
# modelo instalacao; normal e leitura ficam ~7-10%). Entao o OCR le so a FAIXA
# SUPERIOR da pagina (com folga), em vez da pagina inteira: muito mais rapido,
# sem perder o numero. Se um dia algum modelo trouxer o numero mais abaixo,
# basta aumentar esta fracao.
FRACAO_TOPO_OCR = 0.48
# resolucao do OCR (DPI). Menos DPI = OCR mais rapido; os numeros da O.S. sao
# grandes o bastante para serem lidos com folga em ~150 DPI.
DPI_OCR = 150


def numero_por_ocr(page):
    """
    Reserva para paginas SEM camada de texto (scan puro): rasteriza a FAIXA
    SUPERIOR da pagina, roda o OCR e aplica a MESMA regra dos 3 modelos
    (rotulo -> numero a direita). Se nao achar por rotulo, tenta um numero no
    canto sup. direito.
    """
    try:
        import numpy as np
        zoom = DPI_OCR / 72.0
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
        if pix.n == 4:
            img = img[:, :, :3]
        # corta so o topo da pagina (onde ficam os rotulos dos 3 modelos)
        topo = max(1, int(img.shape[0] * FRACAO_TOPO_OCR))
        img = np.ascontiguousarray(img[0:topo, :])
        achados = get_reader().readtext(img, detail=1, paragraph=False)
        # converte as caixas do OCR em itens (texto, x0, y0, x1, y1)
        itens = []
        for bbox, texto, conf in achados:
            xs = [p[0] for p in bbox]
            ys = [p[1] for p in bbox]
            itens.append((texto, min(xs), min(ys), max(xs), max(ys)))
        # tolerancia de linha maior no OCR (a altura do texto em pixels e grande)
        num = _numero_por_rotulo(itens, ytol=18)
        if num:
            return num
        # ultimo recurso: qualquer numero no canto superior direito
        h, w = img.shape[:2]
        for texto, x0, y0, x1, y1 in itens:
            if y0 < h * 0.22 and x0 > w * 0.50:
                m = RE_OS.search(texto)
                if m:
                    return m.group(0)
        return None
    except Exception as e:
        log(f"  OCR falhou: {e}")
        return None


# Cache da leitura: como os scans passam por OCR (lento), guardamos o resultado
# por arquivo, marcado por data-de-modificacao + tamanho. Se o PDF nao mudou,
# devolve na hora — isso evita reprocessar na Fase 3 o que a Fase 1 ja leu, e
# torna reexecucoes instantaneas. O cache NAO vai para o Git.
CACHE_JSON = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache_leitura_os.json")


def _assinatura_arquivo(caminho):
    try:
        st = os.stat(caminho)
        return f"{int(st.st_mtime)}-{st.st_size}"
    except OSError:
        return None


def _cache_carregar():
    try:
        with open(CACHE_JSON, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _cache_salvar(cache):
    try:
        with open(CACHE_JSON, "w", encoding="utf-8") as f:
            json.dump(cache, f)
    except Exception:
        pass


def extrair_os_do_pdf(caminho):
    """
    Retorna a lista de numeros de O.S. do PDF, na ordem das paginas, SEM repetir
    (o mesmo numero costuma vir em 2 paginas seguidas). Le o numero pela POSICAO
    do rotulo, cobrindo os 3 modelos (padrao, instalacao e leitura); se nenhum
    rotulo aparecer na pagina, ela e pulada. Usa cache por arquivo (evita OCR
    repetido quando o PDF nao mudou).
    """
    # 0) cache: se o arquivo nao mudou, devolve o resultado ja lido
    assinatura = _assinatura_arquivo(caminho)
    chave = os.path.abspath(caminho)
    cache = _cache_carregar()
    ent = cache.get(chave)
    if assinatura and ent and ent.get("sig") == assinatura:
        return list(ent.get("nums", []))

    numeros = []
    vistos = set()
    try:
        doc = fitz.open(caminho)
    except Exception as e:
        log(f"  ERRO ao abrir {caminho}: {e}")
        return numeros

    try:
        for i, page in enumerate(doc, 1):
            num = numero_por_geometria(page)          # 1) posicao do rotulo (3 modelos)
            if not num:                               # 2) regex simples (reserva)
                num = numero_do_texto(page.get_text() or "")
            if not num:                               # 3) OCR (pagina sem texto)
                num = numero_por_ocr(page)
            if not num:
                log(f"  pagina {i}: sem numero de O.S. valido (pulada)")
                continue
            if num in vistos:
                continue  # ja tratamos essa O.S. (pagina repetida)
            vistos.add(num)
            numeros.append(num)
    finally:
        doc.close()

    # guarda no cache para nao reprocessar este arquivo enquanto nao mudar
    if assinatura:
        cache[chave] = {"sig": assinatura, "nums": numeros}
        _cache_salvar(cache)
    return numeros


def listar_pdfs(pasta):
    return sorted(f for f in os.listdir(pasta) if f.lower().endswith(".pdf"))


def nome_tecnico(arquivo_pdf):
    return os.path.splitext(arquivo_pdf)[0].strip()


def coletar(pasta):
    """Le todos os PDFs e devolve lista de (tecnico, os_num, arquivo)."""
    tarefas = []
    for arq in listar_pdfs(pasta):
        tec = nome_tecnico(arq)
        log(f"\n{arq}  (tecnico: {tec})")
        oss = extrair_os_do_pdf(os.path.join(pasta, arq))
        if not oss:
            log("  nenhuma O.S. encontrada")
        for num in oss:
            log(f"  O.S. {num}")
            tarefas.append((tec, num, arq))
    return tarefas


# ---------------------------------------------------------------------------
# Calibracao das coordenadas do ILUX
# ---------------------------------------------------------------------------
def carregar_config():
    if os.path.exists(CONFIG_JSON):
        with open(CONFIG_JSON, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def salvar_config(cfg):
    with open(CONFIG_JSON, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)
    log(f"Coordenadas salvas em {CONFIG_JSON}")


def capturar_ponto(descricao, right_click_antes=None):
    """
    Captura uma coordenada por contagem regressiva (evita problema de foco no
    terminal). Se 'right_click_antes' for uma coord, da botao direito la antes
    de contar (para o menu de contexto ficar aberto na hora da captura).
    """
    _ensure_gui()
    log(f"\n>>> {descricao}")
    if right_click_antes:
        input("    (Enter: eu abro o menu com botao direito e conto 5s "
              "para voce posicionar o mouse no item) ")
        pyautogui.rightClick(right_click_antes[0], right_click_antes[1])
    else:
        input("    (Enter, e voce tera 5s para colocar o mouse em cima) ")
    for s in range(5, 0, -1):
        x, y = pyautogui.position()
        print(f"    capturando em {s}s...  mouse=({x},{y})   ", end="\r", flush=True)
        time.sleep(1)
    x, y = pyautogui.position()
    print(f"    capturado: ({x},{y})                         ")
    return [x, y]


def calibrar(cfg):
    _ensure_gui()
    log("=" * 64)
    log("CALIBRACAO — deixe o ILUX aberto e visivel.")
    log("Para cada item: leia a descricao, aperte Enter e posicione o mouse.")
    log("=" * 64)
    for chave, desc in CAMPOS_COORD:
        if chave == "menu_alterar_tecnico":
            # o menu so existe depois do clique direito na grade
            cfg[chave] = capturar_ponto(desc, right_click_antes=cfg.get("area_grid"))
        else:
            cfg[chave] = capturar_ponto(desc)
    salvar_config(cfg)
    log("\nCalibracao concluida.")
    return cfg


def config_completa(cfg):
    return all(chave in cfg for chave, _ in CAMPOS_COORD)


# ---------------------------------------------------------------------------
# Fluxo no ILUX
# ---------------------------------------------------------------------------
def _ensure_gui():
    global pyautogui, pyperclip
    if pyautogui is None:
        import pyautogui as _pg
        import pyperclip as _pc
        _pg.FAILSAFE = True   # mouse no canto sup. esquerdo aborta
        _pg.PAUSE = 0.3
        pyautogui = _pg
        pyperclip = _pc


def esperar(segundos, delay_mult):
    time.sleep(segundos * delay_mult)


def digitar_texto(texto):
    """Cola via clipboard (seguro para acentos)."""
    pyperclip.copy(texto)
    pyautogui.hotkey("ctrl", "v")


def _sem_acento(s):
    """Deixa minusculo e sem acento, para comparar textos com seguranca."""
    nfkd = unicodedata.normalize("NFKD", s)
    return "".join(c for c in nfkd if not unicodedata.combining(c)).lower()


def regiao_menu(clique):
    """
    Regiao da tela ao redor do ponto do clique direito onde o menu costuma
    abrir. Pega um pouco para cima e para baixo, porque conforme o status da
    O.S. o menu tem mais/menos itens e pode ate abrir para cima.
    """
    larg, alt = pyautogui.size()
    x, y = clique
    left = max(0, x - 30)
    top = max(0, y - 300)
    w = min(470, larg - left)
    h = min(720, alt - top)
    return (left, top, w, h)


def localizar_texto_na_tela(termos, regiao, escala=2):
    """
    Procura um item na tela pelo TEXTO (via OCR), dentro de 'regiao'
    (left, top, w, h). 'termos' e uma lista de textos aceitos, em ordem de
    preferencia (ja sem acento e minusculos). Retorna (x, y) absoluto do
    centro do texto encontrado, ou None se nao achar.
    """
    import numpy as np
    left, top, w, h = regiao
    shot = pyautogui.screenshot(region=(left, top, w, h))
    if escala != 1:
        shot = shot.resize((w * escala, h * escala))
    img = np.array(shot)  # RGB

    achados = get_reader().readtext(img, detail=1, paragraph=False)
    # normaliza e guarda o centro (ja convertido para coordenada absoluta)
    itens = []
    for bbox, texto, conf in achados:
        xs = [p[0] for p in bbox]
        ys = [p[1] for p in bbox]
        cx = left + (sum(xs) / len(xs)) / escala
        cy = top + (sum(ys) / len(ys)) / escala
        itens.append((_sem_acento(texto), int(cx), int(cy)))

    # tenta cada termo na ordem de preferencia
    for termo in termos:
        for texto, cx, cy in itens:
            if termo in texto:
                return (cx, cy)
    return None


class ItemNaoEncontrado(Exception):
    """O item 'Alterar Tecnico' nao foi encontrado no menu -> pular a O.S."""
    pass


def localizar_item_imagem(cfg):
    """
    Procura o item 'Alterar Tecnico' na tela por CASAMENTO DE IMAGEM (OpenCV).
    Muito mais rapido que OCR: acha a figurinha salva em ITEM_IMG dentro da
    regiao do menu, funcionando mesmo quando o menu muda de posicao.
    Retorna (x, y) do centro ou None (se nao achar ou nao houver imagem/ref).
    """
    if not os.path.exists(ITEM_IMG):
        return None
    try:
        regiao = regiao_menu(cfg["area_grid"])  # (left, top, w, h)
        p = pyautogui.locateCenterOnScreen(ITEM_IMG, region=regiao, confidence=0.75)
        return (int(p.x), int(p.y)) if p else None
    except Exception as e:
        log(f"  (busca por imagem falhou: {e})")
        return None


def clicar_alterar_tecnico(cfg, dm):
    """
    Da o clique direito na grade e clica em 'Alterar Tecnico'.
    Ordem de deteccao (rapido -> lento), funcionando mesmo quando o menu muda
    de posicao conforme o status da O.S.:
      1) casamento de imagem (OpenCV, rapido)
      2) OCR pelo NOME (EasyOCR, reserva)
    Se NENHUM achar, fecha o menu e levanta ItemNaoEncontrado (a O.S. e PULADA).
    """
    pyautogui.rightClick(*cfg["area_grid"])
    esperar(0.6, dm)

    ponto = localizar_item_imagem(cfg)  # 1) rapido, por imagem
    if not ponto:
        ponto = localizar_texto_na_tela(  # 2) reserva por OCR
            ["alterar tecnico", "tecnico"],
            regiao_menu(cfg["area_grid"]),
        )
    if ponto:
        pyautogui.click(*ponto)
    else:
        # Fecha o menu de contexto com um clique simples num lugar SEGURO
        # (o campo de pesquisa). NAO usar Esc: no ILUX o Esc costuma acionar
        # o botao "Cancelar" e FECHAR a tela inteira. Um clique fora do menu
        # apenas dispensa o menu, sem mexer no resto.
        pyautogui.click(*cfg["campo_pesquisa"])
        esperar(0.3, dm)
        raise ItemNaoEncontrado("'Alterar Tecnico' nao encontrado no menu")


def direcionar_os(os_num, tecnico, cfg, dm):
    """Executa a sequencia completa no ILUX para 1 O.S."""
    c = cfg

    # 1) pesquisa a O.S. (duplo clique seleciona o conteudo do campo)
    pyautogui.doubleClick(*c["campo_pesquisa"])
    pyautogui.press("delete")
    digitar_texto(os_num)
    esperar(0.3, dm)
    pyautogui.click(*c["btn_pesquisar"])
    esperar(2.0, dm)  # aguarda carregar o resultado

    # 2) botao direito na grade -> menu -> Alterar Tecnico
    #    (acha o item pelo NOME via OCR; a posicao muda conforme o status da O.S.)
    clicar_alterar_tecnico(c, dm)
    esperar(1.2, dm)  # aguarda abrir a janela de alteracao

    # 3) duplo clique no campo de tecnico, digita o nome, seleciona
    pyautogui.doubleClick(*c["campo_tecnico"])
    esperar(0.6, dm)
    digitar_texto(tecnico)
    esperar(1.0, dm)  # aguarda a lista filtrar
    pyautogui.doubleClick(*c["primeiro_resultado"])
    esperar(0.6, dm)

    # 4) confirma
    pyautogui.click(*c["btn_verificar"])
    esperar(1.5, dm)  # aguarda concluir


# ---------------------------------------------------------------------------
# Execucao
# ---------------------------------------------------------------------------
def abrir_log():
    nome = f"log_direcionamento_{datetime.now():%Y%m%d_%H%M}.csv"
    caminho = os.path.join(os.path.dirname(os.path.abspath(__file__)), nome)
    f = open(caminho, "w", newline="", encoding="utf-8-sig")
    w = csv.writer(f, delimiter=";")
    w.writerow(["hora", "arquivo", "tecnico", "os", "status"])
    return f, w, caminho


def executar(pasta, limite, dm):
    _ensure_gui()
    cfg = carregar_config()
    if not config_completa(cfg):
        log("Faltam coordenadas. Vamos calibrar primeiro.")
        cfg = calibrar(cfg)

    tarefas = coletar(pasta)
    if limite:
        tarefas = tarefas[:limite]
    if not tarefas:
        log("\nNada a fazer.")
        return

    log("\n" + "=" * 64)
    log(f"Serao direcionadas {len(tarefas)} O.S.:")
    for tec, num, _ in tarefas:
        log(f"  {num} -> {tec}")
    log("=" * 64)
    log("ATENCAO: o script vai controlar o mouse/teclado no ILUX.")
    log("Deixe o ILUX aberto e visivel. Mouse no canto sup. esquerdo = abortar.")
    resp = input("Digite 'SIM' para comecar: ").strip().upper()
    if resp != "SIM":
        log("Cancelado.")
        return

    log("Comecando em 5 segundos... clique na janela do ILUX.")
    time.sleep(5)

    f, w, caminho_log = abrir_log()
    ok = falhas = pulados = 0
    try:
        for i, (tec, num, arq) in enumerate(tarefas, 1):
            log(f"[{i}/{len(tarefas)}] O.S. {num} -> {tec}")
            try:
                direcionar_os(num, tec, cfg, dm)
                w.writerow([datetime.now().strftime("%H:%M:%S"), arq, tec, num, "ok"])
                ok += 1
            except pyautogui.FailSafeException:
                log("ABORTADO pelo usuario (mouse no canto).")
                w.writerow([datetime.now().strftime("%H:%M:%S"), arq, tec, num, "abortado"])
                break
            except ItemNaoEncontrado:
                log("  PULADO: 'Alterar Tecnico' nao encontrado -> proxima O.S.")
                w.writerow([datetime.now().strftime("%H:%M:%S"), arq, tec, num,
                            "pulado: 'Alterar Tecnico' nao encontrado"])
                pulados += 1
            except Exception as e:
                log(f"  ERRO: {e}")
                w.writerow([datetime.now().strftime("%H:%M:%S"), arq, tec, num, f"erro: {e}"])
                falhas += 1
            f.flush()
    finally:
        f.close()

    log("\n" + "=" * 64)
    log(f"Concluido. OK: {ok} | Pulados: {pulados} | Falhas: {falhas} | Log: {caminho_log}")


def capturar_item_menu(cfg):
    """
    Captura UMA VEZ a imagem do item 'Alterar Tecnico' para o casamento de
    imagem. Abre o menu com o botao direito e recorta um pedaco ao redor da
    coordenada calibrada 'menu_alterar_tecnico', salvando em ITEM_IMG.
    Importante: o recorte e feito com o mouse ainda na grade (item NAO
    destacado em azul), igual ao momento da busca real.
    """
    _ensure_gui()
    if "area_grid" not in cfg or "menu_alterar_tecnico" not in cfg:
        log("Faltam coordenadas. Rode 'python direcionar_os.py --calibrar' primeiro.")
        return
    log("=" * 64)
    log("CAPTURA da imagem do item 'Alterar Tecnico'.")
    log("Deixe o ILUX aberto com uma O.S. SELECIONADA na grade.")
    log("=" * 64)
    input("Enter para abrir o menu e capturar o item... ")
    log("Clicando com o botao direito em 3s...")
    time.sleep(3)
    pyautogui.rightClick(*cfg["area_grid"])
    esperar(0.6, 1.0)

    x, y = cfg["menu_alterar_tecnico"]
    left, top, w, h = max(0, x - 90), max(0, y - 15), 260, 30
    shot = pyautogui.screenshot(region=(left, top, w, h))
    shot.save(ITEM_IMG)
    log(f"Imagem salva em {ITEM_IMG}")
    log("Abra o arquivo e confira se 'Alterar Tecnico' aparece INTEIRO e nitido.")

    # confirma na hora que consegue reencontrar a imagem que acabou de salvar
    esperar(0.3, 1.0)
    p = localizar_item_imagem(cfg)
    if p:
        log(f"Teste OK: item localizado por imagem em {p}. Movendo o mouse para la (sem clicar).")
        pyautogui.moveTo(*p)
    else:
        log("NAO consegui reencontrar a imagem. Pode ser preciso aumentar a caixa "
            "(largura/altura) ou recapturar. Me avise.")
    log("Aperte Esc no ILUX para fechar o menu.")


def testar_menu(dm):
    """
    Testa (sem clicar) a deteccao do 'Alterar Tecnico' pelo nome.
    Da o clique direito na grade e move o mouse ate o item encontrado,
    para voce conferir visualmente se o OCR acertou.
    """
    _ensure_gui()
    cfg = carregar_config()
    if "area_grid" not in cfg:
        log("Faltam coordenadas. Rode 'python direcionar_os.py --calibrar' primeiro.")
        return
    log("=" * 64)
    log("TESTE da deteccao do 'Alterar Tecnico' pelo nome (nao clica).")
    log("Deixe o ILUX aberto com uma O.S. SELECIONADA na grade.")
    log("=" * 64)
    input("Enter para dar o clique direito e procurar o item... ")
    log("Clicando com o botao direito em 3s...")
    time.sleep(3)
    pyautogui.rightClick(*cfg["area_grid"])
    esperar(0.6, dm)
    ponto = localizar_item_imagem(cfg)
    origem = "imagem"
    if not ponto:
        ponto = localizar_texto_na_tela(
            ["alterar tecnico", "tecnico"],
            regiao_menu(cfg["area_grid"]),
        )
        origem = "OCR"
    if ponto:
        log(f"ACHOU 'Alterar Tecnico' por {origem} em {ponto}. Movendo o mouse para la (sem clicar).")
        pyautogui.moveTo(*ponto)
    else:
        log("NAO achou por imagem nem OCR. Na execucao real esta O.S. seria PULADA.")
    log("Teste concluido. Aperte Esc no ILUX para fechar o menu.")


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Direciona O.S. no ILUX a partir de PDFs.")
    ap.add_argument("--pasta", default=PASTA_PADRAO, help="pasta dos PDFs")
    ap.add_argument("--simular", action="store_true", help="so le os PDFs; nao mexe no ILUX")
    ap.add_argument("--calibrar", action="store_true", help="redefine as coordenadas do ILUX")
    ap.add_argument("--testar-menu", dest="testar_menu", action="store_true",
                    help="testa (sem clicar) a deteccao do 'Alterar Tecnico'")
    ap.add_argument("--capturar-item", dest="capturar_item", action="store_true",
                    help="captura a imagem do 'Alterar Tecnico' para busca rapida por imagem")
    ap.add_argument("--limite", type=int, default=0, help="processa no maximo N O.S.")
    ap.add_argument("--delay", type=float, default=1.0, help="multiplicador dos tempos de espera")
    args = ap.parse_args()

    if not os.path.isdir(args.pasta):
        log(f"ERRO: pasta nao encontrada: {args.pasta}")
        sys.exit(1)

    if args.calibrar:
        calibrar(carregar_config())
        return

    if args.capturar_item:
        capturar_item_menu(carregar_config())
        return

    if args.testar_menu:
        testar_menu(args.delay)
        return

    if args.simular:
        tarefas = coletar(args.pasta)
        log("\n" + "=" * 64)
        log(f"SIMULACAO: {len(tarefas)} O.S. seriam direcionadas.")
        for tec, num, _ in tarefas:
            log(f"  {num} -> {tec}")
        return

    executar(args.pasta, args.limite, args.delay)


if __name__ == "__main__":
    main()
