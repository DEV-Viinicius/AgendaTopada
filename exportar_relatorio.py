# -*- coding: utf-8 -*-
r"""
FASE 2 do Agenda Topada — Exporta o relatorio de Ordem de Servico do ILUX.

Sequencia automatizada (apos o direcionamento das O.S.):
    Esc (fecha a tela do direcionamento)
    -> menu 'Relatorios'  ->  'Ordem de Servico'
    -> clica no campo 'STATUS' (abre uma tela)
    -> clica no campo da lista e escolhe a ULTIMA opcao
    -> clica em 'Verificado'
    -> clica no simbolo do Excel  ->  abre a tela 'Salvar como'
    -> VOCE salva o arquivo na pasta do Agenda Topada (aparece um aviso na tela)

Como usar
---------
  python exportar_relatorio.py --calibrar
        Define as coordenadas dessas telas do ILUX e salva em
        config_relatorio.json. Faca isto UMA vez (ou quando a tela mudar).

  python exportar_relatorio.py
        Executa de verdade. Na 1a vez pede calibracao. Confirma antes.

Seguranca
---------
  - PyAutoGUI FAILSAFE: jogue o mouse para o CANTO SUPERIOR ESQUERDO p/ abortar.
  - Confirmacao obrigatoria antes de mexer no ILUX.
"""

import os
import sys
import json
import time
import argparse
from datetime import datetime

pyautogui = None
pyperclip = None


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = app_dir()
CONFIG_REL = os.path.join(APP_DIR, "config_relatorio.json")
CONFIG_TOPADA = os.path.join(APP_DIR, "config_topada.json")

# elementos das telas de relatorio que precisam de coordenada (x, y)
CAMPOS_COORD = [
    ("menu_relatorios",     "Menu 'Relatorios' na barra do ILUX"),
    ("item_ordem_servico",  "Item 'Ordem de Servico' dentro de Relatorios"),
    ("campo_status",        "Campo/coluna 'STATUS' (o clique que ABRE a tela)"),
    ("campo_lista_status",  "Na tela aberta: o CAMPO onde clica p/ abrir a lista de status"),
    ("opcao_ultima_status", "A ULTIMA opcao da lista de status"),
    ("btn_verificado",      "Botao/opcao 'Verificado'"),
    ("icone_excel",         "Simbolo do Excel (exportar)"),
]


def log(msg):
    print(msg, flush=True)


# ---------------------------------------------------------------------------
# GUI / config
# ---------------------------------------------------------------------------
def _ensure_gui():
    global pyautogui, pyperclip
    if pyautogui is None:
        import pyautogui as _pg
        import pyperclip as _pc
        _pg.FAILSAFE = True
        _pg.PAUSE = 0.3
        pyautogui = _pg
        pyperclip = _pc


def carregar_config(caminho):
    if os.path.exists(caminho):
        try:
            with open(caminho, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def salvar_config(cfg, caminho):
    with open(caminho, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


def config_completa(cfg):
    return all(chave in cfg for chave, _ in CAMPOS_COORD)


def esperar(segundos, dm=1.0):
    time.sleep(segundos * dm)


# ---------------------------------------------------------------------------
# Calibracao
# ---------------------------------------------------------------------------
def capturar_ponto(descricao):
    _ensure_gui()
    log(f"\n>>> {descricao}")
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
    log("CALIBRACAO (relatorio) — deixe o ILUX aberto e visivel.")
    log("Dica: deixe uma O.S. ja aberta e navegue ate a tela do relatorio")
    log("conforme for pedindo cada ponto.")
    log("=" * 64)
    for chave, desc in CAMPOS_COORD:
        cfg[chave] = capturar_ponto(desc)
    salvar_config(cfg, CONFIG_REL)
    log(f"\nCalibracao salva em {CONFIG_REL}")
    return cfg


# ---------------------------------------------------------------------------
# Salvar (tela 'Salvar como' do Windows) — QUEM SALVA E O USUARIO
# ---------------------------------------------------------------------------
# O script apenas ABRE a tela 'Salvar como' (clicando no icone do Excel). A
# escolha da pasta e o clique em 'Salvar' ficam com o USUARIO. Antes disso
# mostramos um aviso com a pasta certa (a do Agenda Topada) e o nome sugerido.
# Assim some o clique automatico no botao (fragil) e a espera fixa de 20s.
def _aviso_salvar_console(pasta, nome):
    """Aviso padrao (execucao pelo console): instrui e espera Enter."""
    log("")
    log("=" * 64)
    log(">>> AGORA SALVE O RELATORIO <<<")
    log("Na tela 'Salvar como' que abriu no ILUX, salve NESTA pasta:")
    log(f"    {pasta}")
    log(f"Nome sugerido: {nome}")
    log("Depois de clicar em 'Salvar', volte aqui.")
    log("=" * 64)
    input("Ja salvou? Pressione ENTER para continuar... ")


# ---------------------------------------------------------------------------
# Fluxo do relatorio
# ---------------------------------------------------------------------------
# nome do arquivo (SEM extensao: o ILUX decide se salva .xls ou .xlsx)
NOME_BASE = "Relatorio de Ordem de Serviço - Completo"


def pasta_saida():
    cfg = carregar_config(CONFIG_TOPADA)
    p = cfg.get("pasta_saida", APP_DIR)
    os.makedirs(p, exist_ok=True)
    return p


def encontrar_relatorio(pasta, base, desde=0):
    """
    Acha o relatorio salvo (.xls/.xlsx). Como quem salva e o usuario, o nome
    pode variar: primeiro tenta o nome sugerido; senao, pega o Excel mais
    recente salvo NESTA execucao (modificado a partir de 'desde'), ignorando a
    AGENDA_FILTRADA. Retorna o caminho ou None.
    """
    exatos, recentes = [], []
    for f in os.listdir(pasta):
        nome_sem_ext, ext = os.path.splitext(f)
        if ext.lower() not in (".xls", ".xlsx"):
            continue
        if nome_sem_ext.upper().startswith("AGENDA_FILTRADA"):
            continue  # e a saida da fase 3, nao o relatorio
        caminho = os.path.join(pasta, f)
        if nome_sem_ext == base:
            exatos.append(caminho)
        elif os.path.getmtime(caminho) >= desde:
            recentes.append(caminho)
    if exatos:
        return max(exatos, key=os.path.getmtime)
    return max(recentes, key=os.path.getmtime) if recentes else None


def registrar_relatorio(caminho):
    cfg = carregar_config(CONFIG_TOPADA)
    cfg["ultimo_relatorio"] = caminho
    salvar_config(cfg, CONFIG_TOPADA)


def exportar(dm=1.0, confirmar=True, avisar_salvar=None):
    _ensure_gui()
    cfg = carregar_config(CONFIG_REL)
    if not config_completa(cfg):
        log("Faltam coordenadas do relatorio. Vamos calibrar primeiro.")
        cfg = calibrar(cfg)

    pasta = pasta_saida()
    destino_base = os.path.join(pasta, NOME_BASE)  # sem extensao

    log("\n" + "=" * 64)
    log("EXPORTAR RELATORIO DE O.S. do ILUX")
    log(f"O arquivo sera salvo em:\n  {destino_base}  (.xls ou .xlsx)")
    log("ATENCAO: o script vai controlar o mouse/teclado no ILUX.")
    log("Deixe o ILUX aberto e visivel. Mouse no canto sup. esquerdo = abortar.")
    log("=" * 64)
    if confirmar:
        resp = input("Digite 'SIM' para comecar: ").strip().upper()
        if resp != "SIM":
            log("Cancelado.")
            return None

    log("Comecando em 5 segundos... clique na janela do ILUX.")
    time.sleep(5)

    inicio = time.time()  # marca o inicio p/ achar o arquivo salvo NESTA execucao
    try:
        # 0) fecha a tela do direcionamento
        pyautogui.press("esc")
        esperar(0.8, dm)

        # 1) menu Relatorios -> Ordem de Servico
        pyautogui.click(*cfg["menu_relatorios"])
        esperar(0.8, dm)
        pyautogui.click(*cfg["item_ordem_servico"])
        esperar(2.0, dm)  # aguarda abrir a tela do relatorio

        # 2) clica no campo STATUS (abre a tela)
        pyautogui.click(*cfg["campo_status"])
        esperar(1.2, dm)

        # 3) abre a lista e escolhe a ULTIMA opcao
        pyautogui.click(*cfg["campo_lista_status"])
        esperar(0.8, dm)
        pyautogui.click(*cfg["opcao_ultima_status"])
        esperar(0.6, dm)

        # 4) Verificado
        pyautogui.click(*cfg["btn_verificado"])
        esperar(2.0, dm)  # aguarda o relatorio filtrar/gerar

        # 5) icone do Excel -> abre a tela 'Salvar como'
        #    QUEM SALVA E O USUARIO: abrimos a tela, mostramos o aviso com a
        #    pasta certa e esperamos ele salvar (some o clique automatico).
        pyautogui.click(*cfg["icone_excel"])
        esperar(2.0, dm)  # tempo p/ a tela 'Salvar como' aparecer
        (avisar_salvar or _aviso_salvar_console)(pasta, NOME_BASE)

    except pyautogui.FailSafeException:
        log("ABORTADO pelo usuario (mouse no canto).")
        return None
    except Exception as e:
        log(f"ERRO durante a exportacao: {e}")
        return None

    # confirma que o arquivo apareceu (aceita .xls ou .xlsx)
    esperar(1.0, dm)
    achado = encontrar_relatorio(pasta, NOME_BASE, inicio)
    if achado:
        registrar_relatorio(achado)
        log(f"\nOK! Relatorio salvo: {achado}")
        return achado
    else:
        log("\nNao encontrei o arquivo salvo. Confira a tela 'Salvar como'.")
        log(f"(esperado: {destino_base}.xls ou .xlsx)")
        return None


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Exporta relatorio de O.S. do ILUX.")
    ap.add_argument("--calibrar", action="store_true", help="redefine as coordenadas")
    ap.add_argument("--delay", type=float, default=1.0, help="multiplicador dos tempos")
    ap.add_argument("--sim", action="store_true", help="pula a confirmacao 'SIM'")
    args = ap.parse_args()

    if args.calibrar:
        calibrar(carregar_config(CONFIG_REL))
        return

    exportar(args.delay, confirmar=not args.sim)


if __name__ == "__main__":
    main()
