# -*- coding: utf-8 -*-
r"""
FASE 3 do Agenda Topada — Gera a AGENDA_FILTRADA a partir do Excel exportado
do ILUX, MAS ordenando as O.S. de cada tecnico na MESMA ORDEM dos PDFs da
pasta NOVA AGENDA (a ordem em que foram digitalizados/direcionados).

Regras (definidas com o usuario):
  - O nome do arquivo PDF (sem .pdf) = nome do tecnico (casamento exato,
    tolerante a maiusculas/acentos/espacos).
  - Para cada tecnico, as O.S. saem na ORDEM do PDF daquele tecnico.
  - O.S. que aparecem no Excel mas NAO estao em nenhum PDF: vao para o FIM
    da lista do tecnico, SEM numero de sequencia.

Baseado no AgendaTech_v3.py (mesmo layout/cores/blocos), so muda a ordenacao
e a numeracao da coluna "Posicao".
"""

import os
import re
import sys
import json
import unicodedata
from datetime import datetime, timedelta

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# ---------------------------------------------------------------------------
# Caminhos
# ---------------------------------------------------------------------------
def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = app_dir()
CONFIG_TOPADA = os.path.join(APP_DIR, "config_topada.json")
LOGO = os.path.join(APP_DIR, "LOGO SOLIVETTI.jpg")

# pasta padrao dos PDFs: subpasta "PDFS" DENTRO do projeto (portatil)
PASTA_PDFS_PADRAO = os.path.join(APP_DIR, "PDFS")

# lista fixa de tecnicos: controla a ORDEM das secoes e os blocos especiais
NOME_TECNICOS = [
    "CILAS", "ESDRAS", "JOSIEL", "ROMERO", "WILSON", "BRITO",
    "ROBERTO", "DIOGO", "ANDRE", "GUEDES", "RONALDO", "ADAUTO",
    "ALBERTO", "MILTON", "SIDRAYTONN", "GOMES", "TI",
    "JUNIOR", "JOAO", "NATALICIO", "FILIPE S",
]

CLIENTES_DESTACAR = [
    "PRESERVE SEGURANCA E TRANSPORTE DE VALORES LTDA",
    "PMC AUTOMOTIVA DO BRASIL LTDA",
    "BROSE DO BRASIL LTDA",
    "PROCURADORIA GERAL DA JUSTICA",
    "UMANA BRASIL ASSESSORIA E CONSULTORIA DE RECURSOS",
]


# ---------------------------------------------------------------------------
# Utilidades de normalizacao (para casar nomes e numeros com seguranca)
# ---------------------------------------------------------------------------
def sem_acento(s):
    nfkd = unicodedata.normalize("NFKD", str(s))
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def norm_nome(s):
    """Nome de tecnico normalizado: maiusculo, sem acento, sem espacos extras."""
    return re.sub(r"\s+", " ", sem_acento(s).strip().upper())


def norm_os(x):
    """Numero de O.S. normalizado: so digitos, sem zeros a esquerda."""
    d = re.sub(r"\D", "", str(x))
    d = d.lstrip("0")
    return d or "0"


# ---------------------------------------------------------------------------
# Leitura dos PDFs -> ordem das O.S. por tecnico
# ---------------------------------------------------------------------------
def carregar_config_topada():
    if os.path.exists(CONFIG_TOPADA):
        try:
            with open(CONFIG_TOPADA, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def salvar_config_topada(cfg):
    with open(CONFIG_TOPADA, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


def _importar_extrator(pasta_pdfs):
    """
    Reaproveita a extracao ja testada do direcionar_os.py (texto + OCR reserva).
    1) tenta importar direto (o menu ja pode ter deixado no sys.path);
    2) senao, procura o direcionar_os.py em locais conhecidos.
    """
    try:
        import direcionar_os as dnav  # ja importavel?
        return dnav
    except Exception:
        pass
    pasta_script = os.path.dirname(os.path.abspath(pasta_pdfs.rstrip("\\/")))
    candidatos = [
        APP_DIR, pasta_pdfs, pasta_script, PASTA_PDFS_PADRAO,
        r"C:\Users\suporte06\Desktop\NOVA AGENDA",
    ]
    for c in candidatos:
        if os.path.exists(os.path.join(c, "direcionar_os.py")):
            if c not in sys.path:
                sys.path.insert(0, c)
            import direcionar_os as dnav  # noqa
            return dnav
    return None


def ler_ordem_pdfs(pasta_pdfs):
    """
    Retorna dict {nome_tecnico_normalizado: [os_str, os_str, ...]} na ordem
    das paginas do PDF. Usa o extrator do direcionar_os se disponivel; senao,
    faz leitura simples pela camada de texto (fitz).
    """
    ordem = {}
    if not os.path.isdir(pasta_pdfs):
        return ordem

    dnav = _importar_extrator(pasta_pdfs)
    pdfs = sorted(f for f in os.listdir(pasta_pdfs) if f.lower().endswith(".pdf"))

    for arq in pdfs:
        tecnico = os.path.splitext(arq)[0].strip()
        caminho = os.path.join(pasta_pdfs, arq)
        if dnav is not None:
            nums = dnav.extrair_os_do_pdf(caminho)
        else:
            nums = _extrair_texto_simples(caminho)
        ordem[norm_nome(tecnico)] = [norm_os(n) for n in nums]
    return ordem


def _extrair_texto_simples(caminho):
    """Reserva local (sem OCR) caso o direcionar_os.py nao seja encontrado."""
    import fitz
    RE_NUM = re.compile(r"N[uú]mero\s*[:.\-]?\s*(\d{5,})", re.IGNORECASE)
    nums, vistos = [], set()
    try:
        doc = fitz.open(caminho)
    except Exception:
        return nums
    try:
        for page in doc:
            m = RE_NUM.search(page.get_text() or "")
            if m and m.group(1) not in vistos:
                vistos.add(m.group(1))
                nums.append(m.group(1))
    finally:
        doc.close()
    return nums


# ---------------------------------------------------------------------------
# Geracao da agenda
# ---------------------------------------------------------------------------
def gerar_agenda(arquivo_origem, pasta_pdfs, arquivo_saida=None):
    """
    Le o Excel do relatorio do ILUX e a ordem dos PDFs, e gera a
    AGENDA_FILTRADA.xlsx formatada. Retorna o caminho do arquivo gerado.
    """
    if arquivo_saida is None:
        arquivo_saida = os.path.join(os.path.dirname(arquivo_origem),
                                     "AGENDA_FILTRADA.xlsx")

    df_origem = pd.read_excel(arquivo_origem)
    df_origem["Técnico"] = df_origem["Técnico"].astype(str).str.strip()

    # ordem das O.S. por tecnico, vinda dos PDFs
    ordem_pdfs = ler_ordem_pdfs(pasta_pdfs)

    # relatorio de casamento (para avisar o usuario)
    avisos = []

    df_filtrado = df_origem[df_origem["Técnico"].isin(NOME_TECNICOS)]

    linhas_formatadas = []

    for tecnico in NOME_TECNICOS:
        df_tec = df_filtrado[df_filtrado["Técnico"] == tecnico].copy()
        if df_tec.empty:
            continue

        # lista de O.S. do PDF desse tecnico (na ordem digitalizada)
        ordem = ordem_pdfs.get(norm_nome(tecnico), [])
        pos = {num: i for i, num in enumerate(ordem)}
        if not ordem:
            avisos.append(f"Tecnico '{tecnico}': nenhum PDF correspondente "
                          f"(O.S. saem na ordem original, todas numeradas).")

        # ordena: primeiro as O.S. que estao no PDF (na ordem do PDF),
        # depois as extras (que nao estao no PDF) mantendo a ordem original.
        df_tec = df_tec.reset_index(drop=True)
        df_tec["_os"] = df_tec["Seq. O.S."].map(norm_os)
        # se nao ha PDF para o tecnico, considera TODAS como "no pdf" (numeradas)
        if ordem:
            df_tec["_no_pdf"] = df_tec["_os"].map(lambda o: o in pos)
            df_tec["_ordem"] = df_tec["_os"].map(lambda o: pos.get(o, 10 ** 9))
        else:
            df_tec["_no_pdf"] = True
            df_tec["_ordem"] = range(len(df_tec))
        df_tec["_orig"] = range(len(df_tec))
        df_tec = df_tec.sort_values(["_ordem", "_orig"], kind="stable")

        # cabecalho do tecnico
        linhas_formatadas.append(_linha_titulo(f"TÉCNICO: {tecnico}"))
        linhas_formatadas.append({
            "Sequência": "Posição", "Cliente": "Cliente", "ClienteCompleto": "",
            "Bairro": "Cidade/Bairro", "Número da O.S": "Número da O.S",
            "Tipo": "Status", "Técnico": "Técnico",
        })

        seq = 0
        for _, row in df_tec.iterrows():
            if row["_no_pdf"]:
                seq += 1
                numero = seq
            else:
                numero = ""  # O.S. fora do PDF: sem numero, vai para o fim
            linhas_formatadas.append({
                "Sequência": numero,
                "Cliente": str(row["Cliente (Razão)"])[:35],
                "ClienteCompleto": str(row["Cliente (Razão)"]),
                "Bairro": (str(row["Cidade"]) + " - " + str(row["Bairro"]))[:50],
                "Número da O.S": row["Seq. O.S."],
                "Tipo": str(row["Tipo de Status"])[:10],
                "Técnico": tecnico,
            })

        if tecnico == "ROBERTO":
            linhas_formatadas.append(
                _linha_titulo("Coordenador Responsável: MILTON / Técnicos internos:"))
        elif tecnico == "TI":
            linhas_formatadas.append(_linha_titulo("Portadores:"))

    df_final = pd.DataFrame(linhas_formatadas)
    _escrever_excel(df_final, arquivo_saida)
    return arquivo_saida, avisos


def _linha_titulo(texto):
    return {
        "Sequência": "", "Cliente": texto, "ClienteCompleto": "",
        "Bairro": "", "Número da O.S": "", "Tipo": "", "Técnico": "",
    }


def _escrever_excel(df_final, arquivo_saida):
    """Escreve o Excel formatado (mesmo layout do AgendaTech_v3)."""
    df_exportar = df_final.drop(columns=["ClienteCompleto"])

    with pd.ExcelWriter(arquivo_saida, engine="xlsxwriter") as writer:
        df_exportar.to_excel(writer, index=False, sheet_name="Agenda",
                             startrow=2, header=False)
        workbook = writer.book
        worksheet = writer.sheets["Agenda"]

        if os.path.exists(LOGO):
            worksheet.insert_image("A1", LOGO, {
                "x_offset": 20, "y_offset": 5, "x_scale": 0.15, "y_scale": 0.15,
            })

        f_cabecalho = workbook.add_format({
            "bold": True, "bg_color": "#4F81BD", "font_color": "white",
            "align": "center", "valign": "vcenter", "border": 1})
        f_tecnico = workbook.add_format({
            "bold": True, "bg_color": "#C6EFCE",
            "align": "center", "valign": "vcenter", "border": 1})
        f_dados = workbook.add_format({
            "align": "center", "valign": "vcenter", "border": 1})
        f_cliente_dest = workbook.add_format({
            "bold": True, "bg_color": "#FFD966",
            "align": "center", "valign": "vcenter", "border": 1})
        f_titulo = workbook.add_format({
            "bold": 1, "font_size": 16,
            "align": "center", "valign": "vcenter", "border": 1})
        f_data = workbook.add_format({
            "bold": 1, "align": "center", "valign": "vcenter",
            "bg_color": "#FFFB05", "border": 1, "font_color": "red"})
        f_esquerda = workbook.add_format({
            "bold": True, "align": "center", "valign": "vcenter",
            "bg_color": "#FCC7F3", "border": 1})

        worksheet.merge_range("B1:D2", "AGENDA", f_titulo)
        worksheet.merge_range(
            "A3:D3", "Coordenador Responsável: SIDRAYTONN / Técnicos internos:",
            f_esquerda)

        data_amanha = datetime.today() + timedelta(days=1)
        worksheet.merge_range("E1:E2", data_amanha.strftime("%d/%m/%Y"), f_data)
        worksheet.write("E3", _dia_semana_pt(data_amanha), f_data)

        for idx, coluna in enumerate(df_exportar.columns):
            worksheet.write(3, idx, coluna, f_cabecalho)

        linha_excel = 5
        for _, row in df_final.iterrows():
            cliente = str(row["Cliente"])
            if cliente.startswith("TÉCNICO:"):
                worksheet.merge_range(
                    f"A{linha_excel}:F{linha_excel}", cliente, f_tecnico)
            elif cliente in (
                "Coordenador Responsável: MILTON / Técnicos internos:",
                "Portadores:",
            ):
                worksheet.merge_range(
                    f"A{linha_excel}:F{linha_excel}", cliente, f_esquerda)
            else:
                nome_cli = sem_acento(row["ClienteCompleto"]).upper()
                dest = nome_cli in [sem_acento(c).upper() for c in CLIENTES_DESTACAR]
                f_cli = f_cliente_dest if dest else f_dados
                worksheet.write(f"A{linha_excel}", row["Sequência"], f_dados)
                worksheet.write(f"B{linha_excel}", row["Cliente"], f_cli)
                worksheet.write(f"C{linha_excel}", row["Bairro"], f_dados)
                worksheet.write(f"D{linha_excel}", row["Número da O.S"], f_dados)
                worksheet.write(f"E{linha_excel}", row["Tipo"], f_dados)
                worksheet.write(f"F{linha_excel}", row["Técnico"], f_dados)
            linha_excel += 1

        for idx, coluna in enumerate(df_exportar.columns):
            max_len = max(df_exportar[coluna].astype(str).map(len).max(),
                          len(str(coluna))) + 5
            worksheet.set_column(idx, idx, max_len)


_DIAS_PT = ["Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira",
            "Sexta-feira", "Sábado", "Domingo"]


def _dia_semana_pt(data):
    return _DIAS_PT[data.weekday()]


# ---------------------------------------------------------------------------
# Interface (usada quando roda sozinho ou chamado pelo menu)
# ---------------------------------------------------------------------------
def rodar_interativo(pasta_pdfs=None):
    cfg = carregar_config_topada()
    if pasta_pdfs is None:
        pasta_pdfs = cfg.get("pasta_pdfs", PASTA_PDFS_PADRAO)

    inicial = cfg.get("ultimo_relatorio", "")
    arquivo_origem = filedialog.askopenfilename(
        title="Selecione o Excel exportado do ILUX (relatório de O.S.)",
        initialfile=os.path.basename(inicial) if inicial else "",
        initialdir=os.path.dirname(inicial) if inicial else APP_DIR,
        filetypes=[("Excel files", "*.xls;*.xlsx")],
    )
    if not arquivo_origem:
        return

    try:
        saida, avisos = gerar_agenda(arquivo_origem, pasta_pdfs)
        cfg["ultimo_relatorio"] = arquivo_origem
        cfg["pasta_pdfs"] = pasta_pdfs
        salvar_config_topada(cfg)
        msg = f"Arquivo gerado:\n{saida}"
        if avisos:
            msg += "\n\nAvisos:\n- " + "\n- ".join(avisos)
        messagebox.showinfo("Sucesso", msg)
    except Exception as e:
        messagebox.showerror("Erro", str(e))


def main():
    root = tk.Tk()
    root.withdraw()
    rodar_interativo()


if __name__ == "__main__":
    main()
