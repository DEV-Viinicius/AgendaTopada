# -*- coding: utf-8 -*-
r"""
AGENDA TOPADA — um programa, UM clique.

Botao "GERAR AGENDA COMPLETA": roda as 3 fases em sequencia, sozinho:
  1) direciona todas as O.S. no ILUX (le os PDFs da pasta NOVA AGENDA)
  2) exporta o relatorio de Ordem de Servico do ILUX
  3) gera a AGENDA_FILTRADA.xlsx (cada tecnico na ordem do PDF)

Calibracao (config unica, feita 1 vez): botoes menores de "Calibrar".
Durante a execucao o programa controla o mouse/teclado; para ABORTAR a
qualquer momento jogue o mouse para o CANTO SUPERIOR ESQUERDO da tela.
"""

import os
import sys
import json
import time
import queue
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = app_dir()
CONFIG_TOPADA = os.path.join(APP_DIR, "config_topada.json")
LOGO = os.path.join(APP_DIR, "LOGO SOLIVETTI.jpg")

# pasta dos PDFs: subpasta "PDFS" DENTRO do projeto (portatil)
PASTA_PDFS_PADRAO = os.path.join(APP_DIR, "PDFS")
os.makedirs(PASTA_PDFS_PADRAO, exist_ok=True)

# onde vive o direcionar_os.py (Fase 1)
# Agora ele fica JUNTO do projeto (portatil). Mantemos a pasta antiga
# "NOVA AGENDA" apenas como reserva, para instalacoes antigas.
PASTA_NOVA_AGENDA = r"C:\Users\suporte06\Desktop\NOVA AGENDA"
if os.path.exists(os.path.join(APP_DIR, "direcionar_os.py")):
    PASTA_DIRECIONAR = APP_DIR
else:
    PASTA_DIRECIONAR = PASTA_NOVA_AGENDA
DIRECIONAR_PY = os.path.join(PASTA_DIRECIONAR, "direcionar_os.py")
EXPORTAR_PY = os.path.join(APP_DIR, "exportar_relatorio.py")

# importa os modulos como biblioteca (sem abrir consoles)
sys.path.insert(0, APP_DIR)
sys.path.insert(0, PASTA_DIRECIONAR)
import gerar_agenda_topada as fase3          # noqa: E402
import exportar_relatorio as fase2mod        # noqa: E402
try:
    import direcionar_os as dnav             # noqa: E402
except Exception as _e:
    dnav = None
    _erro_dnav = str(_e)


# ---------------------------------------------------------------------------
# Config (pastas)
# ---------------------------------------------------------------------------
def carregar_cfg():
    if os.path.exists(CONFIG_TOPADA):
        try:
            with open(CONFIG_TOPADA, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def salvar_cfg(cfg):
    with open(CONFIG_TOPADA, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


def get_pasta_pdfs():
    return carregar_cfg().get("pasta_pdfs", PASTA_PDFS_PADRAO)


def get_pasta_saida():
    return carregar_cfg().get("pasta_saida", APP_DIR)


def _python():
    if getattr(sys, "frozen", False):
        return "python"
    return sys.executable


# ---------------------------------------------------------------------------
# Log na tela (thread-safe via fila)
# ---------------------------------------------------------------------------
_fila = queue.Queue()


class _EscritorLog:
    """Redireciona print() para a fila da interface."""
    def write(self, txt):
        if txt:
            _fila.put(txt)

    def flush(self):
        pass


def log(msg):
    _fila.put(str(msg) + "\n")


def _drenar_fila():
    try:
        while True:
            txt = _fila.get_nowait()
            caixa_log.configure(state="normal")
            caixa_log.insert("end", txt)
            caixa_log.see("end")
            caixa_log.configure(state="disabled")
    except queue.Empty:
        pass
    janela.after(120, _drenar_fila)


# ---------------------------------------------------------------------------
# O fluxo completo (roda numa thread separada)
# ---------------------------------------------------------------------------
def _tem_calibracao():
    if dnav is None:
        log("ERRO: nao encontrei o direcionar_os.py em " + PASTA_DIRECIONAR)
        return False
    if not dnav.config_completa(dnav.carregar_config()):
        log(">> Falta calibrar o DIRECIONAMENTO. Clique em 'Calibrar direcionamento' (uma vez).")
        return False
    if not fase2mod.config_completa(fase2mod.carregar_config(fase2mod.CONFIG_REL)):
        log(">> Falta calibrar o RELATORIO. Clique em 'Calibrar relatorio' (uma vez).")
        return False
    return True


def _fase1_direcionar(pasta, cfg_dir, dm):
    dnav._ensure_gui()
    tarefas = dnav.coletar(pasta)
    if not tarefas:
        log("Nenhuma O.S. encontrada nos PDFs de " + pasta)
        return False
    log("=" * 60)
    log(f"FASE 1 — direcionando {len(tarefas)} O.S. no ILUX")
    log("Deixe o ILUX na frente. Mouse no canto sup. esquerdo = abortar.")
    log("Comecando em 5 segundos... CLIQUE na janela do ILUX agora!")
    time.sleep(5)
    ok = pulados = falhas = 0
    for i, (tec, num, arq) in enumerate(tarefas, 1):
        log(f"[{i}/{len(tarefas)}] O.S. {num} -> {tec}")
        try:
            dnav.direcionar_os(num, tec, cfg_dir, dm)
            ok += 1
        except dnav.ItemNaoEncontrado:
            log("   PULADO: 'Alterar Tecnico' nao encontrado")
            pulados += 1
        except dnav.pyautogui.FailSafeException:
            log("ABORTADO pelo usuario (mouse no canto).")
            raise
        except Exception as e:
            log(f"   ERRO: {e}")
            falhas += 1
    log(f"Fase 1 concluida. OK: {ok} | Pulados: {pulados} | Falhas: {falhas}")
    return True


def _avisar_salvar_gui(pasta, nome):
    """
    Mostra na tela (pop-up) o aviso de onde salvar o relatorio e ESPERA o
    usuario clicar OK antes de continuar. Como o pipeline roda numa thread, a
    janela e criada na thread principal (janela.after) e sincronizamos com um
    Event.
    """
    evento = threading.Event()

    def mostrar():
        messagebox.showinfo(
            "Salve o relatório",
            "A tela 'Salvar como' abriu no ILUX.\n\n"
            f"Salve o arquivo NESTA pasta:\n{pasta}\n\n"
            f"Nome sugerido:\n{nome}\n\n"
            "Depois de clicar em 'Salvar', clique OK aqui para continuar.")
        evento.set()

    janela.after(0, mostrar)
    evento.wait()


def _pipeline(dm=1.0):
    antigo_stdout = sys.stdout
    sys.stdout = _EscritorLog()
    try:
        if not _tem_calibracao():
            return
        pasta = get_pasta_pdfs()

        # FASE 1
        if not _fase1_direcionar(pasta, dnav.carregar_config(), dm):
            return

        # FASE 2
        log("=" * 60)
        log("FASE 2 — exportando relatorio do ILUX")
        caminho = fase2mod.exportar(dm, confirmar=False, avisar_salvar=_avisar_salvar_gui)
        if not caminho:
            log("Falha ao exportar o relatorio. Parei aqui.")
            return

        # FASE 3
        log("=" * 60)
        log("FASE 3 — gerando a AGENDA_FILTRADA")
        saida, avisos = fase3.gerar_agenda(caminho, pasta)
        for a in avisos:
            log("  aviso: " + a)
        log("=" * 60)
        log("PRONTO! Agenda gerada em:")
        log("  " + saida)
        janela.after(0, lambda: messagebox.showinfo(
            "Concluido", f"Agenda gerada:\n{saida}"))
    except Exception as e:
        log("ERRO geral: " + str(e))
        janela.after(0, lambda: messagebox.showerror("Erro", str(e)))
    finally:
        sys.stdout = antigo_stdout
        janela.after(0, lambda: btn_rodar.config(
            state="normal", text="▶  GERAR AGENDA COMPLETA"))


def gerar_agenda_completa():
    if not messagebox.askyesno(
        "Gerar agenda completa",
        "O programa vai, SOZINHO:\n\n"
        "  1) direcionar todas as O.S. no ILUX\n"
        "  2) exportar o relatorio\n"
        "  3) gerar a AGENDA_FILTRADA\n\n"
        "Deixe o ILUX aberto e visivel. Nao mexa no mouse durante a execucao.\n"
        "Para abortar: jogue o mouse para o CANTO SUPERIOR ESQUERDO.\n\n"
        "Comecar?"):
        return
    btn_rodar.config(state="disabled", text="Executando... (mouse no canto = abortar)")
    threading.Thread(target=_pipeline, daemon=True).start()


# ---------------------------------------------------------------------------
# Extras (config unica / fallback)
# ---------------------------------------------------------------------------
def rodar_console(script, *args, cwd=None):
    if not os.path.exists(script):
        messagebox.showerror("Arquivo nao encontrado", script)
        return
    cmd = [_python(), script, *args]
    creation = subprocess.CREATE_NEW_CONSOLE if os.name == "nt" else 0
    subprocess.Popen(cmd, cwd=cwd or os.path.dirname(script), creationflags=creation)


def calibrar_direcionamento():
    rodar_console(DIRECIONAR_PY, "--calibrar", cwd=PASTA_NOVA_AGENDA)


def calibrar_relatorio():
    rodar_console(EXPORTAR_PY, "--calibrar", cwd=APP_DIR)


def so_gerar_agenda():
    """Fallback: ja tenho o relatorio; so quero gerar a agenda."""
    fase3.rodar_interativo(pasta_pdfs=get_pasta_pdfs())


def escolher_pasta_pdfs():
    cfg = carregar_cfg()
    p = filedialog.askdirectory(title="Pasta dos PDFs (NOVA AGENDA)",
                                initialdir=get_pasta_pdfs())
    if p:
        cfg["pasta_pdfs"] = p
        salvar_cfg(cfg)
        lbl_pdfs.config(text=f"PDFs: {p}")


def escolher_pasta_saida():
    cfg = carregar_cfg()
    p = filedialog.askdirectory(title="Pasta onde salvar o relatorio/agenda",
                                initialdir=get_pasta_saida())
    if p:
        cfg["pasta_saida"] = p
        salvar_cfg(cfg)
        lbl_saida.config(text=f"Saida: {p}")


# ---------------------------------------------------------------------------
# Interface
# ---------------------------------------------------------------------------
def construir_janela():
    global janela, caixa_log, btn_rodar, lbl_pdfs, lbl_saida
    janela = tk.Tk()
    janela.title("Agenda Topada — Solivetti")
    janela.geometry("560x620")
    janela.configure(bg="white")

    topo = tk.Frame(janela, bg="white")
    topo.pack(fill="x", pady=(10, 4))
    if os.path.exists(LOGO):
        try:
            img = tk.PhotoImage(file=LOGO)
            fator = max(1, img.width() // 200)
            if fator > 1:
                img = img.subsample(fator, fator)
            l = tk.Label(topo, image=img, bg="white")
            l.image = img
            l.pack()
        except Exception:
            pass
    tk.Label(topo, text="AGENDA TOPADA", font=("Segoe UI", 16, "bold"),
             bg="white", fg="#4F81BD").pack(pady=(2, 0))

    # BOTAO PRINCIPAL
    btn_rodar = tk.Button(
        janela, text="▶  GERAR AGENDA COMPLETA", command=gerar_agenda_completa,
        height=2, width=42, bg="#2E7D32", fg="white",
        font=("Segoe UI", 12, "bold"), relief="flat", cursor="hand2")
    btn_rodar.pack(pady=(8, 6))

    # setup (config unica) + fallback
    linha = tk.Frame(janela, bg="white")
    linha.pack(pady=2)
    def mini(txt, cmd):
        tk.Button(linha, text=txt, command=cmd, font=("Segoe UI", 8),
                  relief="flat", fg="#4F81BD", bg="#EAF0F8",
                  cursor="hand2", padx=6, pady=3).pack(side="left", padx=3)
    mini("Calibrar direcionamento", calibrar_direcionamento)
    mini("Calibrar relatorio", calibrar_relatorio)
    mini("So gerar agenda", so_gerar_agenda)

    # log
    tk.Label(janela, text="Andamento:", bg="white", fg="#555",
             font=("Segoe UI", 8)).pack(anchor="w", padx=12, pady=(8, 0))
    caixa_log = tk.Text(janela, height=14, width=70, state="disabled",
                        bg="#1E1E1E", fg="#DDDDDD", font=("Consolas", 8),
                        wrap="word", relief="flat")
    caixa_log.pack(padx=12, pady=(2, 6), fill="both", expand=True)

    # rodape com pastas
    frm = tk.Frame(janela, bg="white")
    frm.pack(side="bottom", fill="x", pady=6)
    lbl_pdfs = tk.Label(frm, text=f"PDFs: {get_pasta_pdfs()}", bg="white",
                        fg="#777", font=("Segoe UI", 7), wraplength=540)
    lbl_pdfs.pack()
    lbl_saida = tk.Label(frm, text=f"Saida: {get_pasta_saida()}", bg="white",
                         fg="#777", font=("Segoe UI", 7), wraplength=540)
    lbl_saida.pack()
    lnk = tk.Frame(frm, bg="white")
    lnk.pack()
    tk.Button(lnk, text="alterar pasta PDFs", command=escolher_pasta_pdfs,
              font=("Segoe UI", 7), relief="flat", fg="#4F81BD",
              bg="white", cursor="hand2").pack(side="left", padx=6)
    tk.Button(lnk, text="alterar pasta saida", command=escolher_pasta_saida,
              font=("Segoe UI", 7), relief="flat", fg="#4F81BD",
              bg="white", cursor="hand2").pack(side="left", padx=6)

    if dnav is None:
        log("ATENCAO: nao consegui carregar o direcionar_os.py:")
        log("  " + _erro_dnav)

    janela.after(200, _drenar_fila)
    janela.mainloop()


if __name__ == "__main__":
    construir_janela()
