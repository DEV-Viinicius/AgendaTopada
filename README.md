# Agenda Topada

Automação da rotina de agenda de Ordens de Serviço (O.S.) da Solivetti sobre o
sistema **ILUX**. Reúne, num único programa, as três fases do processo:

1. **Direcionar O.S.** – lê os PDFs de cada técnico e, para cada O.S., altera o
   técnico responsável no ILUX (via automação de mouse/teclado).
2. **Exportar relatório** – exporta do ILUX o relatório de O.S. em Excel.
3. **Gerar agenda** – monta a `AGENDA_FILTRADA.xlsx` já ordenada, com cada
   técnico na mesma ordem em que as O.S. aparecem no PDF dele.

## Como funciona a leitura do número da O.S.

O número é lido pela **posição do rótulo** na folha (camada de texto do PDF ou,
quando o PDF é um scan, via OCR com EasyOCR). São cobertos **3 modelos** de
folha, nesta ordem de preferência:

| Modelo      | Rótulo      | Onde fica o número           |
|-------------|-------------|------------------------------|
| normal      | `Número :`  | mesma linha, à direita       |
| instalação  | `Número :`  | mesma linha, à direita       |
| leitura     | `OS Nº :`   | mesma linha, à direita       |

Se nenhum rótulo aparecer na página, ela é pulada. O número da O.S. é validado
como tendo **5 dígitos ou mais e não começar com zero** — isso descarta outros
números longos da folha (Inscrição Estadual, CNPJ, CEP, telefone).

## Estrutura

| Arquivo                  | Função                                              |
|--------------------------|-----------------------------------------------------|
| `agenda_topada.py`       | App principal (menu de 1 clique que roda as 3 fases)|
| `direcionar_os.py`       | Fase 1 – direciona as O.S. no ILUX; extrai o número |
| `exportar_relatorio.py`  | Fase 2 – exporta o relatório do ILUX                |
| `gerar_agenda_topada.py` | Fase 3 – gera a agenda ordenada pelos PDFs          |
| `config_*.json`          | Coordenadas calibradas de cada fase                 |
| `PDFS/`                  | 1 PDF por técnico (nome do arquivo = nome do técnico)|

> Os PDFs (`PDFS/` e exemplos) **não estão no repositório** por conterem dados
> de clientes. Crie a pasta `PDFS` localmente e coloque um PDF por técnico
> (ex.: `WILSON.pdf`).

## Requisitos

- Python 3.x
- Bibliotecas: `pymupdf` (fitz), `pandas`, `xlsxwriter`, `xlrd`, `openpyxl`,
  `pyautogui`, `pyperclip`, `easyocr`

```
pip install pymupdf pandas xlsxwriter xlrd openpyxl pyautogui pyperclip easyocr
```

## Como usar

Programa completo (menu):

```
python agenda_topada.py
```

Fase de direcionamento isolada, com opções úteis para teste:

```
python direcionar_os.py --simular          # só lê os PDFs e mostra as O.S. (não mexe no ILUX)
python direcionar_os.py --calibrar         # (re)define as coordenadas do ILUX
python direcionar_os.py --limite 1         # executa de verdade, só a 1ª O.S. (teste seguro)
python direcionar_os.py                     # executa tudo
```

### Segurança da automação

- **FAILSAFE**: jogue o mouse para o **canto superior esquerdo** da tela para
  abortar a qualquer momento.
- Toda execução real pede confirmação antes de começar a mexer no ILUX.
- É gerado um log em CSV a cada execução (`log_direcionamento_*.csv`).

## Observações

- A pasta padrão dos PDFs é a subpasta `PDFS` dentro do projeto (portátil).
- O nome do PDF (sem `.pdf`) precisa bater com a coluna **Técnico** do Excel
  exportado (comparação tolerante a maiúsculas/acentos/espaços).
- Os PDFs reais são scans (imagem), então a leitura passa por OCR na CPU — é
  normal a CPU subir enquanto processa.
