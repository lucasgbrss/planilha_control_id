import pdfplumber
from pathlib import Path
from openpyxl import Workbook
import re
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

def selecionar_pdfs():
    """Abre dialog para selecionar múltiplos arquivos PDF"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    pdf_paths = filedialog.askopenfilenames(
        title="Selecione os arquivos PDF",
        filetypes=[("Arquivos PDF", "*.pdf")],
        initialdir=Path.home() / "Downloads"
    )
    root.destroy()
    return list(pdf_paths)

def selecionar_local_salvar(qtd_pdfs):
    """Abre dialog para selecionar onde salvar o arquivo Excel"""
    # Nome sugerido baseado na quantidade de PDFs
    nome_base = f"Ponto_{qtd_pdfs}PDFs"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_sugerido = f"{nome_base}_{timestamp}.xlsx"

    # Pasta inicial sugerida (Documentos)
    documentos = Path.home() / "OneDrive" / "Documentos"
    if not documentos.exists():
        documentos = Path.home() / "Documentos"

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    # Abre dialog para salvar arquivo
    excel_path = filedialog.asksaveasfilename(
        title="Salvar planilha Excel",
        defaultextension=".xlsx",
        initialfile=nome_sugerido,
        initialdir=documentos,
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )
    root.destroy()
    return excel_path

# Selecionar múltiplos PDFs
pdf_paths = selecionar_pdfs()
if not pdf_paths:
    print("Nenhum arquivo PDF selecionado.")
    exit()

print(f"{len(pdf_paths)} arquivo(s) PDF selecionado(s):")
for path in pdf_paths:
    print(f"  - {path}")

# Selecionar local para salvar o Excel
excel_path = selecionar_local_salvar(len(pdf_paths))
if not excel_path:
    print("Operação cancelada.")
    exit()

print(f"Planilha será salva em: {excel_path}")

# Extrair texto completo de todos os PDFs para informações do funcionário
todos_textos = []
for pdf_path in pdf_paths:
    print(f"Processando: {pdf_path}")
    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            todos_textos.append(pagina.extract_text() or "")


def get_item(lista, indice):
    """Acessa índice da lista com segurança, retorna None se fora de alcance"""
    return lista[indice] if indice < len(lista) else None

def extrair_info_funcionario(texto):
    """Extrai informações do funcionário a partir do texto de um PDF"""
    info = {}

    # EMPRESA e CNPJ na mesma linha
    match = re.search(r'EMPRESA:\s*(.+?)\s*CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', texto, re.IGNORECASE)
    if match:
        info["empresa"] = match.group(1).strip()
        info["cnpj"] = match.group(2).strip()
    else:
        match = re.search(r'EMPRESA:\s*([^\n]+)', texto, re.IGNORECASE)
        if match:
            info["empresa"] = match.group(1).strip()
        match = re.search(r'CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', texto, re.IGNORECASE)
        if match:
            info["cnpj"] = match.group(1).strip()

    # ENDEREÇO
    match = re.search(r'ENDERE[Ç]O:\s*(.+?)(?:\n|$)', texto, re.IGNORECASE)
    if match:
        info["endereco"] = match.group(1).strip()

    # NOME (pegar apenas o nome, antes de PIS/PASEP)
    match = re.search(r'NOME:\s*(.+?)\s+PIS/PASEP:', texto, re.IGNORECASE)
    if match:
        info["nome"] = match.group(1).strip()

    # PIS/PASEP
    match = re.search(r'PIS/PASEP:\s*(\d+)', texto, re.IGNORECASE)
    if match:
        info["pis"] = match.group(1).strip()

    # ADMISSÃO
    match = re.search(r'ADMISS[Ã]O:\s*(\d{2}/\d{2}/\d{4})', texto, re.IGNORECASE)
    if match:
        info["admissao"] = match.group(1).strip()

    # CPF
    match = re.search(r'CPF:\s*(\d+)', texto, re.IGNORECASE)
    if match:
        info["cpf"] = match.group(1).strip()

    # MATRÍCULA
    match = re.search(r'MATR[Í]CULA:\s*(\d+)', texto, re.IGNORECASE)
    if match:
        info["matricula"] = match.group(1).strip()

    # CENTRO DE CUSTO
    match = re.search(r'CENTRO DE CUSTO:\s*(\S+)', texto, re.IGNORECASE)
    if match:
        info["centro_custo"] = match.group(1).strip()

    # DEPARTAMENTO
    match = re.search(r'DEPARTAMENTO:\s*(\S+)', texto, re.IGNORECASE)
    if match:
        info["departamento"] = match.group(1).strip()

    # CARGO
    match = re.search(r'CARGO:\s*(.+?)(?:\n|$)', texto, re.IGNORECASE)
    if match:
        info["cargo"] = match.group(1).strip()

    return info

def extrair_dados_ponto(tabelas):
    """Extrai dados de ponto de uma lista de tabelas, retornando uma lista de dias"""
    dias = []
    dia_atual = None

    for row in tabelas:
        # Detectar nova linha de dia (formato: DD/MM/YY - DIA)
        if get_item(row, 0) and any(mes in str(get_item(row, 0)).upper() for mes in ["SEG", "TER", "QUA", "QUI", "SEX", "SAB", "DOM"]):
            dia_atual = str(get_item(row, 0)).strip()
            dados_dia = {
                "dia": dia_atual,
                "marcacoes": "",
                "ent1": "", "sai1": "", "ent2": "", "sai2": "", "ent3": "", "sai3": "",
                "duracao": "", "ch": ""
            }
            dias.append(dados_dia)

        if dia_atual:
            # Marcações registradas
            if get_item(row, 1):
                dias[-1]["marcacoes"] = str(get_item(row, 1)).strip()

            # Jornada realizada
            if get_item(row, 2):
                dias[-1]["ent1"] = str(get_item(row, 2)).strip()
            if get_item(row, 3):
                dias[-1]["sai1"] = str(get_item(row, 3)).strip()
            if get_item(row, 4):
                dias[-1]["ent2"] = str(get_item(row, 4)).strip()
            if get_item(row, 5):
                dias[-1]["sai2"] = str(get_item(row, 5)).strip()
            if get_item(row, 6):
                dias[-1]["ent3"] = str(get_item(row, 6)).strip()
            if get_item(row, 7):
                dias[-1]["sai3"] = str(get_item(row, 7)).strip()

            # Duração e CH
            if get_item(row, 8):
                dias[-1]["duracao"] = str(get_item(row, 8)).strip()
            if get_item(row, 9):
                dias[-1]["ch"] = str(get_item(row, 9)).strip()

    return dias

# Parse das informações do primeiro PDF para dados gerais (usado apenas como fallback)
info_geral = extrair_info_funcionario(todos_textos[0]) if todos_textos else {}

# Extrair dados de ponto de todos os PDFs
# Processar cada PDF separadamente para manter dados únicos
todos_dias = []
tabela_atual = 0

for idx, pdf_path in enumerate(pdf_paths):
    # Extrair texto e informações do funcionário deste PDF
    texto_do_pdf = ""
    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            texto_do_pdf += pagina.extract_text() or ""

    info_funcionario = extrair_info_funcionario(texto_do_pdf)

    # Extrair tabelas apenas deste PDF
    tabelas_do_pdf = []
    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            tables = pagina.extract_tables()
            for table in tables:
                if table:
                    tabelas_do_pdf.extend(table)

    # Extrair dias deste PDF e associar às informações do funcionário
    dias_do_pdf = extrair_dados_ponto(tabelas_do_pdf)

    # Adicionar informações do funcionário a cada dia
    for dia in dias_do_pdf:
        dia["nome"] = info_funcionario.get("nome", "")
        dia["empresa"] = info_funcionario.get("empresa", "")
        dia["cpf"] = info_funcionario.get("cpf", "")

    todos_dias.extend(dias_do_pdf)
    print(f"  -> {len(dias_do_pdf)} dias extraídos de {Path(pdf_path).name}")

dados_ponto = todos_dias

# Criar Excel formatado
wb = Workbook()
ws = wb.active
ws.title = "Ponto"

linha_atual = 1

# Cabeçalhos da tabela de ponto - todos na mesma linha
ws.cell(row=linha_atual, column=1, value="FUNCIONÁRIO")
ws.cell(row=linha_atual, column=2, value="DIA")
ws.cell(row=linha_atual, column=3, value="MARCAÇÕES REGISTRADAS\nNO PONTO ELETRÔNICO")
ws.cell(row=linha_atual, column=4, value="ENT. 1")
ws.cell(row=linha_atual, column=5, value="SAÍ. 1")
ws.cell(row=linha_atual, column=6, value="ENT. 2")
ws.cell(row=linha_atual, column=7, value="SAÍ. 2")
ws.cell(row=linha_atual, column=8, value="ENT. 3")
ws.cell(row=linha_atual, column=9, value="SAÍ. 3")
ws.cell(row=linha_atual, column=10, value="DURAÇÃO")
ws.cell(row=linha_atual, column=11, value="CH")
linha_atual += 1


# Dados de cada dia
for dia_dados in dados_ponto:
    ws.cell(row=linha_atual, column=1, value=dia_dados.get("nome", ""))
    ws.cell(row=linha_atual, column=2, value=dia_dados["dia"])
    ws.cell(row=linha_atual, column=3, value=dia_dados["marcacoes"])
    ws.cell(row=linha_atual, column=4, value=dia_dados["ent1"])
    ws.cell(row=linha_atual, column=5, value=dia_dados["sai1"])
    ws.cell(row=linha_atual, column=6, value=dia_dados["ent2"])
    ws.cell(row=linha_atual, column=7, value=dia_dados["sai2"])
    ws.cell(row=linha_atual, column=8, value=dia_dados["ent3"])
    ws.cell(row=linha_atual, column=9, value=dia_dados["sai3"])
    ws.cell(row=linha_atual, column=10, value=dia_dados["duracao"])
    ws.cell(row=linha_atual, column=11, value=dia_dados["ch"])
    linha_atual += 1

# Ajustar largura das colunas
for col in ws.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        if cell.value:
            max_length = max(max_length, len(str(cell.value)))
    ws.column_dimensions[column].width = min(max_length + 2, 30)

wb.save(excel_path)
print(f"\nPlanilha criada com sucesso: {excel_path}")
print(f"Total de PDFs processados: {len(pdf_paths)}")
print(f"Total de dias registrados: {len(dados_ponto)}")
print("\nInformações extraídas (do primeiro PDF):")
for k, v in info_geral.items():
    print(f"  {k}: {v}")
