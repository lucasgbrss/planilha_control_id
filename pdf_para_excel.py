import pdfplumber
from pathlib import Path
from openpyxl import Workbook
import re
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

def selecionar_pdf():
    """Abre dialog para selecionar arquivo PDF"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    pdf_path = filedialog.askopenfilename(
        title="Selecione o arquivo PDF",
        filetypes=[("Arquivos PDF", "*.pdf")],
        initialdir=Path.home() / "Downloads"
    )
    root.destroy()
    return pdf_path

def obter_caminho_saida(nome_pdf):
    """Retorna caminho na pasta Documentos com nome baseado no PDF"""
    documentos = Path.home() / "OneDrive" / "Documentos"
    if not documentos.exists():
        documentos = Path.home() / "Documentos"

    # Nome do Excel baseado no nome do PDF
    nome_base = Path(nome_pdf).stem
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_nome = f"{nome_base}_{timestamp}.xlsx"

    return documentos / excel_nome

# Selecionar PDF
pdf_path = selecionar_pdf()
if not pdf_path:
    print("Nenhum arquivo PDF selecionado.")
    exit()

print(f"PDF selecionado: {pdf_path}")

# Definir caminho de saída
excel_path = obter_caminho_saida(pdf_path)
print(f"Planilha será salva em: {excel_path}")

# Extrair tabelas do PDF
tabelas = []
with pdfplumber.open(pdf_path) as pdf:
    for pagina in pdf.pages:
        tables = pagina.extract_tables()
        for table in tables:
            if table:
                tabelas.extend(table)

# Extrair texto completo para informações
with pdfplumber.open(pdf_path) as pdf:
    texto = ""
    for pagina in pdf.pages:
        texto += pagina.extract_text() or ""

# Parse das informações usando regex
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
match = re.search(r'ENDERE[Ç�]O:\s*(.+?)(?:\n|$)', texto, re.IGNORECASE)
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
match = re.search(r'ADMISS[Ã�]O:\s*(\d{2}/\d{2}/\d{4})', texto, re.IGNORECASE)
if match:
    info["admissao"] = match.group(1).strip()

# CPF
match = re.search(r'CPF:\s*(\d+)', texto, re.IGNORECASE)
if match:
    info["cpf"] = match.group(1).strip()

# MATRÍCULA
match = re.search(r'MATR[Í�]CULA:\s*(\d+)', texto, re.IGNORECASE)
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

# Extrair dados de ponto por dia
dados_ponto = {}
dia_atual = None

for row in tabelas:
    # Detectar nova linha de dia (formato: DD/MM/YY - DIA)
    if row[0] and any(mes in str(row[0]).upper() for mes in ["SEG", "TER", "QUA", "QUI", "SEX", "SAB", "DOM"]):
        dia_atual = str(row[0]).strip()
        dados_ponto[dia_atual] = {
            "marcacoes": "",
            "ent1": "", "sai1": "", "ent2": "", "sai2": "", "ent3": "", "sai3": "",
            "duracao": "", "ch": ""
        }

    if dia_atual:
        # Marcações registradas
        if row[1]:
            dados_ponto[dia_atual]["marcacoes"] = str(row[1]).strip()

        # Jornada realizada
        if row[2]:
            dados_ponto[dia_atual]["ent1"] = str(row[2]).strip()
        if row[3]:
            dados_ponto[dia_atual]["sai1"] = str(row[3]).strip()
        if row[4]:
            dados_ponto[dia_atual]["ent2"] = str(row[4]).strip()
        if row[5]:
            dados_ponto[dia_atual]["sai2"] = str(row[5]).strip()
        if row[6]:
            dados_ponto[dia_atual]["ent3"] = str(row[6]).strip()
        if row[7]:
            dados_ponto[dia_atual]["sai3"] = str(row[7]).strip()

        # Duração e CH
        if row[8]:
            dados_ponto[dia_atual]["duracao"] = str(row[8]).strip()
        if row[9]:
            dados_ponto[dia_atual]["ch"] = str(row[9]).strip()

# Criar Excel formatado
wb = Workbook()
ws = wb.active
ws.title = "Ponto"

linha_atual = 1

# Informações da empresa
ws.cell(row=linha_atual, column=1, value="EMPRESA:")
ws.cell(row=linha_atual, column=2, value=info.get("empresa", ""))
ws.cell(row=linha_atual, column=5, value="CNPJ:")
ws.cell(row=linha_atual, column=6, value=info.get("cnpj", ""))
linha_atual += 1

ws.cell(row=linha_atual, column=1, value="ENDEREÇO:")
ws.cell(row=linha_atual, column=2, value=info.get("endereco", ""))
linha_atual += 2  # Linha em branco

# Informações do funcionário
ws.cell(row=linha_atual, column=1, value="NOME:")
ws.cell(row=linha_atual, column=2, value=info.get("nome", ""))
linha_atual += 1

ws.cell(row=linha_atual, column=1, value="PIS/PASEP:")
ws.cell(row=linha_atual, column=2, value=info.get("pis", ""))
ws.cell(row=linha_atual, column=5, value="ADMISSÃO:")
ws.cell(row=linha_atual, column=6, value=info.get("admissao", ""))
linha_atual += 1

ws.cell(row=linha_atual, column=1, value="CPF:")
ws.cell(row=linha_atual, column=2, value=info.get("cpf", ""))
ws.cell(row=linha_atual, column=5, value="MATRÍCULA:")
ws.cell(row=linha_atual, column=6, value=info.get("matricula", ""))
linha_atual += 1

ws.cell(row=linha_atual, column=1, value="CENTRO DE CUSTO:")
ws.cell(row=linha_atual, column=2, value=info.get("centro_custo", ""))
linha_atual += 1

ws.cell(row=linha_atual, column=1, value="DEPARTAMENTO:")
ws.cell(row=linha_atual, column=2, value=info.get("departamento", ""))
ws.cell(row=linha_atual, column=5, value="CARGO:")
ws.cell(row=linha_atual, column=6, value=info.get("cargo", ""))
linha_atual += 2  # Linha em branco de separação

# Cabeçalhos da tabela de ponto - todos na mesma linha
ws.cell(row=linha_atual, column=1, value="DIA")
ws.cell(row=linha_atual, column=2, value="MARCAÇÕES REGISTRADAS\nNO PONTO ELETRÔNICO")
ws.cell(row=linha_atual, column=3, value="JORNADA REALIZADA")
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
for dia, dados in dados_ponto.items():
    ws.cell(row=linha_atual, column=1, value=dia)
    ws.cell(row=linha_atual, column=2, value=dados["marcacoes"])
    ws.cell(row=linha_atual, column=4, value=dados["ent1"])
    ws.cell(row=linha_atual, column=5, value=dados["sai1"])
    ws.cell(row=linha_atual, column=6, value=dados["ent2"])
    ws.cell(row=linha_atual, column=7, value=dados["sai2"])
    ws.cell(row=linha_atual, column=8, value=dados["ent3"])
    ws.cell(row=linha_atual, column=9, value=dados["sai3"])
    ws.cell(row=linha_atual, column=10, value=dados["duracao"])
    ws.cell(row=linha_atual, column=11, value=dados["ch"])
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
print(f"Total de dias registrados: {len(dados_ponto)}")
print("\nInformações extraídas:")
for k, v in info.items():
    print(f"  {k}: {v}")
