import io
import re
import zipfile

import pdfplumber
import pypdf

from control_id_reader.utils import nome_para_arquivo


def detect_employee_page_groups_from_pages(pages):
    """Detecta grupos de paginas por funcionario a partir de objetos pagina."""
    grupos = []
    grupo_atual = None

    for index, pagina in enumerate(pages):
        texto = pagina.extract_text() or ""
        match = re.search(r"NOME:\s*(.+?)\s+PIS/PASEP:", texto, re.IGNORECASE)
        if match:
            nome = match.group(1).strip()
            grupo_atual = {"nome": nome, "paginas": [index]}
            grupos.append(grupo_atual)
        elif grupo_atual:
            grupo_atual["paginas"].append(index)

    return grupos


def detect_employee_page_groups(pdf_path):
    """Detecta grupos de paginas por funcionario em um PDF."""
    with pdfplumber.open(pdf_path) as pdf:
        return detect_employee_page_groups_from_pages(pdf.pages)


def unique_pdf_filename(nome, nomes_usados):
    """Gera nome de arquivo PDF unico para um funcionario."""
    nome_base = nome_para_arquivo(nome)

    if nome_base in nomes_usados:
        nomes_usados[nome_base] += 1
        return f"{nome_base}_{nomes_usados[nome_base]}.pdf"

    nomes_usados[nome_base] = 1
    return f"{nome_base}.pdf"


def split_pdf_by_employee(pdf_path, zip_path, progress_callback=None):
    """Separa um PDF multi-funcionario em um ZIP com um PDF por funcionario."""
    grupos = detect_employee_page_groups(pdf_path)
    if not grupos:
        return {"status": "no_employees", "groups": [], "files": []}
    if len(grupos) == 1:
        return {"status": "single_employee", "groups": grupos, "files": []}

    arquivos = write_employee_zip(pdf_path, zip_path, grupos, progress_callback=progress_callback)
    return {"status": "ok", "groups": grupos, "files": arquivos}


def write_employee_zip(pdf_path, zip_path, grupos, progress_callback=None):
    """Grava um ZIP com um PDF por grupo de paginas detectado."""
    reader = pypdf.PdfReader(pdf_path)
    arquivos = []

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        nomes_usados = {}
        total = len(grupos)
        for index, grupo in enumerate(grupos, 1):
            writer = pypdf.PdfWriter()
            for idx_pagina in grupo["paginas"]:
                writer.add_page(reader.pages[idx_pagina])

            nome_arquivo = unique_pdf_filename(grupo["nome"], nomes_usados)
            buffer = io.BytesIO()
            writer.write(buffer)
            zf.writestr(nome_arquivo, buffer.getvalue())

            arquivo = {
                "nome": grupo["nome"],
                "arquivo": nome_arquivo,
                "paginas": len(grupo["paginas"]),
            }
            arquivos.append(arquivo)
            if progress_callback:
                progress_callback(arquivo, index, total)

    return arquivos
