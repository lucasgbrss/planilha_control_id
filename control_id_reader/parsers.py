import email as email_lib
import re

from control_id_reader.utils import normalizar_cpf, parse_data_ponto


def extract_employee_info_from_text(texto):
    """Extrai informacoes do funcionario a partir do texto de um PDF."""
    info = {}

    if not texto or not texto.strip():
        return info, "texto do arquivo está vazio"

    match = re.search(
        r"EMPRESA:\s*(.+?)\s*CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})",
        texto,
        re.IGNORECASE,
    )
    if match:
        info["empresa"] = match.group(1).strip()
        info["cnpj"] = match.group(2).strip()
    else:
        match = re.search(r"EMPRESA:\s*([^\n]+)", texto, re.IGNORECASE)
        if match:
            info["empresa"] = match.group(1).strip()
        match = re.search(r"CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto, re.IGNORECASE)
        if match:
            info["cnpj"] = match.group(1).strip()

    match = re.search(r"ENDERE[Ç]O:\s*(.+?)(?:\n|$)", texto, re.IGNORECASE)
    if match:
        info["endereco"] = match.group(1).strip()

    match = re.search(r"NOME:\s*(.+?)\s+PIS/PASEP:", texto, re.IGNORECASE)
    if match:
        info["nome"] = match.group(1).strip()

    match = re.search(r"PIS/PASEP:\s*(\d+)", texto, re.IGNORECASE)
    if match:
        info["pis"] = match.group(1).strip()

    match = re.search(r"ADMISS[Ã]O:\s*(\d{2}/\d{2}/\d{4})", texto, re.IGNORECASE)
    if match:
        info["admissao"] = match.group(1).strip()

    match = re.search(r"CPF:\s*([\d\.\-]+)", texto, re.IGNORECASE)
    if match:
        info["cpf"] = normalizar_cpf(match.group(1))

    match = re.search(r"MATR[Í]CULA:\s*(\d+)", texto, re.IGNORECASE)
    if match:
        info["matricula"] = match.group(1).strip()

    match = re.search(
        r"CENTRO DE CUSTO:\s*(.+?)(?=\s+(?:CPF|MATR[Í]CULA|DEPARTAMENTO|CARGO|ADMISS[Ã]O|$))",
        texto,
        re.IGNORECASE,
    )
    if match:
        info["centro_custo"] = match.group(1).strip()

    match = re.search(
        r"DEPARTAMENTO:\s*(.+?)(?=\s+(?:CARGO|CENTRO DE CUSTO|CPF|MATR[Í]CULA|ADMISS[Ã]O|$))",
        texto,
        re.IGNORECASE,
    )
    if match:
        info["departamento"] = match.group(1).strip()

    match = re.search(r"CARGO:\s*(.+?)(?:\n|$)", texto, re.IGNORECASE)
    if match:
        info["cargo"] = match.group(1).strip()

    if not info.get("nome"):
        tem_empresa = "empresa" in info
        tem_cpf = "cpf" in info
        if not tem_empresa and not tem_cpf:
            motivo = "nenhum campo reconhecido — formato do arquivo pode ser diferente do esperado"
        elif tem_empresa and not tem_cpf:
            motivo = "empresa identificada mas campo NOME não encontrado no padrão 'NOME: ... PIS/PASEP:'"
        else:
            motivo = "campo NOME não encontrado no padrão esperado"
        return info, motivo

    return info, None


def extract_punch_rows_from_tables(tabelas):
    """Extrai registros de ponto e horarios contratuais de tabelas pdfplumber."""
    dias = []
    horarios_contratuais = []
    capturando_horarios = False

    time_re = re.compile(r"^\d{1,2}:\d{2}$")
    ch_re = re.compile(r"^\d{5}$")
    day_re = re.compile(
        r"^\d{2}/\d{2}/\d{2,4}\s*-\s*(SEG|TER|QUA|QUI|SEX|SAB|DOM)",
        re.IGNORECASE,
    )

    def is_time(valor):
        return bool(valor and time_re.match(str(valor).strip()))

    def is_ch(valor):
        return bool(valor and ch_re.match(str(valor).strip()))

    def is_schedule_header(row_text, row_values):
        first_cell = str(row_values[0] or "").lower() if row_values else ""
        ent_count = sum(1 for value in row_values if str(value or "").strip().lower() == "ent")
        sai_count = sum(1 for value in row_values if str(value or "").strip().lower() == "sai")
        known_marker = any(marker in row_text for marker in (
            "horários contratuais", "horarios contratuais",
            "código do horário", "codigo do horario",
        ))
        garbled_marker = "ch" in first_cell and ("hor" in first_cell or "rio" in first_cell)
        header_shape = "ch" in first_cell and ent_count >= 1 and sai_count >= 1
        return known_marker or garbled_marker or header_shape

    for row in tabelas:
        if not row:
            continue

        def cel(indice):
            return str(row[indice]).strip() if indice < len(row) and row[indice] else ""

        linha_texto = " ".join(str(c) for c in row if c).lower()

        if is_schedule_header(linha_texto, row):
            capturando_horarios = True
            continue

        if capturando_horarios:
            codigo = cel(0)
            if codigo:
                horario = extract_schedule_row(row, is_time)
                if horario:
                    horarios_contratuais.append(horario)
            continue

        if not (row[0] and day_re.match(str(row[0]).strip())):
            continue

        dia = {
            "dia": str(row[0]).strip(),
            "marcacoes": cel(1),
            "ent1": cel(2),
            "sai1": cel(3),
            "ent2": "",
            "sai2": "",
            "ent3": "",
            "sai3": "",
            "duracao": "",
            "ch": "",
        }

        cel4_raw = cel(4)
        cel5_raw = cel(5)

        if cel5_raw.startswith(":") and cel4_raw:
            parts4 = cel4_raw.split()
            if parts4:
                hours_part = parts4[-1]
                cel4_raw = " ".join(parts4[:-1])
                cel5_raw = f"{hours_part}{cel5_raw}"

        time_parts4 = [p for p in cel4_raw.split() if is_time(p)] if cel4_raw else []

        if len(time_parts4) >= 2:
            dia["ent2"] = time_parts4[0]
            dia["sai2"] = time_parts4[1]
            merged = True
        else:
            dia["ent2"] = cel4_raw if is_time(cel4_raw) else ""
            dia["sai2"] = cel5_raw if is_time(cel5_raw) else ""
            merged = False

        ch_positions = [9, 8, 7, 6, 5] if not merged else [7, 6, 5]

        found = False
        for ch_i in ch_positions:
            valor = cel(ch_i)
            if is_ch(valor):
                dia["ch"] = valor
                prev = cel(ch_i - 1)
                if is_time(prev):
                    dia["duracao"] = prev
                if ch_i == 9:
                    dia["ent3"] = cel(6)
                    dia["sai3"] = cel(7)
                found = True
                break

        if not found:
            is_format_b = is_ch(cel(9))
            search_order = [8, 7, 6] if is_format_b else [6, 7, 8]

            for dur_i in search_order:
                if is_time(cel(dur_i)):
                    dia["duracao"] = cel(dur_i)
                    if dur_i == 8 and is_format_b:
                        dia["ent3"] = cel(6)
                        dia["sai3"] = cel(7)
                    break

        dias.append(dia)

    return dias, horarios_contratuais


def extract_schedule_row(row, is_time_func=None):
    """Extrai um horario contratual a partir de uma linha de tabela CH."""
    if not row:
        return None

    codigo = str(row[0] or "").strip()
    if not codigo:
        return None

    is_time_func = is_time_func or (lambda value: bool(re.match(r"^\d{1,2}:\d{2}$", str(value or "").strip())))
    times = [str(value).strip() for value in row[1:] if is_time_func(value)]
    if not times:
        return None

    horario = {"codigo": codigo}
    par = 1
    for idx in range(0, len(times), 2):
        ent = times[idx]
        sai = times[idx + 1] if idx + 1 < len(times) else ""
        horario[f"ent{par}"] = ent
        horario[f"sai{par}"] = sai
        par += 1

    return horario


def extract_html_from_mhtml(caminho):
    """Abre um arquivo .mhtml e retorna o HTML interno como string."""
    with open(caminho, "rb") as f:
        msg = email_lib.message_from_bytes(f.read())
    for part in msg.walk():
        if "html" in part.get_content_type():
            payload = part.get_payload(decode=True)
            charset = part.get_content_charset() or "utf-8"
            return payload.decode(charset, errors="replace")
    return None


def extract_employee_info_from_mhtml(soup):
    """Extrai dados do funcionario a partir do HTML parseado."""
    info = {}

    campo_map_lower = {
        "empresa": "empresa",
        "cnpj": "cnpj",
        "cei": "cei",
        "endereço": "endereco",
        "nome": "nome",
        "pis/pasep": "pis",
        "admissão": "admissao",
        "centro de custo": "centro_custo",
        "cpf": "cpf",
        "matrícula": "matricula",
        "departamento": "departamento",
        "cargo": "cargo",
    }

    for td in soup.find_all(["td", "th"]):
        texto = td.get_text(separator=" ", strip=True)
        if ":" not in texto:
            continue
        chave, _, valor = texto.partition(":")
        chave_norm = chave.strip().lower()
        valor = valor.strip()
        if chave_norm in campo_map_lower and valor:
            campo_destino = campo_map_lower[chave_norm]
            info[campo_destino] = normalizar_cpf(valor) if campo_destino == "cpf" else valor

    if not info.get("nome"):
        tem_empresa = "empresa" in info
        if not tem_empresa:
            motivo = "nenhum campo reconhecido — estrutura do MHTML diferente do esperado"
        else:
            motivo = "empresa identificada mas campo NOME não encontrado"
        return info, motivo

    return info, None


def extract_punch_rows_from_mhtml(soup):
    """Extrai registros de ponto da tabela HTML."""
    dias = []
    tables = soup.find_all("table")

    tabela_ponto = None
    for table in tables:
        primeira = table.find("tr")
        if primeira and "DIA" in primeira.get_text():
            tabela_ponto = table
            break

    if not tabela_ponto:
        return dias

    data_re = re.compile(r"(\d{2}/\d{2}/\d{2,4})")

    for tr in tabela_ponto.find_all("tr"):
        cells = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]

        if len(cells) < 17:
            continue
        match = data_re.match(cells[0])
        if not match:
            continue

        data_raw = cells[0]
        data_part = match.group(1)
        data = parse_data_ponto(data_part)
        data_fmt = data.strftime("%d/%m/%Y") if data != data.min else data_part
        dia_semana = data_raw.split("-")[-1].strip() if "-" in data_raw else ""

        ent1 = cells[3].strip()
        sai1 = cells[4].strip()
        ent2 = cells[5].strip()
        sai2 = cells[6].strip()
        duracao = cells[10].strip() if len(cells) > 10 else ""
        ch = cells[12].strip() if len(cells) > 12 else ""

        marcacoes = f"{ent1} {sai1}".strip() if ent1 or sai1 else ""

        dias.append({
            "dia": f"{data_fmt} - {dia_semana}",
            "marcacoes": marcacoes,
            "ent1": ent1,
            "sai1": sai1,
            "ent2": ent2,
            "sai2": sai2,
            "ent3": "",
            "sai3": "",
            "duracao": duracao,
            "ch": ch,
        })

    return dias


def extract_pdf_employees(pdf, logger=None):
    """Processa um PDF com um ou mais funcionarios, pagina a pagina."""
    resultados = []
    texto_atual = ""
    tabelas_atual = []

    def finalizar_secao(texto, tabelas):
        if not texto.strip():
            return None
        info, motivo = extract_employee_info_from_text(texto)
        if not info.get("nome"):
            if logger:
                logger(f"    ⚠️ Seção sem funcionário identificado: {motivo}", "error")
            return None
        dias, horarios = extract_punch_rows_from_tables(tabelas)
        if horarios:
            info["horarios_contratuais"] = horarios
        return info, dias

    for pagina in pdf.pages:
        texto_pagina = pagina.extract_text() or ""
        tabelas_pagina = []
        for tabela in (pagina.extract_tables() or []):
            if tabela:
                tabelas_pagina.extend(tabela)

        e_novo_funcionario = bool(
            re.search(r"NOME:\s*.+?\s+PIS/PASEP:", texto_pagina, re.IGNORECASE)
        )

        if e_novo_funcionario and texto_atual:
            resultado = finalizar_secao(texto_atual, tabelas_atual)
            if resultado:
                resultados.append(resultado)
            texto_atual = ""
            tabelas_atual = []

        texto_atual += texto_pagina
        tabelas_atual += tabelas_pagina

    if texto_atual:
        resultado = finalizar_secao(texto_atual, tabelas_atual)
        if resultado:
            resultados.append(resultado)

    return resultados
