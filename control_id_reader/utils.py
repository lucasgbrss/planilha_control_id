from collections import Counter
from datetime import datetime
import re
import unicodedata


MESES_ABREV = {
    1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"
}


def normalizar_cpf(valor):
    """Retorna apenas os digitos do CPF ou string vazia."""
    return re.sub(r"\D", "", valor or "")


def formatar_cpf(valor):
    """Formata um CPF normalizado quando possivel."""
    cpf = normalizar_cpf(valor)
    if len(cpf) != 11:
        return (valor or "").strip()
    return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"


def parse_data_ponto(valor):
    """Converte o trecho de data de um registro de ponto para datetime."""
    data_str = (valor or "").split(" ")[0]
    partes = data_str.split("/")
    if len(partes) != 3:
        return datetime.min

    fmt = "%d/%m/%y" if len(partes[2]) == 2 else "%d/%m/%Y"
    try:
        return datetime.strptime(data_str, fmt)
    except ValueError:
        return datetime.min


def mes_ano_predominante(dias_lista):
    """Detecta mes/ano predominante nas datas dos dias."""
    meses_anos = []
    for dia in dias_lista:
        data = parse_data_ponto(dia.get("dia", ""))
        if data != datetime.min:
            meses_anos.append((data.month, data.year))

    if not meses_anos:
        return None

    mes, ano = Counter(meses_anos).most_common(1)[0][0]
    return f"{MESES_ABREV[mes]}{ano}"


def somar_duracoes(duracoes):
    """Soma uma lista de strings 'HH:MM' e retorna o total em 'HH:MM'."""
    total_min = 0
    for duracao in duracoes:
        if not duracao or ":" not in duracao:
            continue
        try:
            horas, minutos = duracao.strip().split(":")
            total_min += int(horas) * 60 + int(minutos)
        except ValueError:
            continue
    return f"{total_min // 60:02d}:{total_min % 60:02d}"


def media_horas(total_horas, total_dias):
    """Calcula a media de horas por dia a partir de 'HH:MM'."""
    if not total_dias or ":" not in total_horas:
        return "00:00"
    try:
        horas, minutos = total_horas.split(":")
        total_min = int(horas) * 60 + int(minutos)
        media_min = total_min // total_dias
        return f"{media_min // 60:02d}:{media_min % 60:02d}"
    except ValueError:
        return "00:00"


def nome_para_arquivo(nome):
    """Sanitiza texto para uso em nome de arquivo."""
    nfkd = unicodedata.normalize("NFKD", nome or "")
    ascii_nome = nfkd.encode("ASCII", "ignore").decode()
    nome_limpo = re.sub(r"[^\w\s-]", "", ascii_nome).strip()
    nome_limpo = re.sub(r"[\s_]+", "_", nome_limpo)
    return nome_limpo or "Funcionario"
