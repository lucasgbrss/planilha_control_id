from collections import Counter, defaultdict
import re

from control_id_reader.excel_writer import dias_faltados, dias_trabalhados, dias_trabalho_totais
from control_id_reader.utils import formatar_cpf, parse_data_ponto


PUNCH_FIELDS = ("ent1", "sai1", "ent2", "sai2", "ent3", "sai3")


def analyze_inconsistencies(dados, todos_funcionarios):
    """Retorna uma lista de inconsistencias encontradas nos dados extraidos."""
    issues = []

    for _chave, info in todos_funcionarios.items():
        nome = info.get("nome", "")
        cpf = info.get("cpf", "")
        dias_funcionario = info.get("_dias", [])

        if not cpf:
            issues.append(_issue("Atenção", nome, cpf, "", "CPF", "Funcionário sem CPF identificado."))

        if not dias_funcionario:
            issues.append(_issue("Atenção", nome, cpf, "", "Dias", "Funcionário sem registros de ponto."))

    datas_por_funcionario = defaultdict(list)
    for dia in dados:
        chave_funcionario = dia.get("cpf") or dia.get("nome") or "SEM_IDENTIFICACAO"
        data_norm = _data_normalizada(dia.get("dia", ""))
        if data_norm:
            datas_por_funcionario[(chave_funcionario, data_norm)].append(dia)

        issues.extend(_auditar_dia(dia))

    for (_funcionario, _data), ocorrencias in datas_por_funcionario.items():
        if len(ocorrencias) <= 1:
            continue
        base = ocorrencias[0]
        issues.append(_issue(
            "Atenção",
            base.get("nome", ""),
            base.get("cpf", ""),
            base.get("dia", ""),
            "Data",
            f"Data duplicada para o funcionário ({len(ocorrencias)} ocorrências).",
        ))

    return issues


def build_preview_summary(dados, todos_funcionarios, inconsistencias):
    """Monta um resumo textual/estruturado para a pre-visualizacao."""
    total_funcionarios = len(todos_funcionarios)
    total_dias = len(dados)
    total_trabalhados = len(dias_trabalhados(dados))
    total_faltados = len(dias_faltados(dados))
    total_trabalho = len(dias_trabalho_totais(dados))
    por_severidade = Counter(issue["severidade"] for issue in inconsistencias)

    funcionarios = []
    for _chave, info in sorted(todos_funcionarios.items(), key=lambda item: item[1].get("nome", "")):
        dias = info.get("_dias", [])
        funcionarios.append({
            "nome": info.get("nome", ""),
            "cpf": formatar_cpf(info.get("cpf", "")),
            "dias": len(dias),
            "dias_trabalhados": len(dias_trabalhados(dias)),
            "dias_faltados": len(dias_faltados(dias)),
            "dias_trabalho_totais": len(dias_trabalho_totais(dias)),
        })

    return {
        "total_funcionarios": total_funcionarios,
        "total_dias": total_dias,
        "total_dias_trabalhados": total_trabalhados,
        "total_dias_faltados": total_faltados,
        "total_dias_trabalho": total_trabalho,
        "total_inconsistencias": len(inconsistencias),
        "inconsistencias_por_severidade": dict(por_severidade),
        "funcionarios": funcionarios,
    }


def _auditar_dia(dia):
    issues = []
    nome = dia.get("nome", "")
    cpf = dia.get("cpf", "")
    data = dia.get("dia", "")

    batidas = [dia.get(campo, "").strip() for campo in PUNCH_FIELDS]
    preenchidas = [valor for valor in batidas if valor]
    ch = dia.get("ch", "").strip()

    if not preenchidas:
        if ch:
            issues.append(_issue("Atenção", nome, cpf, data, "Marcações", "Dia de trabalho sem marcações de ponto."))
        return issues

    if len(preenchidas) % 2 != 0:
        issues.append(_issue("Alta", nome, cpf, data, "Marcações", "Quantidade ímpar de marcações."))

    for entrada, saida, par in (("ent1", "sai1", "1"), ("ent2", "sai2", "2"), ("ent3", "sai3", "3")):
        ent = dia.get(entrada, "").strip()
        sai = dia.get(saida, "").strip()
        if ent and not sai:
            issues.append(_issue("Alta", nome, cpf, data, f"Par {par}", f"Entrada {par} sem saída correspondente."))
        if sai and not ent:
            issues.append(_issue("Alta", nome, cpf, data, f"Par {par}", f"Saída {par} sem entrada correspondente."))

    duracao = dia.get("duracao", "").strip()
    if not duracao:
        issues.append(_issue("Atenção", nome, cpf, data, "Duração", "Duração ausente em dia com marcações."))
    elif not _hora_valida(duracao):
        issues.append(_issue("Atenção", nome, cpf, data, "Duração", f"Duração com formato inválido: {duracao}."))

    if not ch:
        issues.append(_issue("Info", nome, cpf, data, "CH", "Código de horário ausente em dia com marcações."))

    return issues


def _issue(severidade, funcionario, cpf, dia, campo, mensagem):
    return {
        "severidade": severidade,
        "funcionario": funcionario,
        "cpf": formatar_cpf(cpf),
        "dia": dia,
        "campo": campo,
        "mensagem": mensagem,
    }


def _data_normalizada(valor):
    data = parse_data_ponto(valor)
    if not valor or data == data.min:
        return ""
    return data.strftime("%Y-%m-%d")


def _hora_valida(valor):
    match = re.match(r"^(\d{1,3}):(\d{2})$", valor)
    if not match:
        return False
    return int(match.group(2)) < 60
