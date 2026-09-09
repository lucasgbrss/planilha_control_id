import unittest

from control_id_reader.audit import analyze_inconsistencies, build_preview_summary


class AuditTest(unittest.TestCase):
    def test_analyze_inconsistencies_detects_missing_cpf_and_odd_punches(self):
        dados = [{
            "nome": "Ana",
            "cpf": "",
            "dia": "05/04/26 - SEG",
            "ent1": "08:00",
            "sai1": "",
            "ent2": "",
            "sai2": "",
            "ent3": "",
            "sai3": "",
            "duracao": "",
            "ch": "",
        }]
        funcionarios = {"Ana": {"nome": "Ana", "cpf": "", "_dias": dados}}

        issues = analyze_inconsistencies(dados, funcionarios)
        mensagens = [issue["mensagem"] for issue in issues]

        self.assertIn("Funcionário sem CPF identificado.", mensagens)
        self.assertIn("Quantidade ímpar de marcações.", mensagens)
        self.assertIn("Entrada 1 sem saída correspondente.", mensagens)
        self.assertIn("Duração ausente em dia com marcações.", mensagens)
        self.assertIn("Código de horário ausente em dia com marcações.", mensagens)

    def test_analyze_inconsistencies_detects_duplicate_date(self):
        dados = [
            {"nome": "Bia", "cpf": "123.456.789-01", "dia": "05/04/26 - SEG", "ent1": "08:00", "sai1": "17:00", "duracao": "09:00", "ch": "00021"},
            {"nome": "Bia", "cpf": "123.456.789-01", "dia": "05/04/2026 - SEG", "ent1": "08:10", "sai1": "17:10", "duracao": "09:00", "ch": "00021"},
        ]
        funcionarios = {"12345678901": {"nome": "Bia", "cpf": "12345678901", "_dias": dados}}

        issues = analyze_inconsistencies(dados, funcionarios)

        self.assertTrue(any("Data duplicada" in issue["mensagem"] for issue in issues))

    def test_build_preview_summary_counts_employees_and_issues(self):
        dados = [
            {"nome": "Ana", "cpf": "", "dia": "05/04/26 - SEG", "ent1": "", "sai1": "", "ch": "00021"},
            {"nome": "Ana", "cpf": "", "dia": "06/04/26 - TER", "ent1": "", "sai1": "", "ch": ""},
        ]
        funcionarios = {"Ana": {"nome": "Ana", "cpf": "", "_dias": dados}}
        issues = analyze_inconsistencies(dados, funcionarios)

        resumo = build_preview_summary(dados, funcionarios, issues)

        self.assertEqual(resumo["total_funcionarios"], 1)
        self.assertEqual(resumo["total_dias"], 2)
        self.assertEqual(resumo["total_dias_trabalhados"], 0)
        self.assertEqual(resumo["total_dias_faltados"], 1)
        self.assertEqual(resumo["total_dias_trabalho"], 1)
        self.assertGreaterEqual(resumo["total_inconsistencias"], 1)
        self.assertEqual(resumo["funcionarios"][0]["nome"], "Ana")

    def test_ch_vazio_sem_batidas_eh_folga(self):
        dados = [{
            "nome": "Ana",
            "cpf": "12345678901",
            "dia": "06/04/26 - TER",
            "ent1": "",
            "sai1": "",
            "ent2": "",
            "sai2": "",
            "ent3": "",
            "sai3": "",
            "duracao": "",
            "ch": "",
        }]
        funcionarios = {"12345678901": {"nome": "Ana", "cpf": "12345678901", "_dias": dados}}

        issues = analyze_inconsistencies(dados, funcionarios)

        self.assertEqual(issues, [])

    def test_ch_preenchido_sem_batidas_eh_falta(self):
        dados = [{
            "nome": "Ana",
            "cpf": "12345678901",
            "dia": "07/04/26 - QUA",
            "ent1": "",
            "sai1": "",
            "ent2": "",
            "sai2": "",
            "ent3": "",
            "sai3": "",
            "duracao": "",
            "ch": "00021",
        }]
        funcionarios = {"12345678901": {"nome": "Ana", "cpf": "12345678901", "_dias": dados}}

        issues = analyze_inconsistencies(dados, funcionarios)

        self.assertEqual(len(issues), 1)
        self.assertEqual(issues[0]["mensagem"], "Dia de trabalho sem marcações de ponto.")


if __name__ == "__main__":
    unittest.main()
