from datetime import datetime
import unittest

from control_id_reader.excel_writer import build_ch_summary_rows, build_workbook, generate_suggested_filename


class ExcelWriterTest(unittest.TestCase):
    def test_generate_suggested_filename_single_employee(self):
        funcionarios = {
            "12345678901": {
                "nome": "Gabriela Almeida",
                "_dias": [{"dia": "05/04/26 - SEG"}],
            }
        }

        nome = generate_suggested_filename([], funcionarios, input_count=1)

        self.assertEqual(nome, "Ponto_Gabriela_Almeida_Abr2026.xlsx")

    def test_generate_suggested_filename_multiple_employees_uses_current_month(self):
        funcionarios = {
            "1": {"nome": "A"},
            "2": {"nome": "B"},
        }

        nome = generate_suggested_filename([], funcionarios, input_count=2, now=datetime(2026, 8, 21))

        self.assertEqual(nome, "Ponto_2Funcionarios_Ago2026.xlsx")

    def test_build_workbook_single_employee(self):
        dados = [{
            "dia": "05/04/26 - SEG",
            "marcacoes": "07:00 12:00 13:00 17:00",
            "ent1": "07:00",
            "sai1": "12:00",
            "ent2": "13:00",
            "sai2": "17:00",
            "ent3": "",
            "sai3": "",
            "duracao": "08:00",
            "ch": "00021",
        }, {
            "dia": "06/04/26 - TER",
            "marcacoes": "",
            "ent1": "",
            "sai1": "",
            "ent2": "",
            "sai2": "",
            "ent3": "",
            "sai3": "",
            "duracao": "",
            "ch": "00021",
        }, {
            "dia": "07/04/26 - QUA",
            "marcacoes": "",
            "ent1": "",
            "sai1": "",
            "ent2": "",
            "sai2": "",
            "ent3": "",
            "sai3": "",
            "duracao": "",
            "ch": "",
        }]
        funcionarios = {
            "12345678901": {
                "nome": "Gabriela Almeida",
                "cpf": "12345678901",
                "horarios_contratuais": [{"codigo": "00021", "ent1": "07:00", "sai1": "12:00", "ent2": "13:00", "sai2": "17:00"}],
                "_dias": dados,
            }
        }

        wb = build_workbook(dados, funcionarios)
        ws = wb["Ponto"]
        valores_coluna_a = [ws.cell(row=row, column=1).value for row in range(1, ws.max_row + 1)]

        self.assertIn("Ponto", wb.sheetnames)
        self.assertIn("Resumo CH", wb.sheetnames)
        self.assertEqual(ws["A1"].value, "DADOS DO FUNCIONÁRIO")
        self.assertIn("RESUMO", valores_coluna_a)
        self.assertIn("DIA", valores_coluna_a)
        self.assertIn("05/04/26 - SEG", valores_coluna_a)
        linha_resumo_cab = valores_coluna_a.index("NOME") + 1
        linha_resumo_valores = linha_resumo_cab + 1
        self.assertEqual(ws.cell(row=linha_resumo_cab, column=9).value, "DIAS TRABALHADOS")
        self.assertEqual(ws.cell(row=linha_resumo_cab, column=10).value, "DIAS FALTADOS")
        self.assertEqual(ws.cell(row=linha_resumo_cab, column=11).value, "DIAS DE TRABALHO TOTAIS")
        self.assertEqual(ws.cell(row=linha_resumo_valores, column=9).value, 1)
        self.assertEqual(ws.cell(row=linha_resumo_valores, column=10).value, 1)
        self.assertEqual(ws.cell(row=linha_resumo_valores, column=11).value, 2)
        self.assertEqual(wb["Resumo CH"]["A2"].value, "Gabriela Almeida")
        self.assertEqual(wb["Resumo CH"]["C2"].value, "00021")
        self.assertEqual(wb["Resumo CH"]["D2"].value, "07:00-12:00 / 13:00-17:00")
        self.assertEqual(wb["Resumo CH"]["E2"].value, 1)
        self.assertEqual(wb["Resumo CH"]["F2"].value, 1)
        self.assertEqual(wb["Resumo CH"]["G2"].value, 2)

    def test_build_workbook_multiple_employees_creates_summary_sheet(self):
        dados = [
            {"nome": "Bia", "cpf": "222.222.222-22", "dia": "06/04/26 - TER", "ent1": "08:00", "sai1": "17:00", "ch": "00021"},
            {"nome": "Ana", "cpf": "111.111.111-11", "dia": "05/04/26 - SEG", "ent1": "08:00", "sai1": "17:00", "ch": "00021"},
        ]
        funcionarios = {
            "22222222222": {"nome": "Bia", "cpf": "22222222222", "_dias": [dados[0]]},
            "11111111111": {"nome": "Ana", "cpf": "11111111111", "_dias": [dados[1]]},
        }

        wb = build_workbook(dados, funcionarios)

        self.assertEqual(wb.sheetnames, ["Ponto", "Resumo CH", "Resumo"])
        self.assertEqual(wb["Ponto"]["A1"].value, "FUNCIONÁRIO")
        self.assertEqual(wb["Ponto"]["A2"].value, "Ana")
        self.assertEqual(wb["Resumo"]["A1"].value, "NOME")
        self.assertEqual(wb["Resumo"]["J1"].value, "DIAS FALTADOS")
        self.assertEqual(wb["Resumo"]["K1"].value, "DIAS DE TRABALHO TOTAIS")

    def test_build_workbook_orders_sheets_with_inconsistencies(self):
        dados = [
            {"nome": "Ana", "cpf": "111.111.111-11", "dia": "05/04/26 - SEG", "ent1": "08:00", "sai1": "", "ch": "00021"},
            {"nome": "Bia", "cpf": "222.222.222-22", "dia": "06/04/26 - TER", "ent1": "08:00", "sai1": "17:00", "ch": "00021"},
        ]
        funcionarios = {
            "11111111111": {"nome": "Ana", "cpf": "11111111111", "_dias": [dados[0]]},
            "22222222222": {"nome": "Bia", "cpf": "22222222222", "_dias": [dados[1]]},
        }
        inconsistencias = [{
            "severidade": "Alta",
            "funcionario": "Ana",
            "cpf": "111.111.111-11",
            "dia": "05/04/26 - SEG",
            "campo": "Marcações",
            "mensagem": "Quantidade ímpar de marcações.",
        }]

        wb = build_workbook(dados, funcionarios, inconsistencias=inconsistencias)

        self.assertEqual(wb.sheetnames, ["Ponto", "Inconsistências", "Resumo CH", "Resumo"])

    def test_build_ch_summary_rows_counts_by_employee_and_ch(self):
        funcionarios = {
            "12345678901": {
                "nome": "Ana",
                "cpf": "12345678901",
                "horarios_contratuais": [{"codigo": "00001", "ent1": "08:00", "sai1": "12:00"}],
                "_dias": [
                    {"ch": "00001", "ent1": "08:00", "sai1": "12:00"},
                    {"ch": "00001", "ent1": "", "sai1": ""},
                    {"ch": "00002", "ent1": "14:00", "sai1": "18:00"},
                    {"ch": "", "ent1": "", "sai1": ""},
                ],
            }
        }

        rows = build_ch_summary_rows(funcionarios)

        self.assertEqual(rows[0], {
            "funcionario": "Ana",
            "cpf": "123.456.789-01",
            "ch": "00001",
            "horario": "08:00-12:00",
            "dias_trabalhados": 1,
            "dias_faltados": 1,
            "dias_trabalho_totais": 2,
        })
        self.assertEqual(rows[1]["ch"], "00002")
        self.assertEqual(rows[1]["dias_trabalhados"], 1)

    def test_build_workbook_adds_inconsistencies_sheet(self):
        dados = [{"nome": "Ana", "cpf": "", "dia": "05/04/26 - SEG", "ent1": "08:00", "sai1": ""}]
        funcionarios = {"Ana": {"nome": "Ana", "cpf": "", "_dias": dados}}
        inconsistencias = [{
            "severidade": "Alta",
            "funcionario": "Ana",
            "cpf": "",
            "dia": "05/04/26 - SEG",
            "campo": "Marcações",
            "mensagem": "Quantidade ímpar de marcações.",
        }]

        wb = build_workbook(dados, funcionarios, inconsistencias=inconsistencias)

        self.assertIn("Inconsistências", wb.sheetnames)
        self.assertEqual(wb["Inconsistências"]["A1"].value, "SEVERIDADE")
        self.assertEqual(wb["Inconsistências"]["F2"].value, "Quantidade ímpar de marcações.")


if __name__ == "__main__":
    unittest.main()
