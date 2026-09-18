import unittest

from control_id_reader.parsers import (
    extract_card_employee_info_from_tables,
    extract_card_punch_rows_from_tables,
    extract_pdf_employees,
    extract_punch_rows_from_tables,
    extract_schedule_row,
)
from tests.card_fixtures import card_tables


class FakePdfPage:
    def __init__(self, tables):
        self.tables = tables

    def extract_text(self):
        return "Cart�o de Ponto"

    def extract_tables(self):
        return self.tables


class FakePdf:
    def __init__(self, pages):
        self.pages = pages


class CardParserTest(unittest.TestCase):
    def test_time_inside_justification_is_not_a_punch(self):
        tables = card_tables()
        tables[1].insert(-1, [
            "06/08/2026 - QUI", "08:00-12:00 13:00-16:00",
            "Atestado a partir de 08:00", "", "", "", "", "", "", "", "",
        ])

        dia = extract_card_punch_rows_from_tables(tables)[-1]

        self.assertEqual(dia["ent1"], "")
        self.assertTrue(dia["ausencia_justificada"])

    def test_card_employee_and_daily_rows(self):
        tables = card_tables()
        info = extract_card_employee_info_from_tables(tables)
        dias = extract_card_punch_rows_from_tables(tables)

        self.assertEqual(info["nome"], "Ana Souza")
        self.assertEqual(info["cpf"], "12345678901")
        self.assertEqual(info["matricula"], "7")
        self.assertEqual(len(dias), 5)
        self.assertEqual(dias[0]["marcacoes"], "Folga")
        self.assertEqual(dias[1]["ent1"], "23:48")
        self.assertEqual(dias[1]["sai2"], "08:07")
        self.assertEqual(dias[1]["duracao"], "07:34")
        self.assertEqual(dias[1]["total_noturno"], "05:57")
        self.assertEqual(dias[1]["ch"], "")
        self.assertEqual(dias[2]["ent1"], "")
        self.assertFalse(dias[2]["ausencia_justificada"])
        self.assertTrue(dias[3]["ausencia_justificada"])
        self.assertTrue(dias[4]["ausencia_justificada"])

    def test_pdf_pages_are_separate_employees(self):
        pdf = FakePdf([FakePdfPage(card_tables("Ana Souza")), FakePdfPage(card_tables("Bia Lima"))])

        resultados = extract_pdf_employees(pdf)

        self.assertEqual([info["nome"] for info, _dias in resultados], ["Ana Souza", "Bia Lima"])
        self.assertEqual([len(dias) for _info, dias in resultados], [5, 5])


class ParserScheduleTest(unittest.TestCase):
    def test_extract_schedule_row_ignores_empty_interleaved_cells(self):
        row = ["00021", "07:00", "", "12:00", "", "13:00", "", "17:00"]

        horario = extract_schedule_row(row)

        self.assertEqual(horario, {
            "codigo": "00021",
            "ent1": "07:00",
            "sai1": "12:00",
            "ent2": "13:00",
            "sai2": "17:00",
        })

    def test_extract_schedule_table_with_garbled_control_id_header(self):
        tabelas = [
            ["C�DIGO DO HOR�RIO(CH)", "ENT", "", "SAI", "", "ENT", "", "SAI"],
            ["00021", "07:00", "", "12:00", "", "13:00", "", "17:00"],
            ["00022", "07:00", "", "12:00", "", "13:00", "", "16:00"],
        ]

        dias, horarios = extract_punch_rows_from_tables(tabelas)

        self.assertEqual(dias, [])
        self.assertEqual(horarios, [
            {
                "codigo": "00021",
                "ent1": "07:00",
                "sai1": "12:00",
                "ent2": "13:00",
                "sai2": "17:00",
            },
            {
                "codigo": "00022",
                "ent1": "07:00",
                "sai1": "12:00",
                "ent2": "13:00",
                "sai2": "16:00",
            },
        ])


if __name__ == "__main__":
    unittest.main()
