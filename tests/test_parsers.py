import unittest

from control_id_reader.parsers import extract_punch_rows_from_tables, extract_schedule_row


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
