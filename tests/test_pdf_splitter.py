import unittest

from control_id_reader.pdf_splitter import detect_employee_page_groups_from_pages, unique_pdf_filename
from tests.card_fixtures import card_tables


class FakePage:
    def __init__(self, text, tables=None):
        self.text = text
        self.tables = tables or []

    def extract_text(self):
        return self.text

    def extract_tables(self):
        return self.tables


class PdfSplitterTest(unittest.TestCase):
    def test_detect_employee_page_groups_from_pages(self):
        pages = [
            FakePage("NOME: Ana Souza PIS/PASEP: 123"),
            FakePage("continuação da Ana"),
            FakePage("NOME: Bia Lima PIS/PASEP: 456"),
            FakePage("continuação da Bia"),
        ]

        grupos = detect_employee_page_groups_from_pages(pages)

        self.assertEqual(grupos, [
            {"nome": "Ana Souza", "paginas": [0, 1]},
            {"nome": "Bia Lima", "paginas": [2, 3]},
        ])

    def test_unique_pdf_filename_sanitizes_and_avoids_collisions(self):
        nomes_usados = {}

        primeiro = unique_pdf_filename("José Santos", nomes_usados)
        segundo = unique_pdf_filename("Jose Santos", nomes_usados)

        self.assertEqual(primeiro, "Jose_Santos.pdf")
        self.assertEqual(segundo, "Jose_Santos_2.pdf")

    def test_detects_card_pages_by_table_header(self):
        pages = [
            FakePage("Cart�o de Ponto", card_tables("Ana Souza")),
            FakePage("Cart�o de Ponto", card_tables("Bia Lima")),
        ]

        self.assertEqual(detect_employee_page_groups_from_pages(pages), [
            {"nome": "Ana Souza", "paginas": [0]},
            {"nome": "Bia Lima", "paginas": [1]},
        ])


if __name__ == "__main__":
    unittest.main()
