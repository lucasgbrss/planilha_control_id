import unittest

from bs4 import BeautifulSoup

from control_id_reader.parsers import (
    extract_employee_info_from_text,
    extract_punch_rows_from_mhtml,
    extract_punch_rows_from_tables,
)
from control_id_reader.utils import (
    formatar_cpf,
    media_horas,
    mes_ano_predominante,
    nome_para_arquivo,
    normalizar_cpf,
    parse_data_ponto,
    somar_duracoes,
)
from pdf_para_excel import (
    PdfToExcelApp,
)


class UtilsTest(unittest.TestCase):
    def test_normaliza_e_formata_cpf(self):
        self.assertEqual(normalizar_cpf("123.456.789-01"), "12345678901")
        self.assertEqual(formatar_cpf("12345678901"), "123.456.789-01")

    def test_parse_data_aceita_ano_curto_e_longo(self):
        self.assertEqual(parse_data_ponto("05/04/26 - SEG").year, 2026)
        self.assertEqual(parse_data_ponto("05/04/2026 - SEG").year, 2026)

    def test_mes_ano_predominante_expande_ano_curto(self):
        dias = [
            {"dia": "05/04/26 - SEG"},
            {"dia": "06/04/26 - TER"},
            {"dia": "01/05/26 - SEX"},
        ]
        self.assertEqual(mes_ano_predominante(dias), "Abr2026")

    def test_soma_e_media_duracoes(self):
        total = somar_duracoes(["08:00", "07:30", "", "invalido"])
        self.assertEqual(total, "15:30")
        self.assertEqual(media_horas(total, 2), "07:45")

    def test_nome_para_arquivo_remove_acentos_e_sinais(self):
        self.assertEqual(nome_para_arquivo("Joao da Silva"), "Joao_da_Silva")
        self.assertEqual(nome_para_arquivo("José / Santos"), "Jose_Santos")


class PdfToExcelAppPureMethodsTest(unittest.TestCase):
    def setUp(self):
        self.app = PdfToExcelApp.__new__(PdfToExcelApp)
        self.app.pdf_paths = ["ponto.pdf"]

    def test_gerar_nome_arquivo_usa_ano_completo(self):
        funcionarios = {
            "12345678901": {
                "nome": "Gabriela Almeida",
                "_dias": [{"dia": "05/04/26 - SEG"}],
            }
        }
        self.assertEqual(
            self.app._gerar_nome_arquivo([], funcionarios),
            "Ponto_Gabriela_Almeida_Abr2026.xlsx",
        )

    def test_extrair_info_funcionario_campos_com_espacos(self):
        texto = (
            "EMPRESA: ACME LTDA CNPJ: 12.345.678/0001-90\n"
            "NOME: Maria Souza PIS/PASEP: 123456789\n"
            "ADMISSÃO: 01/01/2024 CPF: 123.456.789-01 MATRÍCULA: 42 "
            "CENTRO DE CUSTO: ADMINISTRATIVO CENTRAL DEPARTAMENTO: RECURSOS HUMANOS "
            "CARGO: ANALISTA DE DP\n"
        )
        info, motivo = extract_employee_info_from_text(texto)
        self.assertIsNone(motivo)
        self.assertEqual(info["cpf"], "12345678901")
        self.assertEqual(info["centro_custo"], "ADMINISTRATIVO CENTRAL")
        self.assertEqual(info["departamento"], "RECURSOS HUMANOS")


class ParserTest(unittest.TestCase):
    def test_extract_punch_rows_from_tables_mesclado(self):
        tabelas = [
            ["05/04/26 - SEG", "07:00 12:00 13:00 17:00", "07:00", "12:00", "13:00 17:00", "", "08:00", "00021"],
        ]

        dias, horarios = extract_punch_rows_from_tables(tabelas)

        self.assertEqual(horarios, [])
        self.assertEqual(len(dias), 1)
        self.assertEqual(dias[0]["ent2"], "13:00")
        self.assertEqual(dias[0]["sai2"], "17:00")
        self.assertEqual(dias[0]["duracao"], "08:00")
        self.assertEqual(dias[0]["ch"], "00021")

    def test_extract_punch_rows_from_tables_tempo_quebrado(self):
        tabelas = [
            ["05/04/26 - SEG", "", "19:00", "23:00", "19:00 00", ":05", "04:05", "00021"],
        ]

        dias, _ = extract_punch_rows_from_tables(tabelas)

        self.assertEqual(dias[0]["ent2"], "19:00")
        self.assertEqual(dias[0]["sai2"], "00:05")

    def test_extract_punch_rows_from_mhtml(self):
        html = """
        <table>
          <tr><th>DIA</th><th></th></tr>
          <tr>
            <td>05/04/26 - SEG</td><td></td><td></td><td>07:00</td><td>12:00</td>
            <td>13:00</td><td>17:00</td><td></td><td></td><td></td><td>08:00</td>
            <td></td><td>00021</td><td></td><td></td><td></td><td></td>
          </tr>
        </table>
        """
        soup = BeautifulSoup(html, "html.parser")

        dias = extract_punch_rows_from_mhtml(soup)

        self.assertEqual(len(dias), 1)
        self.assertEqual(dias[0]["dia"], "05/04/2026 - SEG")
        self.assertEqual(dias[0]["ent1"], "07:00")
        self.assertEqual(dias[0]["duracao"], "08:00")


if __name__ == "__main__":
    unittest.main()
