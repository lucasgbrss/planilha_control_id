import json
import unittest

from control_id_reader.config_store import parse_config_content, serialize_config


class ConfigStoreTest(unittest.TestCase):
    def test_serialize_and_parse_config(self):
        content = serialize_config({
            "dir_abrir_pdf": "C:/Entrada",
            "dir_salvar_excel": "C:/Saida",
            "tema": "dark",
        })
        config = parse_config_content(content)

        self.assertEqual(config["dir_abrir_pdf"], "C:/Entrada")
        self.assertEqual(config["dir_salvar_excel"], "C:/Saida")
        self.assertEqual(config["tema"], "dark")
        self.assertIn("data_ultima_execucao", config)
        self.assertIsInstance(json.loads(content), dict)

    def test_load_config_returns_empty_dict_for_invalid_json(self):
        self.assertEqual(parse_config_content("{invalid"), {})


if __name__ == "__main__":
    unittest.main()
