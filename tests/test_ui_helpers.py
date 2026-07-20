import unittest

from main import mapear_tela_para_nav, resumo_ultimo_registro


class UiHelpersTests(unittest.TestCase):
    def test_mapear_tela_para_nav(self):
        self.assertEqual(mapear_tela_para_nav("frame_dashboard"), "dashboard")
        self.assertEqual(mapear_tela_para_nav("frame_menu_phoenix"), "phoenix")
        self.assertEqual(mapear_tela_para_nav("frame_menu_pegasus"), "pegasus")
        self.assertIsNone(mapear_tela_para_nav("frame_login"))

    def test_resumo_ultimo_registro(self):
        historico = [
            {"description": "Primeiro", "status": "done", "linha": "10", "data_abertura": "01/01/2024"},
            {"description": "Segundo", "status": "on going", "linha": "11", "data_abertura": "02/01/2024"},
        ]
        self.assertEqual(
            resumo_ultimo_registro(historico),
            "Último registro: Segundo • Linha 11 • Em andamento"
        )

        self.assertEqual(resumo_ultimo_registro([]), "Último registro: nenhum")


if __name__ == "__main__":
    unittest.main()
