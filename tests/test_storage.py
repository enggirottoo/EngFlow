import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from services import storage


class StorageTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.old_base_dir = storage.BASE_DIR
        self.old_config_path = storage.CONFIG_PATH
        self.old_hist_path = storage.HISTORICO_PATH
        storage.BASE_DIR = self.temp_dir.name
        storage.CONFIG_PATH = os.path.join(self.temp_dir.name, "config_tool.json")
        storage.HISTORICO_PATH = os.path.join(self.temp_dir.name, "historico_solicitacoes.json")

    def tearDown(self):
        storage.BASE_DIR = self.old_base_dir
        storage.CONFIG_PATH = self.old_config_path
        storage.HISTORICO_PATH = self.old_hist_path
        # Evita deixar credenciais de teste no cofre real do sistema.
        try:
            storage._remover_senha_do_cofre("user")
        except Exception:
            pass

    def test_salvar_e_carregar_login(self):
        storage.salvar_login("user", "pass")
        dados = storage.carregar_login()
        self.assertEqual(dados["user"], "user")
        self.assertEqual(dados["password"], "pass")
        self.assertTrue(dados["remember"])

    def test_salvar_login_sem_lembrar_nao_persiste_senha(self):
        storage.salvar_login("user", "pass", lembrar=False)
        dados = storage.carregar_login()
        self.assertEqual(dados["user"], "user")
        self.assertEqual(dados["password"], "")
        self.assertFalse(dados["remember"])

    def test_encontrar_por_linha(self):
        storage.salvar_historico([
            {"linha": 281, "description": "teste", "status": "ON GOING"}
        ])
        self.assertEqual(storage.encontrar_por_linha("281")["description"], "teste")

    def test_criar_registro_descricao(self):
        registro = storage.criar_registro_descricao("Nova descrição", "user")
        self.assertEqual(registro["description"], "Nova descrição")
        self.assertEqual(registro["status"], "ON GOING")
        self.assertTrue(registro["linha"] >= 1)

    def test_salvar_e_carregar_estado_app(self):
        storage.salvar_estado_app("frame_dashboard")
        self.assertEqual(storage.carregar_estado_app().get("last_screen"), "frame_dashboard")


if __name__ == "__main__":
    unittest.main()
