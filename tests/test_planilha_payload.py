import json
import os
import tempfile
import unittest

from automocoes.planilha.planilha import construir_payload_planilha, salvar_payload_planilha


class PlanilhaPayloadTests(unittest.TestCase):
    def test_construir_payload_planilha(self):
        payload = construir_payload_planilha("Descrição teste", linha="12", usuario="estagiaria")
        self.assertEqual(payload["descricao"], "Descrição teste")
        self.assertEqual(payload["linha"], "12")
        self.assertEqual(payload["usuario"], "estagiaria")
        self.assertEqual(payload["status"], "ON GOING")

    def test_salvar_payload_planilha(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            arquivo = os.path.join(tmpdir, "planilha.json")
            payload = {"descricao": "x"}
            salvar_payload_planilha(payload, arquivo=arquivo)
            with open(arquivo, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.assertEqual(data["descricao"], "x")


if __name__ == "__main__":
    unittest.main()
