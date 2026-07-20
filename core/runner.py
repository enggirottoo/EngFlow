"""Utilitários para disparar as automações (scripts em `automocoes/`) como
processos separados e para abrir a planilha no navegador padrão.

Mantido fora de `main.py` para separar "como eu chamo uma automação" da
lógica de construção de telas da GUI.
"""

from __future__ import annotations

import os
import subprocess
import sys
import tkinter as tk
import webbrowser
from tkinter import messagebox
from typing import Optional

from services.logging_config import configurar_logger

logger = configurar_logger("phoenix_tool", arquivo_log="phoenix_tool.log")

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

PLANILHA_URL = (
    "https://exemplo365.sharepoint.com/:x:/r/sites/TestEngineeringExemplo/"
    "_layouts/15/Doc.aspx?sourcedoc=%7B00000000-0000-0000-0000-000000000000%7D&action=edit"
)


def executar_script(*partes: str, arg: Optional[str] = None) -> None:
    """Lança um script de automação em um processo Python separado.

    No Windows, o processo filho recebe seu próprio console (visível para o
    usuário acompanhar prints/`input()` da automação). Se o processo morrer
    logo após iniciar (ex.: dependência ausente), uma mensagem de erro é
    exibida em vez de falhar silenciosamente.
    """
    try:
        caminho = os.path.join(BASE_DIR, *partes)
        if not os.path.isfile(caminho):
            logger.error("Script não encontrado: %s", caminho)
            messagebox.showerror("Erro", f"Script não encontrado:\n{caminho}")
            return

        comando = [sys.executable, caminho]
        if arg:
            comando.append(arg)

        kwargs = {}
        if os.name == "nt":
            kwargs["creationflags"] = subprocess.CREATE_NEW_CONSOLE

        logger.info("Executando automação: %s", caminho)
        processo = subprocess.Popen(comando, cwd=os.path.dirname(caminho), **kwargs)

        def _verificar_saida_imediata() -> None:
            codigo = processo.poll()
            if codigo is not None and codigo != 0:
                logger.error("Automação '%s' encerrou com código %s", os.path.basename(caminho), codigo)
                messagebox.showerror(
                    "Erro na automação",
                    f"O script '{os.path.basename(caminho)}' encerrou logo após iniciar "
                    f"(código {codigo}).\nVerifique se todas as dependências estão instaladas "
                    "no ambiente Python usado pelo Phoenix Tool.",
                )

        root = tk._get_default_root()
        if root is not None:
            root.after(1500, _verificar_saida_imediata)
    except Exception as exc:
        logger.exception("Falha ao executar script")
        messagebox.showerror("Erro", str(exc))


def abrir_planilha() -> None:
    try:
        webbrowser.open(PLANILHA_URL)
    except Exception as exc:
        logger.exception("Falha ao abrir a planilha")
        messagebox.showerror("Erro", str(exc))
