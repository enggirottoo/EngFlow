"""Configuração centralizada de logging para o Phoenix Tool.

Todas as automações (Phoenix, Pegasus, Cost Request, Planilha) e a própria
GUI usam esta função para obter um logger configurado de forma consistente,
em vez de espalhar chamadas a `print()` pelo código.

As automações rodam em um console separado (`CREATE_NEW_CONSOLE`), então o
handler de console continua sendo a forma do usuário acompanhar o progresso
em tempo real; a GUI grava adicionalmente em um arquivo de log para permitir
diagnóstico depois que a janela é fechada.
"""

from __future__ import annotations

import logging
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOG_FORMAT = "%(asctime)s | %(levelname)-7s | %(message)s"
DATE_FORMAT = "%H:%M:%S"


def configurar_logger(nome: str, arquivo_log: str | None = None) -> logging.Logger:
    """Retorna um logger configurado com saída no console (e, opcionalmente, em arquivo).

    Chamar esta função várias vezes com o mesmo `nome` é seguro: os handlers
    só são adicionados uma vez por logger.
    """
    logger = logging.getLogger(nome)
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    formatter = logging.Formatter(LOG_FORMAT, datefmt=DATE_FORMAT)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    if arquivo_log:
        caminho = os.path.join(BASE_DIR, arquivo_log)
        try:
            file_handler = logging.FileHandler(caminho, encoding="utf-8")
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
        except OSError:
            logger.warning("Não foi possível abrir arquivo de log em %s", caminho)

    logger.propagate = False
    return logger
