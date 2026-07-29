"""Backup automático do banco SQLite do Phoenix Tool.

Cria cópias do banco em:
- Na inicialização da ferramenta
- Antes de atualizações críticas (chamadas explicitamente)
- No encerramento da ferramenta

Mantém os últimos MAX_BACKUPS arquivos, removendo os mais antigos.
"""

from __future__ import annotations

import logging
import os
import shutil
from datetime import datetime
from typing import List, Optional

logger = logging.getLogger(__name__)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(BASE_DIR, "phoenix_tool.db")
BACKUP_DIR = os.path.join(BASE_DIR, "backups")
MAX_BACKUPS = 30


def _garantir_dir() -> None:
    os.makedirs(BACKUP_DIR, exist_ok=True)


def fazer_backup(motivo: str = "manual") -> Optional[str]:
    """Cria uma cópia do banco com timestamp.

    Args:
        motivo: Descrição do motivo (ex: 'inicializacao', 'encerramento', 'pre-atualizacao').

    Returns:
        Caminho do arquivo de backup criado, ou None em caso de falha.
    """
    if not os.path.exists(DB_PATH):
        logger.debug("Backup ignorado: banco %s não existe ainda.", DB_PATH)
        return None

    _garantir_dir()
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome = f"backup_{ts}_{motivo[:20].replace(' ', '_')}.db"
    destino = os.path.join(BACKUP_DIR, nome)

    try:
        shutil.copy2(DB_PATH, destino)
        logger.info("Backup criado: %s", destino)
        _limpar_backups_antigos()
        return destino
    except Exception as exc:
        logger.warning("Falha ao criar backup (%s): %s", motivo, exc)
        return None


def listar_backups() -> List[dict]:
    """Retorna lista de backups disponíveis ordenados do mais recente ao mais antigo."""
    _garantir_dir()
    arquivos = []
    for nome in os.listdir(BACKUP_DIR):
        if nome.startswith("backup_") and nome.endswith(".db"):
            caminho = os.path.join(BACKUP_DIR, nome)
            stat = os.stat(caminho)
            tamanho_kb = stat.st_size / 1024
            data_mod = datetime.fromtimestamp(stat.st_mtime).strftime("%d/%m/%Y %H:%M:%S")
            arquivos.append({
                "nome": nome,
                "caminho": caminho,
                "tamanho_kb": round(tamanho_kb, 1),
                "data": data_mod,
            })
    arquivos.sort(key=lambda x: x["nome"], reverse=True)
    return arquivos


def restaurar_backup(caminho: str) -> bool:
    """Substitui o banco atual pelo backup especificado.

    ATENÇÃO: Cria um backup de segurança do banco atual antes de restaurar.
    Deve ser chamado APENAS por código com nível ADMINISTRADOR.

    Returns:
        True se restaurado com sucesso, False caso contrário.
    """
    if not os.path.exists(caminho):
        logger.error("Restauração falhou: arquivo %s não encontrado.", caminho)
        return False

    # Backup de segurança do banco atual antes de restaurar
    fazer_backup("pre-restauracao")

    try:
        shutil.copy2(caminho, DB_PATH)
        logger.info("Banco restaurado a partir de: %s", caminho)
        return True
    except Exception as exc:
        logger.error("Falha ao restaurar backup: %s", exc)
        return False


def _limpar_backups_antigos() -> None:
    """Remove os backups mais antigos mantendo apenas MAX_BACKUPS."""
    _garantir_dir()
    backups = listar_backups()
    if len(backups) > MAX_BACKUPS:
        para_remover = backups[MAX_BACKUPS:]
        for b in para_remover:
            try:
                os.remove(b["caminho"])
                logger.debug("Backup antigo removido: %s", b["nome"])
            except Exception as exc:
                logger.warning("Não foi possível remover backup antigo %s: %s", b["nome"], exc)
