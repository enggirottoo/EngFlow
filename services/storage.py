"""Persistência local do Phoenix Tool.

Este módulo concentra toda a leitura/escrita dos arquivos JSON usados pela
aplicação (`config_tool.json`, `historico_solicitacoes.json`,
`estado_app.json`) e é importado tanto pela GUI (`main.py`) quanto pelas
automações em `automocoes/` — assim existe um único ponto de verdade para
essas operações.

Segurança de credenciais
-------------------------
A senha do usuário nunca deveria viver em texto puro em disco. Por isso,
`salvar_login`/`carregar_login` tentam guardar a senha no cofre de
credenciais do sistema operacional através do pacote `keyring` (no Windows,
o Windows Credential Manager). Isso só é usado quando o usuário optar por
"lembrar" a senha; caso contrário nada é persistido.

Se o `keyring` não estiver disponível no ambiente (pacote não instalado ou
sem backend suportado), o código cai para o comportamento antigo (salvar em
texto puro no `config_tool.json`) para não quebrar a automação, mas registra
um aviso no log para deixar isso visível.
"""

from __future__ import annotations

import json
import logging
import os
from datetime import datetime
from typing import Any, Dict, List, Optional

try:
    import keyring
    from keyring.errors import KeyringError
except ImportError:  # pacote opcional - cai para armazenamento em arquivo
    keyring = None  # type: ignore[assignment]

    class KeyringError(Exception):
        """Usado como stand-in quando o pacote `keyring` não está instalado."""

logger = logging.getLogger(__name__)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(BASE_DIR, "config_tool.json")
HISTORICO_PATH = os.path.join(BASE_DIR, "historico_solicitacoes.json")
ESTADO_APP_PATH = os.path.join(BASE_DIR, "estado_app.json")

# Nome do "serviço" usado para agrupar as credenciais no cofre do sistema.
KEYRING_SERVICE = "PhoenixTool"


# =====================================================
# Config (usuário / senha)
# =====================================================

def _ler_config_bruta() -> Dict[str, Any]:
    """Lê o config_tool.json sem interpretar nenhum campo específico."""
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            dados = json.load(f)
            return dados if isinstance(dados, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def carregar_config() -> Dict[str, Any]:
    """Retorna o conteúdo bruto de config_tool.json como dicionário."""
    return _ler_config_bruta()


def _escrever_config_bruta(dados: Dict[str, Any]) -> None:
    """Grava o dicionário completo no config_tool.json, preservando chaves
    que outras partes do sistema (ex.: cache de caminho da planilha) já
    tenham guardado ali."""
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(dados, f, indent=4, ensure_ascii=False)


def _salvar_senha_no_cofre(usuario: str, senha: str) -> bool:
    """Tenta guardar a senha no cofre de credenciais do SO. Retorna True em caso de sucesso."""
    if keyring is None or not usuario:
        return False
    try:
        keyring.set_password(KEYRING_SERVICE, usuario, senha)
        return True
    except KeyringError as exc:
        logger.warning("Não foi possível salvar a senha no cofre do sistema: %s", exc)
        return False


def _carregar_senha_do_cofre(usuario: str) -> Optional[str]:
    if keyring is None or not usuario:
        return None
    try:
        return keyring.get_password(KEYRING_SERVICE, usuario)
    except KeyringError as exc:
        logger.warning("Não foi possível ler a senha do cofre do sistema: %s", exc)
        return None


def _remover_senha_do_cofre(usuario: str) -> None:
    if keyring is None or not usuario:
        return
    try:
        keyring.delete_password(KEYRING_SERVICE, usuario)
    except KeyringError:
        pass  # não havia senha salva para este usuário - nada a fazer


def salvar_login(usuario: str, senha: str, lembrar: bool = True) -> None:
    """Salva o usuário e, se `lembrar` for True, a senha.

    A senha é priorizada para o cofre de credenciais do sistema operacional;
    só é escrita em texto puro no `config_tool.json` se o `keyring` não
    estiver disponível no ambiente atual.
    """
    dados = _ler_config_bruta()
    dados["user"] = usuario
    dados["remember"] = bool(lembrar) and bool(usuario)
    dados.pop("password", None)

    if dados["remember"]:
        if not _salvar_senha_no_cofre(usuario, senha):
            dados["password"] = senha  # fallback inseguro - só quando o cofre falha
    else:
        _remover_senha_do_cofre(usuario)

    _escrever_config_bruta(dados)


def carregar_login() -> Dict[str, Any]:
    """Retorna `{"user": ..., "password": ..., "remember": ...}`.

    Configs antigos (salvos antes do suporte a "lembrar-me") tinham a senha
    sempre em texto puro no arquivo; nesse caso tratamos como se
    `remember=True` para não quebrar o auto-login de quem já usava a
    ferramenta.
    """
    dados = _ler_config_bruta()
    usuario = str(dados.get("user", ""))
    lembrar = bool(dados.get("remember", "password" in dados))

    senha = ""
    if lembrar:
        senha = _carregar_senha_do_cofre(usuario) or str(dados.get("password", ""))

    return {"user": usuario, "password": senha, "remember": lembrar}


def limpar_senha_salva(usuario: str) -> None:
    """Remove a senha salva (cofre e/ou arquivo) para o usuário informado."""
    dados = _ler_config_bruta()
    dados["remember"] = False
    dados.pop("password", None)
    _escrever_config_bruta(dados)
    _remover_senha_do_cofre(usuario)


# =====================================================
# Histórico de solicitações
# =====================================================

def carregar_historico() -> List[Dict[str, Any]]:
    try:
        with open(HISTORICO_PATH, "r", encoding="utf-8") as f:
            dados = json.load(f)
            if isinstance(dados, list):
                return dados
            if isinstance(dados, dict) and isinstance(dados.get("solicitacoes"), list):
                return dados["solicitacoes"]
    except (OSError, json.JSONDecodeError):
        return []
    return []


def salvar_historico(historico: List[Dict[str, Any]]) -> None:
    with open(HISTORICO_PATH, "w", encoding="utf-8") as f:
        json.dump(historico, f, ensure_ascii=False, indent=4)


def encontrar_por_linha(linha: str) -> Optional[Dict[str, Any]]:
    for item in carregar_historico():
        if str(item.get("linha")) == str(linha):
            return item
    return None


def proxima_linha() -> int:
    historico = carregar_historico()
    linhas = [int(item.get("linha")) for item in historico if str(item.get("linha", "")).isdigit()]
    return max(linhas, default=0) + 1


def criar_registro_descricao(descricao: str, usuario: Optional[str] = None) -> Dict[str, Any]:
    historico = carregar_historico()
    registro = {
        "id": len(historico) + 1,
        "linha": proxima_linha(),
        "ticket": "",
        "description": descricao.strip(),
        "data_abertura": datetime.now().strftime("%d/%m/%Y"),
        "hora_abertura": datetime.now().strftime("%H:%M"),
        "status": "ON GOING",
        "pn": "",
        "data_fechamento": "",
        "user": usuario or "",
    }
    historico.append(registro)
    salvar_historico(historico)
    return registro


def salvar_estado_app(tela: str) -> None:
    with open(ESTADO_APP_PATH, "w", encoding="utf-8") as f:
        json.dump({"last_screen": tela}, f, ensure_ascii=False, indent=4)


def carregar_estado_app() -> Dict[str, str]:
    try:
        with open(ESTADO_APP_PATH, "r", encoding="utf-8") as f:
            dados = json.load(f)
            if isinstance(dados, dict):
                return {"last_screen": str(dados.get("last_screen", ""))}
    except (OSError, json.JSONDecodeError):
        return {"last_screen": ""}
    return {"last_screen": ""}
