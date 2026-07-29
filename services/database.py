"""Motor de persistência SQLite do Phoenix Tool.

Substitui o armazenamento baseado em JSON, provendo:
- UUID único por registro (nunca depende de ticket/description como chave)
- Histórico de eventos imutável (append-only, sem UPDATE/DELETE em historico_eventos)
- Identidade completa do usuário (usuario, hostname, versão, data/hora)
- Soft-delete (status=CANCELADO, nunca exclusão física)
- Log de tentativas bloqueadas e eventos de segurança
- Migração automática do JSON legado na primeira inicialização

A API pública espelha os nomes de `storage.py` para compatibilidade com
todos os importadores existentes.
"""

from __future__ import annotations

import json
import logging
import os
import shutil
import socket
import sqlite3
import uuid
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(BASE_DIR, "phoenix_tool.db")
JSON_LEGADO_PATH = os.path.join(BASE_DIR, "historico_solicitacoes.json")
JSON_BACKUP_PATH = os.path.join(BASE_DIR, "historico_solicitacoes.json.bak")

APP_VERSION = "2.0.0"

# ---------------------------------------------------------------------------
# DDL — esquema do banco
# ---------------------------------------------------------------------------

_DDL = """
PRAGMA journal_mode=WAL;
PRAGMA foreign_keys=ON;

CREATE TABLE IF NOT EXISTS solicitacoes (
    uuid                TEXT PRIMARY KEY,
    linha               INTEGER UNIQUE NOT NULL,
    ticket              TEXT DEFAULT '',
    description         TEXT DEFAULT '',
    produto             TEXT DEFAULT '',
    mpn                 TEXT DEFAULT '',
    origem              TEXT DEFAULT 'IMPORTADO',
    custo               TEXT DEFAULT '',
    custo_finalizado    INTEGER DEFAULT 0,
    pegasus_abertura TEXT DEFAULT '',
    pegasus_fechamento TEXT DEFAULT '',
    custo_abertura TEXT DEFAULT '',
    custo_fechamento TEXT DEFAULT '',
    produto_seq         INTEGER DEFAULT 1,
    total_produtos      INTEGER DEFAULT 1,
    data_abertura       TEXT DEFAULT '',
    hora_abertura       TEXT DEFAULT '',
    status              TEXT DEFAULT 'ON GOING',
    pn                  TEXT DEFAULT '',
    part_number         TEXT DEFAULT '',
    data_fechamento     TEXT DEFAULT '',
    analista            TEXT DEFAULT '',
    data_pn             TEXT DEFAULT '',
    hora_pn             TEXT DEFAULT '',
    criado_por          TEXT DEFAULT '',
    criado_em           TEXT DEFAULT '',
    criado_pc           TEXT DEFAULT '',
    versao_criacao      TEXT DEFAULT '',
    ultima_alteracao    TEXT DEFAULT '',
    ultimo_usuario      TEXT DEFAULT '',
    ultimo_pc           TEXT DEFAULT '',
    ultima_versao       TEXT DEFAULT '',
    solicitante         TEXT DEFAULT '',
    requisitante        TEXT DEFAULT '',
    cancelado           INTEGER DEFAULT 0,
    motivo_cancelamento TEXT DEFAULT ''
);

CREATE TABLE IF NOT EXISTS historico_eventos (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    registro_uuid TEXT NOT NULL,
    data_hora     TEXT NOT NULL,
    usuario       TEXT NOT NULL,
    pc            TEXT NOT NULL,
    versao        TEXT DEFAULT '',
    tipo_evento   TEXT NOT NULL,
    descricao     TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS permissoes (
    usuario_windows TEXT PRIMARY KEY,
    nivel           TEXT NOT NULL DEFAULT 'USUARIO',
    concedido_por   TEXT DEFAULT '',
    concedido_em    TEXT DEFAULT ''
);

CREATE TABLE IF NOT EXISTS log_seguranca (
    id        INTEGER PRIMARY KEY AUTOINCREMENT,
    data_hora TEXT NOT NULL,
    usuario   TEXT NOT NULL,
    pc        TEXT NOT NULL,
    evento    TEXT NOT NULL,
    detalhes  TEXT DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_sol_ticket  ON solicitacoes(ticket);
CREATE INDEX IF NOT EXISTS idx_sol_status  ON solicitacoes(status);
CREATE INDEX IF NOT EXISTS idx_sol_criador ON solicitacoes(criado_por);
CREATE INDEX IF NOT EXISTS idx_hist_uuid   ON historico_eventos(registro_uuid);
"""


# ---------------------------------------------------------------------------
# Conexão e inicialização
# ---------------------------------------------------------------------------

def _conectar() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def inicializar_banco() -> None:
    """Cria as tabelas se não existirem e executa migração do JSON legado."""
    conn = _conectar()
    try:
        conn.executescript(_DDL)
        conn.commit()
        _migrar_json_legado(conn)
        _garantir_primeiro_admin(conn)
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Identidade do usuário/sistema
# ---------------------------------------------------------------------------

def obter_identidade() -> Dict[str, str]:
    """Retorna dicionário com usuário Windows, hostname, data/hora e versão."""
    import getpass
    try:
        usuario = getpass.getuser().upper()
    except Exception:
        usuario = "SISTEMA"
    try:
        pc = socket.gethostname().upper()
    except Exception:
        pc = "DESCONHECIDO"
    return {
        "usuario": usuario,
        "pc": pc,
        "data_hora": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "versao": APP_VERSION,
    }


# ---------------------------------------------------------------------------
# Migração automática do JSON legado
# ---------------------------------------------------------------------------

def _migrar_json_legado(conn: sqlite3.Connection) -> None:
    """Importa registros do JSON legado se o banco ainda estiver vazio."""
    cursor = conn.execute("SELECT COUNT(*) FROM solicitacoes")
    if cursor.fetchone()[0] > 0:
        return  # banco já tem dados — não reimportar

    if not os.path.exists(JSON_LEGADO_PATH):
        return

    try:
        with open(JSON_LEGADO_PATH, "r", encoding="utf-8") as f:
            dados = json.load(f)
        if not isinstance(dados, list):
            return
    except Exception as exc:
        logger.warning("Migração JSON: erro ao ler arquivo legado: %s", exc)
        return

    identidade = obter_identidade()
    agora = identidade["data_hora"]

    migrados = 0
    for item in dados:
        try:
            reg_uuid = str(uuid.uuid4())
            linha = int(item.get("linha", 0)) or migrados + 1
            criador = str(item.get("criado_por") or item.get("user") or identidade["usuario"]).upper()
            criado_em = str(item.get("criado_em") or agora)
            pn = str(item.get("pn") or item.get("part_number") or "")
            origem = str(item.get("origem") or "IMPORTADO")
            custo = str(item.get("custo") or "")
            custo_fin = int(bool(item.get("custo_finalizado")))

            conn.execute("""
                INSERT OR IGNORE INTO solicitacoes (
                    uuid, linha, ticket, description, produto, mpn,
                    origem, custo, custo_finalizado,
                    produto_seq, total_produtos,
                    data_abertura, hora_abertura, status,
                    pn, part_number, data_fechamento,
                    analista, data_pn, hora_pn,
                    criado_por, criado_em, criado_pc, versao_criacao,
                    ultima_alteracao, ultimo_usuario, ultimo_pc, ultima_versao,
                    solicitante, requisitante
                ) VALUES (
                    ?,?,?,?,?,?,
                    ?,?,?,
                    ?,?,
                    ?,?,?,
                    ?,?,?,
                    ?,?,?,
                    ?,?,?,?,
                    ?,?,?,?,
                    ?,?
                )
            """, (
                reg_uuid, linha,
                str(item.get("ticket") or ""),
                str(item.get("description") or ""),
                str(item.get("produto") or ""),
                str(item.get("mpn") or ""),
                origem, custo, custo_fin,
                int(item.get("produto_seq") or 1),
                int(item.get("total_produtos") or 1),
                str(item.get("data_abertura") or ""),
                str(item.get("hora_abertura") or ""),
                str(item.get("status") or "ON GOING"),
                pn, pn,
                str(item.get("data_fechamento") or ""),
                str(item.get("analista") or ""),
                str(item.get("data_pn") or ""),
                str(item.get("hora_pn") or ""),
                criador, criado_em, identidade["pc"], "migrado-json",
                str(item.get("ultima_alteracao") or criado_em),
                str(item.get("ultimo_usuario_alterou") or criador),
                identidade["pc"], "migrado-json",
                str(item.get("solicitante") or ""),
                str(item.get("requisitante") or ""),
            ))

            # Migrar histórico de alterações do JSON
            historico_json = item.get("historico_alteracoes") or []
            if not historico_json:
                historico_json = [{"data": criado_em, "usuario": criador, "descricao": "Abertura da solicitação (migrado)"}]
            for h in historico_json:
                conn.execute("""
                    INSERT INTO historico_eventos (registro_uuid, data_hora, usuario, pc, versao, tipo_evento, descricao)
                    VALUES (?,?,?,?,?,?,?)
                """, (
                    reg_uuid,
                    str(h.get("data") or agora),
                    str(h.get("usuario") or criador).upper(),
                    identidade["pc"],
                    "migrado-json",
                    "MIGRACAO",
                    str(h.get("descricao") or "Evento migrado do JSON"),
                ))
            migrados += 1
        except Exception as exc:
            logger.warning("Migração JSON: erro no registro linha=%s: %s", item.get("linha"), exc)

    conn.commit()
    if migrados > 0:
        logger.info("Migração JSON → SQLite: %s registros importados.", migrados)
        # Renomear JSON original para .bak
        try:
            shutil.move(JSON_LEGADO_PATH, JSON_BACKUP_PATH)
            logger.info("Arquivo JSON original renomeado para .bak")
        except Exception as exc:
            logger.warning("Não foi possível renomear o JSON legado: %s", exc)


# ---------------------------------------------------------------------------
# Garantir primeiro Admin
# ---------------------------------------------------------------------------

def _garantir_primeiro_admin(conn: sqlite3.Connection) -> None:
    """Se não houver nenhum admin cadastrado, registra o usuário atual como admin."""
    cursor = conn.execute("SELECT COUNT(*) FROM permissoes WHERE nivel='ADMINISTRADOR'")
    if cursor.fetchone()[0] > 0:
        return
    identidade = obter_identidade()
    agora = identidade["data_hora"]
    conn.execute("""
        INSERT OR IGNORE INTO permissoes (usuario_windows, nivel, concedido_por, concedido_em)
        VALUES (?, 'ADMINISTRADOR', ?, ?)
    """, (identidade["usuario"], "SISTEMA", agora))
    conn.commit()
    logger.info("Primeiro admin registrado automaticamente: %s", identidade["usuario"])


# ---------------------------------------------------------------------------
# Funções auxiliares internas
# ---------------------------------------------------------------------------

def _row_to_dict(row) -> Dict[str, Any]:
    """Converte sqlite3.Row para dict Python, incluindo historico_alteracoes."""
    if row is None:
        return {}
    d = dict(row)
    # Compatibilidade com código legado que usa 'user' e 'ultimo_usuario_alterou'
    d.setdefault("user", d.get("criado_por", ""))
    d.setdefault("ultimo_usuario_alterou", d.get("ultimo_usuario", ""))
    d["historico_alteracoes"] = _carregar_historico_eventos(d["uuid"])
    return d


def _carregar_historico_eventos(reg_uuid: str) -> List[Dict[str, Any]]:
    conn = _conectar()
    try:
        rows = conn.execute("""
            SELECT data_hora, usuario, pc, versao, tipo_evento, descricao
            FROM historico_eventos WHERE registro_uuid=? ORDER BY id ASC
        """, (reg_uuid,)).fetchall()
        return [
            {
                "data": r["data_hora"],
                "usuario": r["usuario"],
                "pc": r["pc"],
                "versao": r["versao"],
                "tipo": r["tipo_evento"],
                "descricao": r["descricao"],
            }
            for r in rows
        ]
    finally:
        conn.close()


def _registrar_evento(conn: sqlite3.Connection, reg_uuid: str, tipo: str, descricao: str, identidade: Dict[str, str]) -> None:
    """Insere um evento no histórico imutável."""
    conn.execute("""
        INSERT INTO historico_eventos (registro_uuid, data_hora, usuario, pc, versao, tipo_evento, descricao)
        VALUES (?,?,?,?,?,?,?)
    """, (
        reg_uuid,
        identidade.get("data_hora", datetime.now().strftime("%d/%m/%Y %H:%M")),
        identidade.get("usuario", "SISTEMA"),
        identidade.get("pc", "DESCONHECIDO"),
        identidade.get("versao", APP_VERSION),
        tipo,
        descricao,
    ))


def _registrar_log_seguranca(conn: sqlite3.Connection, evento: str, detalhes: str, identidade: Dict[str, str]) -> None:
    conn.execute("""
        INSERT INTO log_seguranca (data_hora, usuario, pc, evento, detalhes)
        VALUES (?,?,?,?,?)
    """, (
        identidade.get("data_hora", datetime.now().strftime("%d/%m/%Y %H:%M")),
        identidade.get("usuario", "SISTEMA"),
        identidade.get("pc", "DESCONHECIDO"),
        evento,
        detalhes,
    ))


def _proxima_linha(conn: sqlite3.Connection) -> int:
    row = conn.execute("SELECT MAX(linha) FROM solicitacoes").fetchone()
    return (row[0] or 0) + 1


def _verificar_e_encerrar(conn: sqlite3.Connection, reg_uuid: str, identidade: Dict[str, str]) -> bool:
    """Se PN + Origem + Custo preenchidos → status ENCERRADO."""
    row = conn.execute(
        "SELECT pn, origem, custo, custo_finalizado, status FROM solicitacoes WHERE uuid=?", (reg_uuid,)
    ).fetchone()
    if not row:
        return False
    pn = str(row["pn"] or "").strip()
    origem = str(row["origem"] or "").strip()
    custo = str(row["custo"] or "").strip()
    custo_fin = bool(row["custo_finalizado"])
    status = str(row["status"] or "").upper()
    if pn and origem and (custo_fin or custo) and status != "ENCERRADO":
        conn.execute(
            "UPDATE solicitacoes SET status='ENCERRADO', ultima_alteracao=?, ultimo_usuario=?, ultimo_pc=?, ultima_versao=? WHERE uuid=?",
            (identidade["data_hora"], identidade["usuario"], identidade["pc"], identidade["versao"], reg_uuid)
        )
        _registrar_evento(conn, reg_uuid, "ENCERRAMENTO", "Encerramento da solicitação", identidade)
        return True
    return False


# ---------------------------------------------------------------------------
# API Pública — espelha storage.py
# ---------------------------------------------------------------------------

def proxima_linha() -> int:
    conn = _conectar()
    try:
        return _proxima_linha(conn)
    finally:
        conn.close()


def carregar_historico() -> List[Dict[str, Any]]:
    """Retorna todos os registros não cancelados como lista de dicts."""
    conn = _conectar()
    try:
        rows = conn.execute(
            "SELECT * FROM solicitacoes WHERE cancelado=0 ORDER BY linha ASC"
        ).fetchall()
        return [_row_to_dict(r) for r in rows]
    finally:
        conn.close()


def salvar_historico(historico: List[Dict[str, Any]]) -> None:
    """Compatibilidade: atualiza campos de cada registro em batch.
    Usado apenas por código legado que ainda monta e reenvia o histórico inteiro."""
    conn = _conectar()
    try:
        identidade = obter_identidade()
        for item in historico:
            reg_uuid = item.get("uuid")
            if not reg_uuid:
                continue
            campos_basicos = {k: v for k, v in item.items() if k not in ("uuid", "historico_alteracoes")}
            _atualizar_campos_por_uuid_interno(conn, reg_uuid, campos_basicos, identidade, registrar_evento=False)
        conn.commit()
    finally:
        conn.close()


def encontrar_por_linha(linha: Any) -> Optional[Dict[str, Any]]:
    conn = _conectar()
    try:
        row = conn.execute("SELECT * FROM solicitacoes WHERE linha=?", (int(linha),)).fetchone()
        return _row_to_dict(row) if row else None
    finally:
        conn.close()


def encontrar_por_uuid(reg_uuid: str) -> Optional[Dict[str, Any]]:
    conn = _conectar()
    try:
        row = conn.execute("SELECT * FROM solicitacoes WHERE uuid=?", (reg_uuid,)).fetchone()
        return _row_to_dict(row) if row else None
    finally:
        conn.close()


def encontrar_por_ticket(ticket: str) -> Optional[Dict[str, Any]]:
    ticket = str(ticket or "").strip()
    if not ticket:
        return None
    conn = _conectar()
    try:
        row = conn.execute(
            "SELECT * FROM solicitacoes WHERE ticket=? AND cancelado=0 LIMIT 1", (ticket,)
        ).fetchone()
        return _row_to_dict(row) if row else None
    finally:
        conn.close()


def encontrar_todos_por_ticket(ticket: str) -> List[Dict[str, Any]]:
    ticket = str(ticket or "").strip()
    if not ticket:
        return []
    conn = _conectar()
    try:
        rows = conn.execute(
            "SELECT * FROM solicitacoes WHERE ticket=? AND cancelado=0 ORDER BY produto_seq ASC", (ticket,)
        ).fetchall()
        return [_row_to_dict(r) for r in rows]
    finally:
        conn.close()


def existe_ticket(ticket: str) -> bool:
    ticket = str(ticket or "").strip()
    if not ticket:
        return False
    conn = _conectar()
    try:
        row = conn.execute("SELECT 1 FROM solicitacoes WHERE ticket=? LIMIT 1", (ticket,)).fetchone()
        return row is not None
    finally:
        conn.close()


def criar_registro_descricao(
    descricao: str,
    usuario: Optional[str] = None,
    solicitante: Optional[str] = None,
    requisitante: Optional[str] = None,
    produto: Optional[str] = None,
    ticket: Optional[str] = None,
    status: Optional[str] = None,
    data_abertura: Optional[str] = None,
    hora_abertura: Optional[str] = None,
    produto_seq: Optional[int] = None,
    total_produtos: Optional[int] = None,
    mpn: Optional[str] = None,
    origem: Optional[str] = None,
    custo: Optional[str] = None,
    custo_finalizado: Optional[bool] = None,
) -> Dict[str, Any]:
    identidade = obter_identidade()
    if usuario:
        identidade["usuario"] = usuario.upper()

    data_ab = data_abertura or datetime.now().strftime("%d/%m/%Y")
    hora_ab = hora_abertura or datetime.now().strftime("%H:%M")
    criado_em = f"{data_ab} {hora_ab}".strip()
    reg_uuid = str(uuid.uuid4())

    conn = _conectar()
    try:
        linha = _proxima_linha(conn)
        conn.execute("""
            INSERT INTO solicitacoes (
                uuid, linha, ticket, description, produto, mpn,
                origem, custo, custo_finalizado,
                produto_seq, total_produtos,
                data_abertura, hora_abertura, status,
                pn, part_number,
                criado_por, criado_em, criado_pc, versao_criacao,
                ultima_alteracao, ultimo_usuario, ultimo_pc, ultima_versao,
                solicitante, requisitante
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            reg_uuid, linha,
            (ticket or "").strip(),
            descricao.strip(),
            (produto or "").strip(),
            (mpn or "").strip(),
            (origem or "IMPORTADO").strip(),
            (custo or "").strip(),
            int(bool(custo_finalizado)),
            int(produto_seq) if produto_seq is not None else 1,
            int(total_produtos) if total_produtos is not None else 1,
            data_ab, hora_ab,
            status or "ON GOING",
            "", "",
            identidade["usuario"], criado_em, identidade["pc"], identidade["versao"],
            criado_em, identidade["usuario"], identidade["pc"], identidade["versao"],
            (solicitante or "").strip(),
            (requisitante or "").strip(),
        ))
        _registrar_evento(conn, reg_uuid, "CRIACAO", "Abertura da solicitação", identidade)
        _verificar_e_encerrar(conn, reg_uuid, identidade)
        conn.commit()
        row = conn.execute("SELECT * FROM solicitacoes WHERE uuid=?", (reg_uuid,)).fetchone()
        return _row_to_dict(row)
    finally:
        conn.close()


def _atualizar_campos_por_uuid_interno(
    conn: sqlite3.Connection,
    reg_uuid: str,
    campos: Dict[str, Any],
    identidade: Dict[str, str],
    registrar_evento: bool = True,
) -> None:
    """Atualiza campos de um registro por UUID (uso interno, conn já aberta)."""
    # Campos permitidos para atualização (exclui chaves imutáveis)
    _IMUTAVEIS = {"uuid", "linha", "id", "criado_por", "criado_em", "criado_pc", "versao_criacao", "historico_alteracoes"}
    campos_validos = {k: v for k, v in campos.items() if k not in _IMUTAVEIS}
    if not campos_validos:
        return

    # Detectar mudanças para eventos específicos
    if registrar_evento:
        row_antes = conn.execute("SELECT * FROM solicitacoes WHERE uuid=?", (reg_uuid,)).fetchone()
        if row_antes:
            pn_antes = str(row_antes["pn"] or "").strip()
            origem_antes = str(row_antes["origem"] or "").strip()
            custo_antes = str(row_antes["custo"] or "").strip()

    set_clause = ", ".join(f"{k}=?" for k in campos_validos)
    values = list(campos_validos.values()) + [
        identidade["data_hora"], identidade["usuario"], identidade["pc"], identidade["versao"], reg_uuid
    ]
    conn.execute(
        f"UPDATE solicitacoes SET {set_clause}, ultima_alteracao=?, ultimo_usuario=?, ultimo_pc=?, ultima_versao=? WHERE uuid=?",
        values
    )

    if registrar_evento and row_antes:
        pn_depois = str(campos_validos.get("pn", pn_antes) or "").strip()
        origem_depois = str(campos_validos.get("origem", origem_antes) or "").strip()
        custo_depois = str(campos_validos.get("custo", custo_antes) or "").strip()
        custo_fin = bool(campos_validos.get("custo_finalizado", False))

        if pn_depois and pn_depois != pn_antes:
            _registrar_evento(conn, reg_uuid, "ATUALIZACAO_PN", f"Atualização de PN: {pn_depois}", identidade)
        if origem_depois and origem_depois != origem_antes:
            _registrar_evento(conn, reg_uuid, "ALTERACAO_ORIGEM", f"Alteração de origem para {origem_depois}", identidade)
        if custo_depois and (custo_depois != custo_antes or custo_fin):
            _registrar_evento(conn, reg_uuid, "FINALIZACAO_CUSTO", "Finalização de custo", identidade)
        if not any([
            pn_depois != pn_antes,
            origem_depois != origem_antes,
            custo_depois != custo_antes,
            custo_fin,
        ]):
            _registrar_evento(conn, reg_uuid, "ALTERACAO", "Informações atualizadas pelo usuário", identidade)


def atualizar_campos_registro(
    linha: Any,
    campos: Dict[str, Any],
    usuario_alteracao: Optional[str] = None,
) -> Optional[Dict[str, Any]]:
    identidade = obter_identidade()
    if usuario_alteracao:
        identidade["usuario"] = usuario_alteracao.upper()

    conn = _conectar()
    try:
        row = conn.execute("SELECT uuid, status FROM solicitacoes WHERE linha=?", (int(linha),)).fetchone()
        if not row:
            return None
        reg_uuid = row["uuid"]
        _atualizar_campos_por_uuid_interno(conn, reg_uuid, campos, identidade)
        _verificar_e_encerrar(conn, reg_uuid, identidade)
        conn.commit()
        updated = conn.execute("SELECT * FROM solicitacoes WHERE uuid=?", (reg_uuid,)).fetchone()
        return _row_to_dict(updated)
    finally:
        conn.close()


def atualizar_campos_por_uuid(
    reg_uuid: str,
    campos: Dict[str, Any],
    usuario_alteracao: Optional[str] = None,
) -> Optional[Dict[str, Any]]:
    identidade = obter_identidade()
    if usuario_alteracao:
        identidade["usuario"] = usuario_alteracao.upper()

    conn = _conectar()
    try:
        row = conn.execute("SELECT uuid FROM solicitacoes WHERE uuid=?", (reg_uuid,)).fetchone()
        if not row:
            return None
        _atualizar_campos_por_uuid_interno(conn, reg_uuid, campos, identidade)
        _verificar_e_encerrar(conn, reg_uuid, identidade)
        conn.commit()
        updated = conn.execute("SELECT * FROM solicitacoes WHERE uuid=?", (reg_uuid,)).fetchone()
        return _row_to_dict(updated)
    finally:
        conn.close()


def atualizar_campos_por_ticket(
    ticket: str,
    campos: Dict[str, Any],
    usuario_alteracao: Optional[str] = None,
) -> Optional[Dict[str, Any]]:
    ticket = str(ticket or "").strip()
    if not ticket:
        return None
    identidade = obter_identidade()
    if usuario_alteracao:
        identidade["usuario"] = usuario_alteracao.upper()

    conn = _conectar()
    try:
        row = conn.execute(
            "SELECT uuid FROM solicitacoes WHERE ticket=? AND cancelado=0 LIMIT 1", (ticket,)
        ).fetchone()
        if not row:
            return None
        reg_uuid = row["uuid"]
        _atualizar_campos_por_uuid_interno(conn, reg_uuid, campos, identidade)
        _verificar_e_encerrar(conn, reg_uuid, identidade)
        conn.commit()
        updated = conn.execute("SELECT * FROM solicitacoes WHERE uuid=?", (reg_uuid,)).fetchone()
        return _row_to_dict(updated)
    finally:
        conn.close()


def atualizar_registros_por_ticket(
    ticket: str,
    lista_pns: List[str],
    dados_step: Dict[str, Any],
    usuario_alteracao: Optional[str] = None,
) -> List[Dict[str, Any]]:
    ticket = str(ticket or "").strip()
    if not ticket:
        return []
    identidade = obter_identidade()
    if usuario_alteracao:
        identidade["usuario"] = usuario_alteracao.upper()

    conn = _conectar()
    try:
        rows = conn.execute(
            "SELECT uuid, produto_seq, pn FROM solicitacoes WHERE ticket=? AND cancelado=0 ORDER BY produto_seq ASC",
            (ticket,)
        ).fetchall()

        atualizados = []
        for idx, row in enumerate(rows):
            reg_uuid = row["uuid"]
            pn_novo = lista_pns[idx] if idx < len(lista_pns) else str(row["pn"] or "")
            pn_atual = str(row["pn"] or "").strip()
            pn_mudou = bool(pn_novo and pn_novo != pn_atual)

            campos = {"pn": pn_novo, "part_number": pn_novo}
            if dados_step.get("status"):
                campos["status"] = dados_step["status"]
            if dados_step.get("data_pn"):
                campos["data_pn"] = dados_step["data_pn"]
            if dados_step.get("hora_pn"):
                campos["hora_pn"] = dados_step["hora_pn"]
            if dados_step.get("analista"):
                campos["analista"] = dados_step["analista"]
            if dados_step.get("data_fechamento"):
                campos["data_fechamento"] = dados_step["data_fechamento"]

            set_clause = ", ".join(f"{k}=?" for k in campos)
            values = list(campos.values()) + [
                identidade["data_hora"], identidade["usuario"], identidade["pc"], identidade["versao"], reg_uuid
            ]
            conn.execute(
                f"UPDATE solicitacoes SET {set_clause}, ultima_alteracao=?, ultimo_usuario=?, ultimo_pc=?, ultima_versao=? WHERE uuid=?",
                values
            )

            if pn_mudou:
                _registrar_evento(conn, reg_uuid, "ATUALIZACAO_PN", f"Atualização de PN: {pn_novo}", identidade)
            else:
                _registrar_evento(conn, reg_uuid, "ALTERACAO", "Status do ticket atualizado", identidade)

            _verificar_e_encerrar(conn, reg_uuid, identidade)
            updated = conn.execute("SELECT * FROM solicitacoes WHERE uuid=?", (reg_uuid,)).fetchone()
            atualizados.append(_row_to_dict(updated))

        if atualizados:
            conn.commit()
        return atualizados
    finally:
        conn.close()


def cancelar_registro(
    reg_uuid: str,
    motivo: str,
    usuario_cancelamento: Optional[str] = None,
) -> bool:
    """Soft-delete: marca o registro como CANCELADO sem remover do banco."""
    identidade = obter_identidade()
    if usuario_cancelamento:
        identidade["usuario"] = usuario_cancelamento.upper()

    conn = _conectar()
    try:
        row = conn.execute("SELECT uuid FROM solicitacoes WHERE uuid=?", (reg_uuid,)).fetchone()
        if not row:
            return False
        conn.execute(
            "UPDATE solicitacoes SET cancelado=1, status='CANCELADO', motivo_cancelamento=?, ultima_alteracao=?, ultimo_usuario=?, ultimo_pc=? WHERE uuid=?",
            (motivo, identidade["data_hora"], identidade["usuario"], identidade["pc"], reg_uuid)
        )
        _registrar_evento(conn, reg_uuid, "CANCELAMENTO", f"Registro cancelado. Motivo: {motivo}", identidade)
        conn.commit()
        return True
    finally:
        conn.close()


def registrar_tentativa_bloqueada(
    descricao: str,
    detalhes: str = "",
    usuario: Optional[str] = None,
) -> None:
    """Registra no log de segurança uma tentativa de ação não permitida."""
    identidade = obter_identidade()
    if usuario:
        identidade["usuario"] = usuario.upper()
    conn = _conectar()
    try:
        _registrar_log_seguranca(conn, descricao, detalhes, identidade)
        conn.commit()
    finally:
        conn.close()


def obter_log_seguranca(limite: int = 100) -> List[Dict[str, Any]]:
    conn = _conectar()
    try:
        rows = conn.execute(
            "SELECT * FROM log_seguranca ORDER BY id DESC LIMIT ?", (limite,)
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Permissões
# ---------------------------------------------------------------------------

NIVEIS = ("USUARIO", "SUPERVISOR", "ADMINISTRADOR")


def obter_nivel_usuario(usuario_windows: Optional[str] = None) -> str:
    if not usuario_windows:
        usuario_windows = obter_identidade()["usuario"]
    usuario_windows = usuario_windows.upper()
    conn = _conectar()
    try:
        row = conn.execute(
            "SELECT nivel FROM permissoes WHERE usuario_windows=?", (usuario_windows,)
        ).fetchone()
        return row["nivel"] if row else "USUARIO"
    finally:
        conn.close()


def definir_nivel_usuario(
    usuario_alvo: str,
    nivel: str,
    usuario_admin: Optional[str] = None,
) -> bool:
    if nivel not in NIVEIS:
        return False
    identidade = obter_identidade()
    if usuario_admin:
        identidade["usuario"] = usuario_admin.upper()
    conn = _conectar()
    try:
        conn.execute("""
            INSERT INTO permissoes (usuario_windows, nivel, concedido_por, concedido_em)
            VALUES (?,?,?,?)
            ON CONFLICT(usuario_windows) DO UPDATE SET
                nivel=excluded.nivel,
                concedido_por=excluded.concedido_por,
                concedido_em=excluded.concedido_em
        """, (usuario_alvo.upper(), nivel, identidade["usuario"], identidade["data_hora"]))
        conn.commit()
        return True
    finally:
        conn.close()


def listar_permissoes() -> List[Dict[str, Any]]:
    conn = _conectar()
    try:
        rows = conn.execute("SELECT * FROM permissoes ORDER BY usuario_windows").fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()











# ---------------------------------------------------------------------------
# Validação de integridade
# ---------------------------------------------------------------------------

def validar_integridade() -> List[str]:
    """Executa verificações de integridade no banco. Retorna lista de problemas."""
    problemas = []
    conn = _conectar()
    try:
        # IDs duplicados (uuid)
        rows = conn.execute(
            "SELECT uuid, COUNT(*) as c FROM solicitacoes GROUP BY uuid HAVING c > 1"
        ).fetchall()
        for r in rows:
            problemas.append(f"UUID duplicado: {r['uuid']}")

        # Linhas duplicadas
        rows = conn.execute(
            "SELECT linha, COUNT(*) as c FROM solicitacoes GROUP BY linha HAVING c > 1"
        ).fetchall()
        for r in rows:
            problemas.append(f"Linha duplicada: {r['linha']}")

        # produto_seq inválido
        rows = conn.execute(
            "SELECT uuid, linha, produto_seq, total_produtos FROM solicitacoes WHERE produto_seq > total_produtos OR produto_seq < 1"
        ).fetchall()
        for r in rows:
            problemas.append(f"produto_seq inválido na linha {r['linha']}: seq={r['produto_seq']}, total={r['total_produtos']}")

        # Campos obrigatórios ausentes
        rows = conn.execute(
            "SELECT uuid, linha FROM solicitacoes WHERE (description IS NULL OR description='') AND cancelado=0"
        ).fetchall()
        for r in rows:
            problemas.append(f"Description ausente na linha {r['linha']}")

        # Verificação de integridade do SQLite
        result = conn.execute("PRAGMA integrity_check").fetchone()
        if result and result[0] != "ok":
            problemas.append(f"Falha na integridade do banco SQLite: {result[0]}")

    finally:
        conn.close()
    return problemas



def iniciar_pegasus(linha):
    return atualizar_campos_registro(
        linha,
        {
            "pegasus_abertura":
            datetime.now().strftime("%d/%m/%Y %H:%M")
        }
    )


def finalizar_pegasus(linha):
    return atualizar_campos_registro(
        linha,
        {
            "pegasus_fechamento":
            datetime.now().strftime("%d/%m/%Y %H:%M")
        }
    )

def iniciar_custo(linha):
    return atualizar_campos_registro(
        linha,
        {
            "custo_abertura":
            datetime.now().strftime("%d/%m/%Y %H:%M")
        }
    )

def finalizar_custo(linha):
    return atualizar_campos_registro(
        linha,
        {
            "custo_fechamento":
            datetime.now().strftime("%d/%m/%Y %H:%M")
        }
    )

