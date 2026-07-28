"""Automação: Atualizar PN Phoenix
================================

Implementa as Partes 3 a 7 do fluxo:

  3) Login automático no Phoenix + navegação até "Solicitações Finalizadas"
  4) Localizar a solicitação certa (Ticket > Description + Data)
  5) Entrar em STEPS STATUS e capturar o PN no step "Criação de Part Number"
  6) Capturar a Data de Fechamento
  7) Atualizar o registro no historico_solicitacoes.json

Chamado pelo main.py assim (mesmo padrão de subprocess já usado no projeto):

    python atualizar_pn.py <linha>

IMPORTANTE - SELETORES
-----------------------
Os trechos marcados com "# AJUSTAR SELETOR" são palpites baseados no padrão
que já existe em `phoenix.py` (ids tipo #FlexUser, localizadores por texto,
etc). Use o Playwright Inspector para confirmar:

    PWDEBUG=1 python atualizar_pn.py 281

ou insira `page.pause()` em qualquer ponto do fluxo para inspecionar o HTML
ao vivo e ajustar o seletor certo.
"""

from __future__ import annotations

import os
import re
import sys
from datetime import datetime
from typing import Any, Dict, List, Optional

from playwright.sync_api import Page, sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# garante que "phoenix.py" (mesma pasta) seja importável mesmo quando
# este arquivo é chamado como subprocess a partir do main.py
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from phoenix import (  # noqa: E402  (reaproveita login já pronto)
    carregar_config,
    _abrir_phoenix,
    _debug_screenshot,
    _diagnostico,
    _ir_para_minhas_solicitacoes,
)
from services.logging_config import configurar_logger  # noqa: E402
from services.storage import (  # noqa: E402
    atualizar_campos_por_ticket,
    atualizar_registros_por_ticket,
    carregar_historico,
    encontrar_por_ticket,
    encontrar_todos_por_ticket,
    salvar_historico,
)

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError  # noqa: E402

# Grava também em arquivo: quando chamada em processo/thread sem console
# visível (botão "Atualizar PN Phoenix" do dashboard), não há terminal para
# mostrar os logs - sem arquivo, uma falha aqui seria invisível.
logger = configurar_logger("atualizar_pn", arquivo_log="phoenix_tool.log")

# Rótulo de status usado quando a solicitação é finalizada.
STATUS_FINALIZADO = "OK"

# Heurística de formato de PN: 3 dígitos - 3 letras - alfanumérico
# (ex: 361-FXB-D202606183J). Ajuste se o padrão real variar.
PN_REGEX = re.compile(r"([0-9]{3}-[A-Z]{3}-[A-Za-z0-9]+)")

# Rótulo amigável salvo no registro quando o step "Criação de Part Number"
# está FINALIZADO no Phoenix.
STATUS_STEP_PN_CRIADO = "Part Number Criado"

# Linhas de texto conhecidas do corpo da resposta do step, ignoradas na
# hora de tentar identificar o nome do analista. AJUSTAR se o texto padrão
# usado pelo Phoenix for diferente.
_BOILERPLATE_RESPOSTA = {
    "hi team",
    "parts",
    "this signal processed in this eco",
}

# Heurística de nome de analista: duas ou mais palavras iniciadas em
# maiúscula (ex: "Jabin Emmanuel"). AJUSTAR se o Phoenix expuser o analista
# num seletor próprio em vez de texto solto na linha do step.
_NOME_ANALISTA_REGEX = re.compile(r"^[A-ZÀ-Ý][a-zà-ÿ]+(?:\s+[A-ZÀ-Ý][a-zà-ÿ]+)+$")


def buscar_registro(historico: List[Dict[str, Any]], linha: str) -> Optional[Dict[str, Any]]:
    for item in historico:
        if str(item.get("linha")) == str(linha):
            return item
    return None


# =====================================================
# NAVEGAÇÃO NO PHOENIX
# =====================================================

def ir_para_finalizadas(page: Page) -> None:
    """A partir da Home, entra em 'Solicitações Finalizadas'."""
    # AJUSTAR SELETOR: texto/menu real da Home do Phoenix
    page.wait_for_selector("text=Solicitações Finalizadas", timeout=15000)
    page.click("text=Solicitações Finalizadas")
    page.wait_for_load_state("networkidle")
    logger.info("Tela 'Solicitações Finalizadas' aberta.")


def pesquisar_por_data(page: Page, data_abertura: str) -> None:
    """Preenche Data Inicial (= data_abertura salva no JSON) e Data Final
    (= hoje) e clica em Pesquisar."""
    data_final = datetime.now().strftime("%d/%m/%Y")

    # AJUSTAR SELETOR: IDs reais dos campos de data
    page.fill("#DataInicial", data_abertura)
    page.fill("#DataFinal", data_final)
    page.click("text=Pesquisar")
    page.wait_for_load_state("networkidle")
    logger.info("Pesquisa realizada: %s até %s", data_abertura, data_final)


def localizar_solicitacao(page: Page, registro: Dict[str, Any]) -> None:
    """Percorre a tabela de resultados e localiza a linha correta.
    Preferência: Ticket. Fallback: Description + Data de abertura."""
    ticket = (registro.get("ticket") or "").strip()
    descricao = (registro.get("description") or "").strip()
    data_abertura = (registro.get("data_abertura") or "").strip()

    # AJUSTAR SELETOR: tabela/linhas de resultado da pesquisa
    linhas = page.locator("table tbody tr")
    total = linhas.count()

    alvo = None

    if ticket:
        for i in range(total):
            candidata = linhas.nth(i)
            if ticket in candidata.inner_text():
                alvo = candidata
                break

    if alvo is None and descricao:
        for i in range(total):
            candidata = linhas.nth(i)
            texto = candidata.inner_text()
            if descricao in texto and (not data_abertura or data_abertura in texto):
                alvo = candidata
                break

    if alvo is None:
        raise RuntimeError(
            "Solicitação não encontrada em 'Solicitações Finalizadas' "
            f"(ticket='{ticket}', description='{descricao}', data='{data_abertura}')."
        )

    alvo.click()
    page.wait_for_load_state("networkidle")
    logger.info("Solicitação localizada e aberta.")


def capturar_ticket(page: Page) -> str:
    """Tenta capturar o número do ticket exibido na tela da solicitação."""
    try:
        # AJUSTAR SELETOR
        return page.locator("#TicketNumber").inner_text().strip()
    except Exception:
        return ""


def abrir_steps_status(page: Page) -> None:
    """Abre a aba 'STEPS STATUS' da solicitação/ticket. Tenta algumas
    variações de seletor (o texto exato pode ter capitalização diferente ou
    ser um link/botão/aba) antes de desistir. AJUSTAR SELETOR se nenhuma
    bater com o Phoenix real - use o screenshot de debug salvo em caso de
    falha para conferir o HTML real da tela."""
    try:
        page.wait_for_load_state("domcontentloaded")
    except Exception:
        pass
    try:
        page.wait_for_load_state("networkidle", timeout=8000)
    except Exception:
        pass

    seletores = [
        "text=STEPS STATUS",
        "text=Steps Status",
        "text=Steps status",
        "a:has-text('STEPS STATUS')",
        "button:has-text('STEPS STATUS')",
        "[role='tab']:has-text('STEPS STATUS')",
    ]

    for selector in seletores:
        locator = page.locator(selector).first
        try:
            locator.wait_for(timeout=6000)
        except PlaywrightTimeoutError:
            continue
        try:
            locator.click()
            try:
                page.wait_for_load_state("domcontentloaded")
            except Exception:
                pass
            logger.info("Aba STEPS STATUS aberta (seletor: %s).", selector)
            return
        except Exception as exc:
            logger.warning("Encontrei '%s' mas não consegui clicar: %s", selector, exc)
            continue

    _diagnostico(page, "steps-status-nao-encontrado")
    _debug_screenshot(page, "steps_status_nao_encontrado")
    raise RuntimeError("Aba 'STEPS STATUS' não encontrada na página da solicitação/ticket.")


def capturar_todos_pns(page: Page) -> List[str]:
    """Localiza a linha do step 'Criação de Part Number' e extrai todos os PNs
    a partir do texto da coluna Resposta."""
    linhas = page.locator("table tbody tr")
    total = linhas.count()

    for i in range(total):
        linha = linhas.nth(i)
        texto = linha.inner_text()

        if "Criação de Part Number" in texto:
            matches = PN_REGEX.findall(texto)
            pns_limpos = list(dict.fromkeys(matches))
            if pns_limpos:
                return pns_limpos

            partes = [p.strip() for p in texto.splitlines() if p.strip()]
            if partes:
                return [partes[-1]]
            return []

    return []


def capturar_pn(page: Page) -> str:
    """Retorna o primeiro PN capturado em STEPS STATUS."""
    pns = capturar_todos_pns(page)
    return pns[0] if pns else ""


def capturar_data_fechamento(page: Page) -> str:
    """Tenta ler a data de finalização exibida na tela; usa hoje como fallback."""
    try:
        # AJUSTAR SELETOR
        texto = page.locator("#ClosedDate").inner_text().strip()
        if texto:
            return texto
    except Exception:
        pass
    return datetime.now().strftime("%d/%m/%Y")


# =====================================================
# ORQUESTRAÇÃO
# =====================================================

def atualizar_pn_phoenix(linha: str) -> None:
    historico = carregar_historico()
    registro = buscar_registro(historico, linha)

    if registro is None:
        logger.warning("Linha %s não encontrada no histórico.", linha)
        return

    status_atual = str(registro.get("status", "")).upper()
    if status_atual != "ON GOING":
        logger.info("Linha %s já está com status '%s'. Nada a fazer.", linha, registro.get("status"))
        return

    config = carregar_config()

    with sync_playwright() as p:
        browser, context, page = _abrir_phoenix(p, config)

        try:
            logger.info("Indo para Solicitações Finalizadas...")
            ir_para_finalizadas(page)

            logger.info("Pesquisando a partir de %s...", registro["data_abertura"])
            pesquisar_por_data(page, registro["data_abertura"])

            logger.info("Localizando a solicitação...")
            localizar_solicitacao(page, registro)

            ticket = capturar_ticket(page) or (registro.get("ticket") or "").strip()
            if ticket:
                registro["ticket"] = ticket
                logger.info("Ticket capturado: %s", ticket)

            logger.info("Abrindo STEPS STATUS...")
            abrir_steps_status(page)

            pns = capturar_todos_pns(page)
            logger.info("PNs capturados (%s): %s", len(pns), pns)

            data_fechamento = capturar_data_fechamento(page)
            logger.info("Data de fechamento: %s", data_fechamento)

            dados_step = {
                "status": STATUS_FINALIZADO,
                "data_fechamento": data_fechamento,
            }
            if ticket:
                atualizar_registros_por_ticket(ticket, pns, dados_step)
            else:
                registro["pn"] = pns[0] if pns else ""
                registro["part_number"] = pns[0] if pns else ""
                registro["status"] = STATUS_FINALIZADO
                registro["data_fechamento"] = data_fechamento
                registro["user"] = registro.get("user") or config.get("user", "")
                salvar_historico(historico)

            logger.info("Histórico atualizado com sucesso.")

        except Exception as exc:
            logger.error("ERRO na automação: %s", exc)
            input("Pressione ENTER para fechar (o registro NÃO foi marcado como finalizado)...")
            return

        input("Concluído. Pressione ENTER para fechar...")
        browser.close()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python atualizar_pn.py <linha>")
        sys.exit(1)

    atualizar_pn_phoenix(sys.argv[1])


# =====================================================
# BOTÃO "ATUALIZAR PN PHOENIX" DO DASHBOARD: busca headless (sem janela
# visível) do Part Number pelo ticket, usando STEPS STATUS
# =====================================================

def capturar_dados_step_pn(page: Page) -> Optional[Dict[str, Any]]:
    """Localiza a linha do step 'Criação de Part Number' dentro de STEPS
    STATUS (já aberta). Retorna None se o step ainda não estiver FINALIZADO
    ou se não encontrar o step. Se FINALIZADO, retorna pns/part_number/status/
    data_pn/hora_pn/analista extraídos do texto da linha."""
    try:
        page.wait_for_selector("table tbody tr", timeout=15000)
    except PlaywrightTimeoutError:
        return None

    linhas = page.locator("table tbody tr")
    total = linhas.count()

    for i in range(total):
        linha = linhas.nth(i)
        texto = linha.inner_text()

        if "Criação de Part Number" not in texto:
            continue

        if "FINALIZADO" not in texto.upper():
            return None

        matches_pn = PN_REGEX.findall(texto)
        pns_limpos = list(dict.fromkeys(matches_pn))
        part_number = pns_limpos[0] if pns_limpos else ""

        data_match = re.search(r"(\d{2}/\d{2}/\d{4})", texto)
        data_pn = data_match.group(1) if data_match else ""

        hora_match = re.search(r"(\d{2}:\d{2}(?::\d{2})?)", texto)
        hora_pn = hora_match.group(1) if hora_match else ""

        partes = [p.strip() for p in texto.splitlines() if p.strip()]
        analista = ""
        for candidato in partes:
            if candidato.lower() in _BOILERPLATE_RESPOSTA:
                continue
            if candidato in pns_limpos or candidato in (data_pn, hora_pn):
                continue
            if "Criação de Part Number" in candidato or "FINALIZADO" in candidato.upper():
                continue
            if _NOME_ANALISTA_REGEX.match(candidato):
                analista = candidato
                break

        return {
            "pns": pns_limpos,
            "part_number": part_number,
            "status": STATUS_STEP_PN_CRIADO,
            "data_pn": data_pn,
            "hora_pn": hora_pn,
            "analista": analista,
        }

    return None


def _indice_coluna_tabela(page: Page, texto_cabecalho: str) -> Optional[int]:
    cabecalhos = page.locator("table thead th")
    total = cabecalhos.count()
    for i in range(total):
        try:
            texto = cabecalhos.nth(i).inner_text().strip().lower()
        except Exception:
            continue
        if texto == texto_cabecalho.strip().lower():
            return i
    return None


def buscar_pn_por_ticket(ticket: str) -> Dict[str, Any]:
    ticket_normalizado = (ticket or "").strip()
    if not ticket_normalizado:
        return {"ok": False, "mensagem": "Este registro não possui ticket.", "part_number": ""}

    registros = encontrar_todos_por_ticket(ticket_normalizado)
    if not registros:
        reg_unico = encontrar_por_ticket(ticket_normalizado)
        if reg_unico is None:
            return {"ok": False, "mensagem": "Ticket não encontrado no histórico.", "part_number": ""}
        registros = [reg_unico]

    pns_existentes = [str(r.get("pn") or r.get("part_number") or "").strip() for r in registros]
    if all(bool(p) for p in pns_existentes):
        return {
            "ok": True,
            "mensagem": f"Part Number(s) já registrados: {', '.join(pns_existentes)}",
            "part_number": pns_existentes[0],
        }

    config = carregar_config()

    try:
        with sync_playwright() as p:
            browser, context, page = _abrir_phoenix(p, config, headless=True)
            try:
                if not _ir_para_minhas_solicitacoes(page):
                    return {
                        "ok": False,
                        "mensagem": "Não encontrei 'Minhas Solicitações' no Phoenix.",
                        "part_number": "",
                    }

                try:
                    page.wait_for_selector("table tbody tr", timeout=20000)
                except PlaywrightTimeoutError:
                    return {
                        "ok": False,
                        "mensagem": "Tabela de 'Minhas Solicitações' não carregou a tempo.",
                        "part_number": "",
                    }

                linhas = page.locator("table tbody tr")
                total = linhas.count()
                alvo = None
                for i in range(total):
                    candidata = linhas.nth(i)
                    if ticket_normalizado in candidata.inner_text():
                        alvo = candidata
                        break

                if alvo is None:
                    return {
                        "ok": False,
                        "mensagem": f"Ticket {ticket_normalizado} não encontrado em 'Minhas Solicitações'.",
                        "part_number": "",
                    }

                indice_steps = _indice_coluna_tabela(page, "Steps")
                if indice_steps is not None:
                    celula_steps = alvo.locator("td").nth(indice_steps)
                    candidatos_clicaveis = celula_steps.locator("a, button, [role='button'], img, i, svg")
                    elemento_clicavel = (
                        candidatos_clicaveis.first if candidatos_clicaveis.count() > 0 else celula_steps
                    )
                else:
                    elemento_clicavel = alvo

                pagina_detalhe = page
                try:
                    with context.expect_page(timeout=6000) as info_nova_pagina:
                        elemento_clicavel.click()
                    pagina_detalhe = info_nova_pagina.value
                    try:
                        pagina_detalhe.wait_for_load_state("domcontentloaded")
                    except Exception:
                        pass
                    logger.info("Steps do ticket %s abriram em uma nova aba.", ticket_normalizado)
                except PlaywrightTimeoutError:
                    try:
                        page.wait_for_load_state("domcontentloaded")
                    except Exception:
                        pass
                    logger.info("Steps do ticket %s abertos.", ticket_normalizado)

                dados_step = capturar_dados_step_pn(pagina_detalhe)
                if dados_step is None:
                    try:
                        abrir_steps_status(pagina_detalhe)
                        dados_step = capturar_dados_step_pn(pagina_detalhe)
                    except Exception:
                        pass

                if not dados_step:
                    return {"ok": False, "mensagem": "PART NUMBER AINDA NÃO FOI GERADO", "part_number": ""}

                pns_encontrados = dados_step.get("pns") or ([dados_step["part_number"]] if dados_step.get("part_number") else [])
                if not pns_encontrados:
                    return {"ok": False, "mensagem": "PART NUMBER AINDA NÃO FOI GERADO", "part_number": ""}

                atualizar_registros_por_ticket(ticket_normalizado, pns_encontrados, dados_step)

                msg_sucesso = f"Part Number(s) capturado(s): {', '.join(pns_encontrados[:len(registros)])}"
                logger.info(
                    "PNs atualizados para o ticket %s: %s", ticket_normalizado, pns_encontrados
                )

                return {
                    "ok": True,
                    "mensagem": msg_sucesso,
                    "part_number": pns_encontrados[0],
                }
            finally:
                try:
                    browser.close()
                except Exception:
                    pass
    except Exception:
        logger.exception("Falha ao buscar PN para o ticket %s.", ticket_normalizado)
        return {
            "ok": False,
            "mensagem": "Falha ao buscar o Part Number. Veja phoenix_tool.log para detalhes.",
            "part_number": "",
        }
