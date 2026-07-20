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

from phoenix import carregar_config, _abrir_phoenix  # noqa: E402  (reaproveita login já pronto)
from services.logging_config import configurar_logger  # noqa: E402
from services.storage import carregar_historico, salvar_historico  # noqa: E402

logger = configurar_logger("atualizar_pn")

# Rótulo de status usado quando a solicitação é finalizada.
STATUS_FINALIZADO = "OK"

# Heurística de formato de PN: 3 dígitos - 3 letras - alfanumérico
# (ex: 361-FXB-D202606183J). Ajuste se o padrão real variar.
PN_REGEX = re.compile(r"([0-9]{3}-[A-Z]{3}-[A-Za-z0-9]+)")


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
    # AJUSTAR SELETOR: aba "STEPS STATUS"
    page.click("text=STEPS STATUS")
    page.wait_for_load_state("networkidle")
    logger.info("Aba STEPS STATUS aberta.")


def capturar_pn(page: Page) -> str:
    """Localiza a linha do step 'Criação de Part Number' e extrai o PN
    a partir do texto da coluna Resposta."""
    # AJUSTAR SELETOR: tabela de steps dentro de STEPS STATUS
    linhas = page.locator("table tbody tr")
    total = linhas.count()

    for i in range(total):
        linha = linhas.nth(i)
        texto = linha.inner_text()

        if "Criação de Part Number" in texto:
            # heurística: o PN normalmente é a última linha não vazia
            # da resposta (ex: "Hi Team\nParts\n361-FXB-D202606183J")
            partes = [p.strip() for p in texto.splitlines() if p.strip()]
            candidato = partes[-1] if partes else ""

            match = PN_REGEX.search(candidato) or PN_REGEX.search(texto)
            if match:
                return match.group(1)

            # não bateu com o padrão esperado - devolve mesmo assim,
            # para você conseguir ver no histórico o que foi capturado
            return candidato

    raise RuntimeError("Step 'Criação de Part Number' não encontrado em STEPS STATUS.")


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

            ticket = capturar_ticket(page)
            if ticket:
                registro["ticket"] = ticket
                logger.info("Ticket capturado: %s", ticket)

            logger.info("Abrindo STEPS STATUS...")
            abrir_steps_status(page)

            pn = capturar_pn(page)
            logger.info("PN capturado: %s", pn)

            data_fechamento = capturar_data_fechamento(page)
            logger.info("Data de fechamento: %s", data_fechamento)

            registro["pn"] = pn or registro.get("pn", "")
            registro["data_fechamento"] = data_fechamento or registro.get("data_fechamento", "")
            registro["status"] = STATUS_FINALIZADO
            registro["ticket"] = registro.get("ticket", "") or ticket
            registro["user"] = registro.get("user") or config.get("user", "")
            if not registro.get("data_fechamento"):
                registro["data_fechamento"] = datetime.now().strftime("%d/%m/%Y")

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
