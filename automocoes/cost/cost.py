"""Automação Cost Request.

Chamada pelo main.py via subprocess:

    python cost.py

Abre o Engineering Portal, faz login, navega até "Other Tools > Cost
Request" e prepara o modal de nova solicitação. Os campos "Part Number" e
"Quotation Price" precisam ser preenchidos manualmente pelo usuário antes de
confirmar (ver aviso impresso no console).
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Optional

from playwright.sync_api import Page, sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = Path(__file__).resolve().parents[2]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from services.logging_config import configurar_logger  # noqa: E402
from services.storage import carregar_login  # noqa: E402

logger = configurar_logger("cost")


def navigate_and_wait(page: Page, url: str, description: str) -> None:
    page.goto(url, wait_until="networkidle")
    logger.info("%s OK", description)


def click_element(page: Page, selector: str, description: Optional[str] = None) -> None:
    page.click(selector)
    if description:
        logger.info("%s clicado", description)


def select_option(page: Page, selector: str, label: str, description: Optional[str] = None) -> None:
    page.select_option(selector, label=label)
    if description:
        logger.info("%s OK", description)


def login(page: Page, user: str, password: str) -> None:
    navigate_and_wait(page, "http://portal.empresa-exemplo.com/", "Portal")
    page.fill("#FlexUser", user)
    page.fill("#Password", password)
    click_element(page, 'xpath=//*[@id="formLogin"]/button')
    page.wait_for_load_state("networkidle")


def nova_solicitacao_cost() -> None:
    config = carregar_login()
    if not config.get("user") or not config.get("password"):
        logger.warning("Usuário/senha não configurados. Faça login no Phoenix Tool primeiro.")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            channel="msedge",
            headless=False,
            args=["--start-maximized"]
        )
        context = browser.new_context(viewport=None)
        page = context.new_page()

        try:
            login(page, config.get("user", ""), config.get("password", ""))

            navigate_and_wait(page, "http://portal.empresa-exemplo.com/br/site01/OtherTools", "Other Tools")
            navigate_and_wait(
                page,
                "http://portal.empresa-exemplo.com/br/site01/OtherTools/CostRequest/Index",
                "Cost Request",
            )

            page.wait_for_timeout(3000)
            click_element(page, "button.createCostRequest", "Open Request")

            page.wait_for_selector("#siteId", timeout=15000)
            logger.info("Modal OK")

            select_option(page, "#siteId", "SITE01", "Site")
            select_option(page, "#companyId", "361 - SITE01", "Company")
            select_option(page, "#CostRequest_Coin", "USD", "Coin")
            select_option(page, "#CostRequest_PartNumberOrigin", "IMPORTADO", "Origin")
            select_option(page, "#CostRequest_Type", "OPS", "Type")

            logger.info("=" * 50)
            logger.info("PREENCHER MANUALMENTE: Part Number, Quotation Price")
            logger.info("=" * 50)

            input("Pressione ENTER para finalizar...")
        except Exception as exc:
            logger.error("Erro na automação de Cost Request: %s", exc)
        finally:
            browser.close()


if __name__ == "__main__":
    nova_solicitacao_cost()
