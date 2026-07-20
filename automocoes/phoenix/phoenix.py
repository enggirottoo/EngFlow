"""Automação Phoenix (Portal + Phoenix).

Duas entradas são usadas pelo Phoenix Tool (via `subprocess`, veja
`services/process_runner.py`):

    python phoenix.py home   -> abrir_home_phoenix()
    python phoenix.py        -> nova_solicitacao_phoenix()

O fluxo de login (`_preencher_login`) e a lógica de detectar se o clique em
um link/tile abriu uma nova aba (`_abrir_phoenix`) são compartilhados com
`pegasus.py`; qualquer ajuste de seletor de login deve ser espelhado lá
também caso o Portal mude.
"""

from __future__ import annotations

import os
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog
from typing import Any, Dict, Optional, Tuple

from playwright.sync_api import Browser, BrowserContext, Page, Playwright
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = Path(__file__).resolve().parents[2]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from automocoes.planilha.planilha import enviar_para_planilha  # noqa: E402
from services.logging_config import configurar_logger  # noqa: E402
from services.storage import carregar_login, criar_registro_descricao  # noqa: E402

logger = configurar_logger("phoenix")


def carregar_config() -> Dict[str, Any]:
    """Usuário/senha vêm do armazenamento seguro central (`services.storage`),
    não mais lidos diretamente de `config_tool.json`."""
    return carregar_login()


def _esperar_elemento(page: Page, selector: str, timeout: int = 10000) -> bool:
    try:
        page.wait_for_selector(selector, timeout=timeout)
        return True
    except PlaywrightTimeoutError:
        return False


def _debug_screenshot(page: Page, nome: str) -> None:
    """Salva um print da tela para ajudar a diagnosticar problema de seletor/login."""
    try:
        pasta = os.path.join(BASE_DIR, "debug")
        os.makedirs(pasta, exist_ok=True)
        caminho = os.path.join(pasta, f"{nome}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
        page.screenshot(path=caminho, full_page=True)
        logger.info("Screenshot salvo em: %s", caminho)
    except Exception as exc:
        logger.warning("Não foi possível salvar screenshot de debug: %s", exc)


def _diagnostico(page: Page, label: str) -> None:
    """Loga URL e título atuais da página para ajudar a diagnosticar sem precisar de screenshot."""
    try:
        logger.info("[%s] URL atual: %s", label, page.url)
        logger.info("[%s] Título: %s", label, page.title())
    except Exception as exc:
        logger.warning("[%s] Não consegui ler URL/título: %s", label, exc)


def _preencher_login(page: Page, config: Dict[str, Any], timeout: int = 15000) -> bool:
    if not _esperar_elemento(page, "#FlexUser", timeout=timeout):
        logger.info(
            "Campo de login (#FlexUser) não apareceu em %sms. "
            "Ou a página já está logada, ou o seletor mudou.",
            timeout,
        )
        _diagnostico(page, "sem-campo-login")
        _debug_screenshot(page, "login_campo_nao_encontrado")
        return False

    if not (config.get("user") or "").strip() or not (config.get("password") or "").strip():
        logger.warning(
            "Usuário/senha configurados estão vazios. Nada será digitado nos campos de login."
        )

    url_antes = page.url
    page.fill("#FlexUser", config.get("user", ""))
    page.fill("#Password", config.get("password", ""))

    clicou = False
    for selector in [
        'xpath=//*[@id="formLogin"]/button',
        'xpath=//*[@id="formLogin"]/button/span',
        'button[type="submit"]',
        'text=Entrar',
        'text=Login',
    ]:
        if _esperar_elemento(page, selector, timeout=3000):
            try:
                page.click(selector)
                clicou = True
                break
            except Exception:
                continue

    if not clicou:
        logger.info("Campos preenchidos, mas não achei o botão de login para clicar. Tentando ENTER.")
        page.keyboard.press("Enter")

    # Espera a navegação acontecer (URL mudar) em vez de esperar a rede ficar
    # ociosa - páginas corporativas costumam ter polling/telemetria contínuos
    # em background, então "networkidle" quase nunca resolve rápido e deixava
    # o login extremamente lento sem necessidade.
    try:
        page.wait_for_url(lambda url: url != url_antes, timeout=8000)
    except PlaywrightTimeoutError:
        pass
    try:
        page.wait_for_load_state("domcontentloaded", timeout=5000)
    except Exception:
        pass

    # Checagem instantânea (não espera timeout inteiro) se o campo de
    # usuário ainda está na tela - indicativo de login rejeitado.
    try:
        ainda_com_login = page.locator("#FlexUser").count() > 0
    except Exception:
        ainda_com_login = False

    if ainda_com_login:
        logger.warning(
            "Depois de enviar o formulário, o campo de usuário ainda está na tela. "
            "O login provavelmente FALHOU (confira o usuário/senha salvos)."
        )
        _diagnostico(page, "login-falhou")
        _debug_screenshot(page, "login_falhou")
        return False

    logger.info("Login enviado e formulário de login não aparece mais na tela.")
    _diagnostico(page, "login-ok")
    return True


def _abrir_phoenix(playwright: Playwright, config: Dict[str, Any]) -> Tuple[Browser, BrowserContext, Page]:
    """Login no Portal + login no Phoenix. Retorna (browser, context, page) já logado no Phoenix."""

    browser = playwright.chromium.launch(
        channel="msedge",
        headless=False,
        args=["--start-maximized"]
    )

    context = browser.new_context(viewport=None)
    page = context.new_page()

    # PORTAL
    page.goto("http://portal.empresa-exemplo.com/", wait_until="domcontentloaded")
    if _esperar_elemento(page, "#FlexUser", timeout=20000):
        _preencher_login(page, config, timeout=20000)
    else:
        logger.info("Portal: formulário de login não apareceu (provavelmente já autenticado).")
    _diagnostico(page, "portal")
    logger.info("Portal OK")

    # ABRIR PHOENIX
    url_antes_click = page.url
    seletor_disponivel: Optional[str] = None
    for selector in [
        'xpath=//*[@id="solution"]/div/div/div[2]/div/div[3]/a',
        'a[href*="Phoenix"]',
        'text=Phoenix',
        'text=PHOENIX',
    ]:
        if _esperar_elemento(page, selector, timeout=5000):
            seletor_disponivel = selector
            break

    abriu_phoenix = False
    nova_pagina = None
    if seletor_disponivel:
        try:
            # O clique costuma abrir o Phoenix numa aba/janela nova. Usar
            # expect_page evita a corrida de checar context.pages depois do
            # clique (a aba nova pode não existir ainda no instante checado).
            with context.expect_page(timeout=6000) as info_nova_pagina:
                page.click(seletor_disponivel)
            nova_pagina = info_nova_pagina.value
            abriu_phoenix = True
        except PlaywrightTimeoutError:
            # Não abriu aba nova: deve ter navegado na própria aba.
            abriu_phoenix = True
        except Exception as exc:
            logger.error("Erro ao clicar no link do Phoenix: %s", exc)
    else:
        logger.warning(
            "Não encontrei nenhum link/botão para abrir o Phoenix na Home do Portal. "
            "O fluxo pode continuar preso na tela do Portal."
        )
        _debug_screenshot(page, "phoenix_link_nao_encontrado")

    if nova_pagina is not None:
        try:
            nova_pagina.wait_for_load_state("domcontentloaded")
        except Exception:
            pass
        page = nova_pagina
        logger.info("Phoenix abriu em uma nova aba.")
    else:
        try:
            page.wait_for_load_state("domcontentloaded")
        except Exception:
            pass
        if abriu_phoenix and page.url == url_antes_click:
            logger.warning(
                "Cliquei no link do Phoenix, mas a URL não mudou (%s). O clique "
                "provavelmente não teve efeito - verifique se o seletor ainda é o correto.",
                page.url,
            )
            _debug_screenshot(page, "phoenix_url_nao_mudou")

    _diagnostico(page, "phoenix-aberto")
    logger.info("Phoenix OK")

    # LOGIN PHOENIX
    try:
        if _esperar_elemento(page, "#FlexUser", timeout=10000):
            _preencher_login(page, config, timeout=10000)
        else:
            logger.info("Phoenix: campo de login não apareceu (já estava logado ou seletor mudou).")
            _diagnostico(page, "phoenix-sem-campo-login")
    except Exception as exc:
        logger.error("Erro ao tentar logar no Phoenix: %s", exc)
        _diagnostico(page, "phoenix-erro-login")

    return browser, context, page


# =====================================================
# FUNÇÃO 1: só abrir a home do Phoenix (botão "Home Phoenix")
# =====================================================

def abrir_home_phoenix() -> None:
    config = carregar_config()

    with sync_playwright() as p:
        browser, context, page = _abrir_phoenix(p, config)

        logger.info("Home Phoenix aberta.")
        input("Pressione ENTER para fechar...")

        browser.close()


# =====================================================
# FUNÇÃO 2: fluxo completo de solicitação P0032 (botão "Nova Solicitação")
# =====================================================

def nova_solicitacao_phoenix() -> None:
    config = carregar_config()

    with sync_playwright() as p:
        browser, context, page = _abrir_phoenix(p, config)

        # P0032
        logger.info("Abrindo P0032...")
        page.fill("#TaskName", "32")
        page.click("text=P0032")
        page.click('xpath=//*[@id="formId"]/div/div[2]/button')
        page.wait_for_load_state("domcontentloaded")
        logger.info("P0032 OK")

        # ESCOLHER TEMPLATE
        logger.info("Abrindo janela para selecionar o template (.xlsx)...")
        root = tk.Tk()
        root.withdraw()
        # Sem isso a janela de seleção de arquivo pode abrir atrás do
        # navegador (que acabou de receber foco) e parecer que "não carrega".
        root.attributes("-topmost", True)
        root.update()

        arquivo = filedialog.askopenfilename(
            parent=root,
            title="Selecione o Template",
            filetypes=[("Excel", "*.xlsx")]
        )

        root.destroy()

        if not arquivo:
            logger.info("Nenhum arquivo selecionado.")
            browser.close()
            return

        logger.info("Arquivo selecionado: %s", arquivo)

        # UPLOAD TEMPLATE
        page.set_input_files("#ExcelFile", arquivo)
        logger.info("Template carregado")

        # LUPA
        page.click('xpath=//*[@id="uploadForm"]/div/div[1]/div/div/div/button')
        logger.info("Lupa clicada")

        try:
            descricao = page.locator('[id$="__Description"]').input_value()
            logger.info("DESCRIPTION: %s", descricao)
        except Exception as exc:
            logger.error("Erro ao capturar Description: %s", exc)
            descricao = ""

        # AREA (só existe para alguns usuários)
        try:
            if page.locator("#areaId").count() > 0:
                page.select_option("#areaId", label="ENGENHARIA DE TESTE (JAGUARIÚNA)")
                logger.info("Area OK - Engenharia de Teste")
                page.wait_for_timeout(1500)
        except Exception as exc:
            logger.error("Erro Area: %s", exc)

        # PRIORITY
        try:
            page.wait_for_selector("#Priority", timeout=15000)
            page.select_option("#Priority", value="URGENT")
            logger.info("Priority OK")
        except Exception as exc:
            logger.error("Erro Priority: %s", exc)

        # COMPANY
        try:
            page.select_option("#CompanyId", value="65bb402c-cb05-441f-9a4f-df9eb51c7ae5")
            logger.info("Company OK")
        except Exception as exc:
            logger.error("Erro Company: %s", exc)

        # ITEM TYPE
        try:
            page.locator('select[id$="__ItemType"]').select_option(value="D")
            logger.info("Item Type OK")
        except Exception as exc:
            logger.error("Erro Item Type: %s", exc)

        # ROHS
        try:
            page.locator('select[id$="__Rohs"]').select_option(value="na")
            logger.info("ROHS OK")
        except Exception as exc:
            logger.error("Erro ROHS: %s", exc)

        # ORIGIN
        try:
            page.locator('select[id$="__Origin"]').select_option(value="IMPORTADO")
            logger.info("Origin OK")
        except Exception as exc:
            logger.error("Erro Origin: %s", exc)

        logger.info("=" * 50)
        logger.info("PREENCHER MANUALMENTE: Assunto Principal, Manufacturer")
        logger.info("=" * 50)
        logger.info("Confirme a solicitação no Phoenix e pressione ENTER para registrar no histórico.")

        input("Pressione ENTER para registrar e finalizar...")

        registro = criar_registro_descricao(descricao or "", config.get("user", ""))
        logger.info("Histórico salvo. Linha criada: %s", registro["linha"])

        enviar_para_planilha(descricao or "", linha=registro.get("linha"), usuario=config.get("user", ""))
        logger.info("Payload enviado para a planilha.")

        browser.close()


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "home":
        abrir_home_phoenix()
    else:
        nova_solicitacao_phoenix()
