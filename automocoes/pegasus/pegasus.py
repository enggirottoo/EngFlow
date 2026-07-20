"""Automação Pegasus (Portal + Pegasus).

Duas entradas são usadas pelo Phoenix Tool (via `subprocess`):

    python pegasus.py home   -> abrir_home_pegasus()
    python pegasus.py        -> nova_solicitacao_pegasus()

O fluxo de login (`_preencher_login`) e a detecção de nova aba
(`_abrir_com_fallback`) espelham a implementação usada em
`automocoes/phoenix/phoenix.py`; ajustes de seletor de login devem ser
replicados nos dois arquivos.
"""

from __future__ import annotations

import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from playwright.sync_api import BrowserContext, Page, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = Path(__file__).resolve().parents[2]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from services.logging_config import configurar_logger  # noqa: E402
from services.storage import carregar_login  # noqa: E402

logger = configurar_logger("pegasus")


def carregar_config() -> Dict[str, Any]:
    """Usuário/senha vêm do armazenamento seguro central (`services.storage`)."""
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


def _preencher_login(page: Page, config: Dict[str, Any], timeout: int = 20000) -> bool:
    """Preenche usuário/senha se o formulário de login aparecer. Retorna True se preencheu."""
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

    page.fill("#FlexUser", config.get("user", ""))
    page.fill("#Password", config.get("password", ""))
    url_antes = page.url

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


def _abrir_com_fallback(
    page: Page,
    context: BrowserContext,
    seletores: List[str],
    aba_timeout: int = 6000,
    elemento_timeout: int = 5000,
) -> Tuple[Page, bool]:
    """Tenta clicar no primeiro seletor disponível. O clique de tiles do Portal
    costuma abrir uma aba/janela nova - usamos expect_page para capturá-la
    sem depender de checar context.pages depois (que sofre corrida, pois a
    aba pode não existir ainda no instante checado). Se nenhuma aba nova
    abrir, assumimos navegação na própria página.

    Retorna (pagina_atual, clicou: bool).
    """
    seletor_disponivel: Optional[str] = None
    for selector in seletores:
        if _esperar_elemento(page, selector, timeout=elemento_timeout):
            seletor_disponivel = selector
            break

    if not seletor_disponivel:
        return page, False

    url_antes = page.url
    nova_pagina = None
    try:
        with context.expect_page(timeout=aba_timeout) as info_nova_pagina:
            page.click(seletor_disponivel)
        nova_pagina = info_nova_pagina.value
    except PlaywrightTimeoutError:
        pass
    except Exception as exc:
        logger.error("Erro ao clicar no link: %s", exc)
        return page, False

    if nova_pagina is not None:
        try:
            nova_pagina.wait_for_load_state("domcontentloaded")
        except Exception:
            pass
        logger.info("Abriu em uma nova aba.")
        return nova_pagina, True

    try:
        page.wait_for_load_state("domcontentloaded")
    except Exception:
        pass
    if page.url == url_antes:
        logger.warning(
            "Cliquei no link, mas a URL não mudou (%s). O clique provavelmente não "
            "teve efeito - verifique se o seletor ainda é o correto.",
            page.url,
        )
    return page, True


def _login_portal(page: Page, config: Dict[str, Any]) -> None:
    """Login no Engineering Portal. Igual nos dois fluxos."""
    page.goto("http://portal.empresa-exemplo.com/", wait_until="domcontentloaded")
    if _preencher_login(page, config, timeout=20000):
        page.wait_for_load_state("domcontentloaded")


# =====================================================
# FUNÇÃO 1: só abrir a home do Pegasus (botão "Home Pegasus")
# =====================================================

def abrir_home_pegasus() -> None:
    config = carregar_config()

    with sync_playwright() as p:

        browser = p.chromium.launch(
            channel="msedge",
            headless=False,
            args=["--start-maximized"]
        )

        context = browser.new_context(viewport=None)
        page = context.new_page()

        # PORTAL
        _login_portal(page, config)
        _diagnostico(page, "portal")
        logger.info("Portal OK")

        # ABRIR PEGASUS
        page, abriu_pegasus = _abrir_com_fallback(
            page,
            context,
            [
                'xpath=//*[@id="solution"]/div/div/div[1]/div/div[3]/a',
                'a[href*="Pegasus"]',
                'text=Pegasus',
                'text=PEGASUS',
            ],
        )
        if not abriu_pegasus:
            logger.warning(
                "Não encontrei nenhum link/botão para abrir o Pegasus na Home do Portal. "
                "O fluxo pode continuar preso na tela do Portal."
            )
            _debug_screenshot(page, "pegasus_link_nao_encontrado")

        _diagnostico(page, "pegasus-aberto")
        logger.info("Pegasus OK")

        # LOGIN PEGASUS
        if not _preencher_login(page, config, timeout=10000):
            logger.info("Pegasus: campo de login não apareceu (já estava logado ou seletor mudou).")

        logger.info("Home Pegasus aberta.")

        input("Pressione ENTER para fechar...")


# =====================================================
# FUNÇÃO 2: fluxo completo de nova solicitação (botão "Nova solicitação")
# =====================================================

def nova_solicitacao_pegasus() -> None:
    config = carregar_config()

    with sync_playwright() as p:

        browser = p.chromium.launch(
            channel="msedge",
            headless=False,
            args=["--window-size=1360,768"]
        )

        context = browser.new_context(
            viewport={"width": 1360, "height": 768}
        )

        page = context.new_page()

        logger.info("Abrindo Engineering Portal...")

        # PORTAL
        _login_portal(page, config)
        _diagnostico(page, "portal")
        logger.info("Portal logado.")

        # ABRIR PEGASUS
        page, abriu_pegasus = _abrir_com_fallback(
            page,
            context,
            [
                'xpath=//*[@id="solution"]/div/div/div[1]/div/div[3]/a',
                'a[href*="Pegasus"]',
                'text=Pegasus',
                'text=PEGASUS',
            ],
        )
        if not abriu_pegasus:
            logger.warning(
                "Não encontrei nenhum link/botão para abrir o Pegasus na Home do Portal. "
                "O fluxo pode continuar preso na tela do Portal."
            )
            _debug_screenshot(page, "pegasus_link_nao_encontrado")
        page.wait_for_timeout(5000)

        page.set_viewport_size({"width": 1360, "height": 768})
        page.bring_to_front()

        logger.info("Viewport: %s", page.viewport_size)

        page.evaluate("""
window.moveTo(0,0);
window.resizeTo(screen.availWidth, screen.availHeight);
""")

        page.wait_for_timeout(2000)

        logger.info("Largura Janela: %s", page.evaluate("window.innerWidth"))
        logger.info("Altura Janela: %s", page.evaluate("window.innerHeight"))
        logger.info("Largura Tela: %s", page.evaluate("screen.width"))
        logger.info("Altura Tela: %s", page.evaluate("screen.height"))

        try:
            page.wait_for_load_state("networkidle", timeout=8000)
        except Exception:
            pass
        _diagnostico(page, "pegasus-aberto")
        logger.info("Pegasus aberto.")

        # LOGIN PEGASUS (SE PEDIR)
        if not _preencher_login(page, config, timeout=15000):
            logger.info("Pegasus: campo de login não apareceu (já estava logado ou seletor mudou).")

        logger.info("Pegasus logado.")

        # NAVEGAÇÃO
        page.click('xpath=/html/body/div[4]/section[2]/section/div[1]/div[1]/div/div/a[1]')
        page.wait_for_load_state("domcontentloaded")

        page.click('xpath=/html/body/div[4]/section[2]/section/div/div[2]/div/div[1]/a')
        page.wait_for_load_state("domcontentloaded")

        page.click('xpath=/html/body/div[4]/section[2]/section/div/div[2]/div/div[1]/a')
        page.wait_for_load_state("domcontentloaded")

        logger.info("Site Jaguariúna.")

        # PREENCHIMENTO INICIAL
        try:
            page.select_option("#companyId", index=1)
            page.select_option("#NonProductiveItemRequest_ItemStatusId", index=1)
            page.select_option("#NonProductiveItemRequest_ItemTypeId", index=2)
            page.select_option("#NonProductiveItemRequest_ItemTransactionId", index=2)
            page.select_option("#siteTaxClassifierGroupId", index=1)
            page.select_option("#taxClassifierGroupId", index=1)
            page.select_option("#NonProductiveItemRequest_RequesterGroupId", index=1)
        except Exception as exc:
            logger.error("Erro ao preencher formulário inicial: %s", exc)

        logger.info("Formulário preparado.")

        input("Pressione ENTER para encerrar...")

        browser.close()


# =====================================================
# Execução via linha de comando (chamado pelo main.py com subprocess)
#
# python pegasus.py home  -> abrir_home_pegasus()
# python pegasus.py       -> nova_solicitacao_pegasus()
# =====================================================

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "home":
        abrir_home_pegasus()
    else:
        nova_solicitacao_pegasus()
