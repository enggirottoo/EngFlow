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
import re
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
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
from services.storage import (  # noqa: E402
    carregar_login,
    criar_registro_descricao,
    existe_ticket,
)

# Grava também em arquivo (phoenix_tool.log): a automação roda num console
# separado (CREATE_NEW_CONSOLE) que pode ser fechado sem que o usuário veja
# o erro - sem log em arquivo, uma falha nessa janela some para sempre e o
# registro no histórico/dashboard nunca é diagnosticado.
logger = configurar_logger("phoenix", arquivo_log="phoenix_tool.log")


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


def _abrir_phoenix(
    playwright: Playwright, config: Dict[str, Any], headless: bool = False
) -> Tuple[Browser, BrowserContext, Page]:
    """Login no Portal + login no Phoenix. Retorna (browser, context, page) já logado no Phoenix.

    `headless=True` é usado por automações que rodam em segundo plano sem
    mostrar nenhuma janela ao usuário (ex.: atualizar_pn.buscar_pn_por_ticket)."""

    browser = playwright.chromium.launch(
        channel="msedge",
        headless=headless,
        args=["--start-maximized"] if not headless else [],
    )

    context = browser.new_context(viewport=None if not headless else {"width": 1366, "height": 768})
    page = context.new_page()

    # PORTAL
    page.goto("http://engenharia.br.flex.com/", wait_until="domcontentloaded")
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

    # Produto/Solicitante/Requisitante não são mais pedidos aqui: o usuário
    # preenche esses campos depois, editando o registro já criado no
    # dashboard (botão "Editar"). Isso evita travar o início da automação
    # esperando um popup da ferramenta.
    with sync_playwright() as p:
        browser, context, page = _abrir_phoenix(p, config)
        try:
            _fluxo_nova_solicitacao(page, browser, config)
        except Exception:
            logger.exception(
                "Automação do Phoenix falhou antes de registrar no histórico/dashboard."
            )
            _debug_screenshot(page, "phoenix_falha_inesperada")
            try:
                messagebox.showerror(
                    "Erro na automação do Phoenix",
                    "A automação falhou e a solicitação NÃO foi registrada no dashboard.\n"
                    "Veja phoenix_tool.log para detalhes.",
                )
            except Exception:
                pass
        finally:
            try:
                browser.close()
            except Exception:
                pass


def _fluxo_nova_solicitacao(
    page: Page,
    browser: Browser,
    config: Dict[str, Any],
) -> None:
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
        return

    logger.info("Arquivo selecionado: %s", arquivo)

    # UPLOAD TEMPLATE
    page.set_input_files("#ExcelFile", arquivo)
    logger.info("Template carregado")

    # LUPA
    page.click('xpath=//*[@id="uploadForm"]/div/div[1]/div/div/div/button')
    logger.info("Lupa clicada")

    itens_form: List[Dict[str, str]] = []
    try:
        inputs_desc = page.locator('[id$="__Description"]')
        inputs_mpn = page.locator('[id$="__Mpn"]')
        total_inputs = inputs_desc.count()
        logger.info("TOTAL DE PRODUTOS NO TEMPLATE: %s", total_inputs)
        for idx in range(total_inputs):
            desc_val = inputs_desc.nth(idx).input_value().strip()
            mpn_val = inputs_mpn.nth(idx).input_value().strip() if inputs_mpn.count() > idx else ""
            if desc_val:
                itens_form.append({"description": desc_val, "mpn": mpn_val})
        logger.info("ITENS CAPTURADOS NO FORMULARIO (%s): %s", len(itens_form), itens_form)
    except Exception as exc:
        logger.error("Erro ao capturar lista de itens do formulário: %s", exc)

    if not itens_form:
        itens_form = [{"description": "", "mpn": ""}]

    descricao = itens_form[0]["description"]

    # AREA (só existe para alguns usuários)
    try:
        if page.locator("#areaId").count() > 0:
            page.select_option("#areaId", value="8eb210b1-bf04-4fbf-848d-5f3aeac9a0b8")
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
        item_types = page.locator('select[id$="__ItemType"]')

        for i in range(item_types.count()):
            item_types.nth(i).select_option(value="D")

        logger.info("Item Type OK")
    except Exception as exc:
        logger.error("Erro Item Type: %s", exc)


    # ROHS
    try:
        rohs = page.locator('select[id$="__Rohs"]')

        for i in range(rohs.count()):
            rohs.nth(i).select_option(value="na")

        logger.info("ROHS OK")
    except Exception as exc:
        logger.error("Erro ROHS: %s", exc)


    
    # ORIGIN
    try:
        origins = page.locator('select[id$="__Origin"]')

        for i in range(origins.count()):
            origins.nth(i).select_option(value="IMPORTADO")

        logger.info("Origin OK")
    except Exception as exc:
        logger.error("Erro Origin: %s", exc)


    logger.info("=" * 50)
    logger.info("PREENCHER MANUALMENTE: Assunto Principal, Manufacturer")
    logger.info("=" * 50)
    logger.info(
        "Quando terminar, clique no botão 'Confirmar Solicitação' dentro do Phoenix. "
        "A automação vai aguardar a solicitação ser criada e capturar o ticket automaticamente."
    )

    # Em vez de um popup pedindo confirmação manual da ferramenta, aguardamos
    # o envio REAL do formulário no Phoenix (botão "Confirmar Solicitação"),
    # detectado pela mudança de URL. O navegador continua aberto - nada é
    # fechado até o ticket ser capturado.
    url_formulario = page.url
    try:
        page.wait_for_url(lambda url: url != url_formulario, timeout=600000)
    except PlaywrightTimeoutError:
        logger.warning(
            "Não detectei o envio do formulário (URL não mudou a tempo). A solicitação "
            "NÃO foi registrada no histórico/dashboard - nenhum ticket foi capturado."
        )
        return

    try:
        page.wait_for_load_state("domcontentloaded", timeout=15000)
    except Exception:
        pass

    logger.info("Solicitação enviada no Phoenix. Aguardando criação e captura do ticket...")
    page.wait_for_timeout(2000)

    descricao_principal = itens_form[0]["description"]
    dados_ticket = _capturar_dados_minhas_solicitacoes(page, descricao_principal)
    if not dados_ticket or not dados_ticket.get("ticket"):
        logger.warning(
            "Não consegui capturar o ticket em 'Minhas Solicitações'. A solicitação NÃO foi "
            "registrada no histórico/dashboard (regra: nunca salvar sem ticket)."
        )
        return

    ticket_num = dados_ticket.get("ticket", "")
    total_produtos = len(itens_form)
    import getpass
    usuario_win = getpass.getuser().upper()
    user_final = (config.get("user") or usuario_win).strip().upper()

    for seq, item_info in enumerate(itens_form, start=1):
        desc_item = item_info.get("description", "")
        mpn_item = item_info.get("mpn", "")
        registro = criar_registro_descricao(
            desc_item,
            user_final,
            ticket=ticket_num,
            status=dados_ticket.get("status") or None,
            data_abertura=dados_ticket.get("data_abertura") or None,
            hora_abertura=dados_ticket.get("hora_abertura") or None,
            produto_seq=seq,
            total_produtos=total_produtos,
            mpn=mpn_item,
        )
        logger.info(
            "Histórico salvo. Linha criada: %s | Ticket: %s | Produto Seq: %s/%s",
            registro["linha"], registro.get("ticket"), seq, total_produtos
        )

        try:
            enviar_para_planilha(desc_item, linha=registro.get("linha"), usuario=user_final)
            logger.info("Payload enviado para a planilha (linha %s).", registro.get("linha"))
        except Exception:
            logger.exception(
                "Falha ao enviar payload para a planilha (o registro %s no histórico já foi salvo).",
                registro.get("linha")
            )


# =====================================================
# CAPTURA: Ticket/Status/Data/Hora em "Minhas Solicitações" logo após o
# envio do formulário (a solicitação só é registrada no histórico/dashboard
# se o ticket for encontrado aqui)
# =====================================================

# Formato observado do ticket: letras seguidas de dígitos (ex: TKGP2026072728).
# AJUSTAR se o padrão real divergir.
TICKET_REGEX = re.compile(r"([A-Z]{2,}\d{6,})")


def _ir_para_minhas_solicitacoes(page: Page) -> bool:
    """Navega até 'Minhas Solicitações' a partir da tela atual do Phoenix."""
    for selector in [
        "text=Minhas Solicitações",
        "text=MINHAS SOLICITAÇÕES",
        'a:has-text("Minhas Solicitações")',
    ]:
        if _esperar_elemento(page, selector, timeout=8000):
            try:
                page.click(selector)
                page.wait_for_load_state("domcontentloaded")
                logger.info("Tela 'Minhas Solicitações' aberta.")
                return True
            except Exception as exc:
                logger.error("Erro ao abrir 'Minhas Solicitações': %s", exc)
                return False

    logger.warning("Não encontrei o menu/link 'Minhas Solicitações'.")
    return False


def _extrair_dados_linha(texto: str) -> Optional[Dict[str, str]]:
    """Extrai ticket/assunto/status/data/hora do texto renderizado."""

    match_ticket = TICKET_REGEX.search(texto)
    ticket = match_ticket.group(1) if match_ticket else ""

    if not ticket:
        return None

    partes = [p.strip() for p in texto.splitlines() if p.strip()]

    status = ""

    for candidato in partes:
        candidato_upper = candidato.upper()

        if candidato_upper in (
            "ON GOING",
            "OPEN",
            "CLOSED",
            "FINALIZADO",
            "PENDING",
            "CANCELLED",
            "APPROVED",
        ):
            status = candidato_upper
            break

    data_match = re.search(r"(\d{2}/\d{2}/\d{4})", texto)
    data_abertura = data_match.group(1) if data_match else ""

    hora_match = re.search(r"(\d{2}:\d{2})", texto)
    hora_abertura = hora_match.group(1) if hora_match else ""

    assunto = ""

    for candidato in partes:
        if (
            candidato == ticket
            or candidato.upper() == status
            or re.fullmatch(r"\d{2}/\d{2}/\d{4}", candidato)
            or re.fullmatch(r"\d{2}:\d{2}", candidato)
        ):
            continue

        if len(candidato) > len(assunto):
            assunto = candidato

    return {
        "ticket": ticket,
        "assunto": assunto,
        "status": status,
        "data_abertura": data_abertura,
        "hora_abertura": hora_abertura,
    }


def _capturar_dados_minhas_solicitacoes(page: Page, descricao: str) -> Optional[Dict[str, str]]:
    """
    Abre 'Minhas Solicitações' e captura Ticket/Status/Data/Hora.
    """

    if not _ir_para_minhas_solicitacoes(page):
        _debug_screenshot(page, "minhas_solicitacoes_nao_encontrada")
        return None

    try:
        page.wait_for_selector("table tbody tr", timeout=20000)
    except PlaywrightTimeoutError:
        logger.warning("Tabela de 'Minhas Solicitações' não carregou a tempo.")
        _debug_screenshot(page, "minhas_solicitacoes_sem_tabela")
        return None

    linhas = page.locator("table tbody tr")
    total = linhas.count()

    if total == 0:
        logger.warning("Tabela de 'Minhas Solicitações' está vazia.")
        return None

    descricao_normalizada = (descricao or "").strip()
    alvo = None

    if descricao_normalizada:
        for i in range(total):
            candidata = linhas.nth(i)

            try:
                texto_linha = candidata.inner_text()
            except Exception:
                continue

            logger.info(
                "Comparando description '%s' com linha '%s'",
                descricao_normalizada,
                texto_linha,
            )

            if descricao_normalizada in texto_linha:
                logger.info("Linha correspondente encontrada.")
                alvo = candidata
                break

    if alvo is None:
        logger.error(
            "Solicitação recém criada não encontrada pela description. "
            "Descrição procurada: %s",
            descricao_normalizada,
        )

        _debug_screenshot(page, "description_nao_encontrada")
        return None

    dados = _extrair_dados_linha(alvo.inner_text())

    if dados is None:
        logger.warning(
            "Não encontrei um ticket no formato esperado na linha capturada."
        )
        return None

    logger.info(
        "Capturado em 'Minhas Solicitações' -> ticket=%s status=%s data=%s hora=%s",
        dados["ticket"],
        dados["status"],
        dados["data_abertura"],
        dados["hora_abertura"],
    )

    return dados
