# =====================================================================
# DISPATCHER DE AUTOMAÇÕES (deve ficar ANTES de todos os imports pesados)
# Quando o EXE é chamado com ``--automacao <nome>``, executa a automação
# diretamente e encerra — sem carregar GUI, matplotlib, openpyxl etc.
# =====================================================================
import sys
import os


def _garantir_chromium_instalado():
    """Verifica se o Chromium do Playwright está instalado; se não, instala automaticamente."""
    browsers_path = os.path.join(os.environ.get("LOCALAPPDATA", ""), "ms-playwright")
    # Verificar se alguma pasta chromium-* existe
    if os.path.isdir(browsers_path):
        for item in os.listdir(browsers_path):
            if item.startswith("chromium-"):
                chromium_dir = os.path.join(browsers_path, item)
                if os.path.isdir(chromium_dir) and os.listdir(chromium_dir):
                    return  # Chromium já instalado
    
    print("=" * 60)
    print("  PRIMEIRA EXECUÇÃO — Instalando navegador Chromium...")
    print("  (isso acontece apenas uma vez, aguarde ~1 minuto)")
    print("=" * 60)
    
    try:
        from playwright._impl._driver import compute_driver_executable
        driver_executable, cli_js = compute_driver_executable()
        import subprocess
        result = subprocess.run(
            [driver_executable, cli_js, "install", "chromium"],
            capture_output=False,
        )
        if result.returncode == 0:
            print("\n[OK] Chromium instalado com sucesso!\n")
        else:
            print(f"\n[AVISO] Instalação retornou código {result.returncode}")
    except Exception as exc:
        print(f"\n[AVISO] Falha ao instalar Chromium automaticamente: {exc}")
        print("Tente instalar manualmente: playwright install chromium")

def _dispatch_automacao() -> bool:
    """Verifica se o EXE foi chamado em modo automação. Retorna True se sim."""
    if "--automacao" not in sys.argv:
        return False

    args = sys.argv[1:]
    idx = args.index("--automacao")
    if idx + 1 >= len(args):
        return False

    nome = args[idx + 1]

    # Parsear argumentos
    arg_valor = None
    cred_user = os.environ.get("PHOENIX_CRED_USER")
    cred_pass = os.environ.get("PHOENIX_CRED_PASS")
    for i, a in enumerate(args):
        if a == "--arg" and i + 1 < len(args):
            arg_valor = args[i + 1]
        elif a == "--cred-user" and i + 1 < len(args):
            cred_user = cred_user or args[i + 1]
        elif a == "--cred-pass" and i + 1 < len(args):
            cred_pass = cred_pass or args[i + 1]

    # No Windows, se o EXE é windowed (console=False), criar console visível
    if os.name == "nt":
        try:
            import ctypes
            kernel32 = ctypes.windll.kernel32
            kernel32.AllocConsole()
            # Reconectar stdout/stderr ao novo console
            sys.stdout = open("CONOUT$", "w", encoding="utf-8", errors="replace")
            sys.stderr = open("CONOUT$", "w", encoding="utf-8", errors="replace")
            sys.stdin = open("CONIN$", "r", encoding="utf-8", errors="replace")
            # Título do console
            kernel32.SetConsoleTitleW(f"Flex-tax classification — Automação {nome.upper()}")
        except Exception:
            pass

    # Reconstruir sys.argv como a automação espera
    new_argv = ["automacao"]
    if arg_valor:
        new_argv.append(arg_valor)
    if cred_user:
        new_argv.extend(["--cred-user", cred_user])
    if cred_pass:
        new_argv.extend(["--cred-pass", cred_pass])
    sys.argv = new_argv

    # Garantir que o navegador Chromium está instalado
    try:
        _garantir_chromium_instalado()
    except Exception as exc:
        print(f"[AVISO] Não foi possível verificar/instalar Chromium: {exc}")

    # Despachar para a automação correta
    try:
        if nome == "phoenix":
            from automocoes.phoenix.phoenix import abrir_home_phoenix, nova_solicitacao_phoenix
            if arg_valor == "home":
                abrir_home_phoenix()
            else:
                nova_solicitacao_phoenix()

        elif nome == "pegasus":
            from automocoes.pegasus.pegasus import abrir_home_pegasus, nova_solicitacao_pegasus
            if arg_valor == "home":
                abrir_home_pegasus()
            else:
                nova_solicitacao_pegasus()

        elif nome == "cost":
            from automocoes.cost.cost import nova_solicitacao_cost
            nova_solicitacao_cost()

        elif nome == "atualizar_pn":
            from automocoes.phoenix.atualizar_pn import atualizar_pn_phoenix
            atualizar_pn_phoenix(arg_valor)

        else:
            print(f"[ERRO] Automação desconhecida: {nome}")
            sys.exit(1)

    except Exception as exc:
        print(f"\n[ERRO] Falha na automação '{nome}': {exc}")
        import traceback
        traceback.print_exc()
        input("\nPressione ENTER para fechar...")
        sys.exit(1)

    return True


if _dispatch_automacao():
    sys.exit(0)

# =====================================================================
# A PARTIR DAQUI: imports normais da GUI (só carrega se NÃO for automação)
# =====================================================================
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from datetime import datetime
from typing import Any, Dict, List, Optional
import threading
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

# matplotlib é importado sob demanda em _renderizar_dashboard_conteudo
# para reduzir consumo de memória no startup (~50 MB a menos).

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from core.paths import get_base_dir, get_db_path, set_db_path
from core.runner import executar_script, logger
# buscar_pn_por_ticket importado sob demanda em _atualizar_pn_dashboard
# para evitar carregar playwright no startup da GUI
from services.logging_config import registrar_log_auditoria, configurar_logger
from services.backup import fazer_backup
from services.database import (
    inicializar_banco,
    validar_usuario,
    validar_usuario_detalhado,
    iniciar_pegasus,
    finalizar_pegasus,
    iniciar_custo,
    finalizar_custo,
    finalizar_phoenix,
    adicionar_usuario,
    editar_usuario,
    desativar_usuario,
    reativar_usuario,
    excluir_usuario,
    esvaziar_lixeira,
    listar_usuarios,
    obter_usuario_por_id,
    obter_credencial,
    salvar_credencial,
    obter_credenciais_usuario,
    contar_por_status,
    exportar_usuarios_json,
    importar_usuarios_json,
    obter_estatisticas_banco,
)
from services.backup import fazer_backup, listar_backups, verificar_e_criar_backup_diario
from services.storage import (
    atualizar_campos_registro,
    carregar_config,
    carregar_estado_app,
    carregar_historico,
    encontrar_por_linha,
    salvar_estado_app,
    cancelar_registro,
)

from ui import theme
from ui.theme import (
    ACCENT,
    ACCENT_HOVER,
    ACCENT_SOFT,
    BG,
    BG_CARD,
    BORDA,
    ENTRY_BG,
    FONT_BOTAO,
    FONT_CAPTION,
    FONT_CARD_TITULO,
    FONT_SUBTITULO,
    FONT_TITULO,
    FOOTER_BG,
    HEADER_BG,
    TEMA_CLARO,
    TEMA_ESCURO,
    TEXTO,
    TEXTO_MUTED,
    definir_tema,
    espacar,
    tema_atual,
)

# Inicializar o banco SQLite na startup
inicializar_banco()

FRAME_LOGIN = "frame_login"
FRAME_MENU = "frame_menu"
FRAME_DASHBOARD = "frame_dashboard"
FRAME_MENU_PHOENIX = "frame_menu_phoenix"
FRAME_MENU_PEGASUS = "frame_menu_pegasus"
FRAME_ATUALIZAR_PN = "frame_atualizar_pn"
FRAME_ADMIN = "frame_admin"
FRAME_CREDENCIAIS = "frame_credenciais"

APP_VERSION = "1.0.0"


def mapear_tela_para_nav(nome: str) -> Optional[str]:
    mapa = {
        FRAME_DASHBOARD: "dashboard",
        FRAME_MENU_PHOENIX: "phoenix",
        FRAME_MENU_PEGASUS: "pegasus",
        FRAME_ATUALIZAR_PN: "phoenix",
        FRAME_ADMIN: "admin",
        FRAME_CREDENCIAIS: "credenciais",
    }
    return mapa.get(nome)


def resumo_ultimo_registro(historico: List[Dict[str, Any]]) -> str:
    if not historico:
        return "Último registro: nenhum"

    item = historico[-1]
    descricao = str(item.get("description") or "Sem description").strip() or "Sem description"
    linha = item.get("linha", "—")
    status = str(item.get("status", "")).strip().upper()

    if status == "ON GOING":
        status_texto = "Em andamento"
    elif status == "CANCELADO":
        status_texto = "Cancelado"
    elif status:
        status_texto = "Finalizado"
    else:
        status_texto = "Sem status"

    return f"Último registro: {descricao} • Linha {linha} • {status_texto}"


class PhoenixTool:

    def _finalizar_phoenix(self, item):
        if not self._validar_permissao("editar"):
            return
        finalizar_phoenix(item["linha"])
        self._abrir_dashboard()

    def _iniciar_pegasus(self, item):
        if not self._validar_permissao("pegasus"):
            return
        iniciar_pegasus(item["linha"])
        self._abrir_dashboard()

    def _finalizar_pegasus(self, item):
        if not self._validar_permissao("pegasus"):
            return
        finalizar_pegasus(item["linha"])
        self._abrir_dashboard()

    def _iniciar_custo(self, item):
        if not self._validar_permissao("cost"):
            return
        iniciar_custo(item["linha"])
        self._abrir_dashboard()

    def _finalizar_custo(self, item):
        if not self._validar_permissao("cost"):
            return
        finalizar_custo(item["linha"])
        self._abrir_dashboard()

    def _excluir_registro(self, item):
        if not self._validar_permissao("excluir"):
            return

        confirmar = messagebox.askyesno(
            "Excluir Registro",
            (
                f"Deseja excluir a linha {item.get('linha')}?\n\n"
                f"Descrição:\n{item.get('description')}\n\n"
                "Esta ação ocultará o registro do dashboard."
            )
        )

        if not confirmar:
            return

        sucesso = cancelar_registro(
            item.get("uuid"),
            "Excluído manualmente pelo usuário",
            usuario_cancelamento=self.usuario,
        )

        if sucesso:
            messagebox.showinfo("Sucesso", "Registro excluído com sucesso.")
            if FRAME_DASHBOARD in self.frames:
                self._build_dashboard()
                self.mostrar(FRAME_DASHBOARD)
        else:
            messagebox.showerror("Erro", "Não foi possível excluir o registro.")

    def __init__(self, root):
        self.root = root
        self.root.overrideredirect(True)
        self._drag_x = 0
        self._drag_y = 0
        self.root.title("Flex-tax classification 1.0")
        self.root.resizable(True, True)
        self.root.minsize(560, 720)
        self.root.state("zoomed")
        self._modo_tela_cheia = False
        self.root.option_add("*Font", "Arial")
        self.root.option_add("*tearOff", False)
        self.root.attributes("-alpha", 0.98)
        self.root.overrideredirect(False)
        self.root.bind("<ButtonPress-1>", self._iniciar_arraste)
        self.root.bind("<B1-Motion>", self._arrastar_janela)

        self.usuario = None
        self.usuario_id = None
        self.permissao = None
        self.history = []
        self.frames = {}
        self._dashboard_filter = "todos"
        self._dashboard_last_signature = None
        self._cards = {}
        self._active_card_key = None

        estado_inicial = carregar_estado_app()
        theme.definir_tema(estado_inicial.get("theme") or theme.TEMA_ESCURO)
        self._sincronizar_cores()

        self.root.configure(bg=BG)
        self._montar_estrutura_chrome()

        self._build_login()
        self._show_loading_state()
        self.root.after(350, self._finalizar_inicializacao)

    def _sincronizar_cores(self):
        global BG, BG_CARD, BORDA, TEXTO, TEXTO_MUTED, ACCENT, ACCENT_SOFT, HEADER_BG, FOOTER_BG
        BG = theme.BG
        BG_CARD = theme.BG_CARD
        BORDA = theme.BORDA
        TEXTO = theme.TEXTO
        TEXTO_MUTED = theme.TEXTO_MUTED
        ACCENT = theme.ACCENT
        ACCENT_SOFT = theme.ACCENT_SOFT
        HEADER_BG = theme.HEADER_BG
        FOOTER_BG = theme.FOOTER_BG

    def _validar_permissao(self, acao: str) -> bool:
        """Validação interna estrita de permissões antes de qualquer ação crítica."""
        if not self.permissao:
            messagebox.showerror("Acesso Negado", "Usuário não autenticado.")
            return False

        perm_upper = str(self.permissao).upper()

        if perm_upper == "ADMIN":
            return True

        if perm_upper == "ENGENHARIA":
            permitidos = {
                "dashboard", "nova_solicitacao", "phoenix", "pegasus", "cost",
                "editar", "atualizar_pn", "exportar", "credenciais"
            }
            if acao in permitidos:
                return True

        if perm_upper == "VISITANTE":
            permitidos = {"dashboard", "consulta", "exportar"}
            if acao in permitidos:
                return True

        messagebox.showerror("Acesso Negado", f"Seu perfil ({perm_upper}) não tem permissão para a ação: {acao}.")
        return False

    def _montar_estrutura_chrome(self):
        self._aplicar_estilo()

        self.header = tk.Frame(self.root, bg=HEADER_BG, height=52, highlightbackground=BORDA, highlightthickness=0)
        self.header.pack(fill="x", side="top")
        self.header.pack_propagate(False)

        tk.Label(self.header, text="FLEX • CLASSIFICAÇÃO FISCAL", bg=HEADER_BG, fg=ACCENT, font=("Arial", 10, "bold")).pack(side="left", padx=18)
        texto_badge = f"USUÁRIO: {self.usuario.upper()} ({self.permissao})" if self.usuario else "USUÁRIO: NÃO LOGADO"
        self.user_badge = tk.Label(self.header, text=texto_badge, bg=HEADER_BG, fg=TEXTO_MUTED, font=("Arial", 8))
        self.user_badge.pack(side="left", padx=(12, 0))
        tk.Label(self.header, text="STATUS: ONLINE", bg=HEADER_BG, fg=TEXTO_MUTED, font=("Arial", 9)).pack(side="right", padx=18)

        self.btn_tema = tk.Button(
            self.header,
            text="☀️ Modo Claro" if tema_atual() == TEMA_ESCURO else "🌙 Modo Escuro",
            command=self._alternar_tema,
            bg=HEADER_BG, fg=ACCENT, activebackground=HEADER_BG, activeforeground=ACCENT,
            relief="solid", bd=1, font=("Arial", 8, "bold"), padx=8, pady=2, cursor="hand2"
        )
        self.btn_tema.pack(side="right", padx=(0, 10))

    def _alternar_tema(self):
        novo_tema = TEMA_CLARO if tema_atual() == TEMA_ESCURO else TEMA_ESCURO
        definir_tema(novo_tema)

        self.root.config(bg=BG)
        if hasattr(self, "header") and self.header.winfo_exists():
            self.header.config(bg=HEADER_BG)
        if hasattr(self, "footer") and self.footer.winfo_exists():
            self.footer.config(bg=FOOTER_BG)
        if hasattr(self, "nav_left") and self.nav_left.winfo_exists():
            self.nav_left.config(bg=FOOTER_BG)
        if hasattr(self, "signature_frame") and self.signature_frame.winfo_exists():
            self.signature_frame.config(bg=BG)
        if hasattr(self, "signature_text") and self.signature_text.winfo_exists():
            self.signature_text.config(bg=BG, fg=TEXTO_MUTED)
        if hasattr(self, "signature_name") and self.signature_name.winfo_exists():
            self.signature_name.config(bg=BG, fg=TEXTO)
        if hasattr(self, "main_canvas") and self.main_canvas.winfo_exists():
            self.main_canvas.config(bg=BG)
        if hasattr(self, "btn_tema") and self.btn_tema.winfo_exists():
            self.btn_tema.config(
                text="☀️ Modo Claro" if novo_tema == TEMA_ESCURO else "🌙 Modo Escuro",
                bg=HEADER_BG, fg=ACCENT, activebackground=HEADER_BG, activeforeground=ACCENT
            )

        if getattr(self, "history", None):
            tela_atual = self.history[-1]
            if tela_atual == FRAME_DASHBOARD:
                self._abrir_dashboard()
            elif tela_atual == FRAME_LOGIN:
                self._build_login()
            elif tela_atual == FRAME_ADMIN:
                self._abrir_administracao()
            elif tela_atual == FRAME_CREDENCIAIS:
                self._abrir_minhas_credenciais()
            elif tela_atual == FRAME_MENU_PHOENIX:
                self._abrir_menu_phoenix()
            elif tela_atual == FRAME_MENU_PEGASUS:
                self._abrir_menu_pegasus()

        self.footer = tk.Frame(self.root, bg=FOOTER_BG, height=30, highlightbackground=BORDA, highlightthickness=0)
        self.footer.pack(fill="x", side="bottom")
        self.footer.pack_propagate(False)

        self.nav_left = tk.Frame(self.footer, bg=FOOTER_BG)
        self.nav_left.pack(side="left", padx=18)
        self.nav_labels = {}

        nav_right = tk.Frame(self.footer, bg=FOOTER_BG)
        nav_right.pack(side="right", padx=18)
        self.footer_status = tk.Label(nav_right, text="PRONTO", bg=FOOTER_BG, fg=TEXTO_MUTED, font=("Arial", 8))
        self.footer_status.pack(side="right")
        tk.Label(nav_right, text="  •  ", bg=FOOTER_BG, fg=TEXTO_MUTED, font=("Arial", 8)).pack(side="right")
        tk.Label(nav_right, text="FLEX-TAX CLASSIFICATION 1.0", bg=FOOTER_BG, fg=TEXTO_MUTED, font=("Arial", 8)).pack(side="right")

        self.signature_frame = tk.Frame(self.root, bg=BG)
        self.signature_frame.pack(side="bottom", pady=5)

        self.signature_text = tk.Label(
            self.signature_frame,
            text="Developed by ",
            bg=BG,
            fg="#6f6f6f",
            font=("Arial", 8)
        )
        self.signature_text.pack(side="left")

        self.signature_name = tk.Label(
            self.signature_frame,
            text="Gabriel Girotto",
            bg=BG,
            fg="white",
            font=("Arial", 8, "bold")
        )
        self.signature_name.pack(side="left")

        area_rolagem = tk.Frame(self.root, bg=BG)
        area_rolagem.pack(fill="both", expand=True)

        self.main_canvas = tk.Canvas(
            area_rolagem,
            bg=BG,
            highlightthickness=0,
            bd=0
        )

        self.vscrollbar = tk.Scrollbar(
            area_rolagem,
            orient="vertical",
            command=self.main_canvas.yview
        )

        self.main_canvas.configure(
            yscrollcommand=self.vscrollbar.set
        )

        self.main_canvas.pack(
            side="left",
            fill="both",
            expand=True
        )

        self.container = tk.Frame(
            self.main_canvas,
            bg=BG
        )

        self.container.grid_columnconfigure(0, weight=1)

        self._container_window = self.main_canvas.create_window(
            (0, 0),
            window=self.container,
            anchor="nw"
        )

        def _atualizar_scrollregion(event=None):
            self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))

        def _ajustar_largura_container(event):
            largura = max(event.width, self.container.winfo_reqwidth())
            self.main_canvas.itemconfig(self._container_window, width=largura)

        self.container.bind("<Configure>", _atualizar_scrollregion)
        self.main_canvas.bind("<Configure>", _ajustar_largura_container)

        def _scroll_vertical(event):
            self.main_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _scroll_horizontal(event):
            self.main_canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")

        def _bind_scroll(event=None):
            self.main_canvas.bind_all("<MouseWheel>", _scroll_vertical)
            self.main_canvas.bind_all("<Shift-MouseWheel>", _scroll_horizontal)

        def _unbind_scroll(event=None):
            self.main_canvas.unbind_all("<MouseWheel>")
            self.main_canvas.unbind_all("<Shift-MouseWheel>")

        self.main_canvas.bind("<Enter>", _bind_scroll)
        self.main_canvas.bind("<Leave>", _unbind_scroll)

    def _recriar_itens_nav(self):
        """Atualiza a barra de navegação inferior de acordo com o perfil."""
        for widget in self.nav_left.winfo_children():
            widget.destroy()
        self.nav_labels = {}

        if not self.permissao:
            return

        perm = str(self.permissao).upper()
        self._criar_item_nav(self.nav_left, "DASHBOARD", "dashboard")

        if perm in ("ADMIN", "ENGENHARIA"):
            tk.Label(self.nav_left, text="  •  ", bg=FOOTER_BG, fg=TEXTO_MUTED, font=("Arial", 8)).pack(side="left")
            self._criar_item_nav(self.nav_left, "PHOENIX", "phoenix")
            tk.Label(self.nav_left, text="  •  ", bg=FOOTER_BG, fg=TEXTO_MUTED, font=("Arial", 8)).pack(side="left")
            self._criar_item_nav(self.nav_left, "PEGASUS", "pegasus")
            tk.Label(self.nav_left, text="  •  ", bg=FOOTER_BG, fg=TEXTO_MUTED, font=("Arial", 8)).pack(side="left")
            self._criar_item_nav(self.nav_left, "CREDENCIAIS", "credenciais")

        if perm == "ADMIN":
            tk.Label(self.nav_left, text="  •  ", bg=FOOTER_BG, fg=TEXTO_MUTED, font=("Arial", 8)).pack(side="left")
            self._criar_item_nav(self.nav_left, "ADMINISTRAÇÃO", "admin")

    def _iniciar_arraste(self, event):
        self._drag_offset = (event.x_root - self.root.winfo_x(), event.y_root - self.root.winfo_y())

    def _arrastar_janela(self, event):
        if getattr(self, "_drag_offset", None) is None:
            return
        x = event.x_root - self._drag_offset[0]
        y = event.y_root - self._drag_offset[1]
        self.root.geometry(f"+{x}+{y}")

    def _alternar_tela_cheia(self):
        if self.root.attributes("-fullscreen"):
            self.root.attributes("-fullscreen", False)
            self.root.geometry("1100x900")
            self._modo_tela_cheia = False
        else:
            self.root.attributes("-fullscreen", True)
            self._modo_tela_cheia = True

    def _criar_item_nav(self, parent, texto, chave):
        def _navegar():
            if chave == "dashboard":
                self._abrir_dashboard()
            elif chave == "phoenix":
                self._abrir_menu_phoenix()
            elif chave == "pegasus":
                self._abrir_menu_pegasus()
            elif chave == "credenciais":
                self._abrir_minhas_credenciais()
            elif chave == "admin":
                self._abrir_administracao()

        label = tk.Label(
            parent,
            text=texto,
            bg=FOOTER_BG,
            fg=TEXTO_MUTED,
            font=("Arial", 8, "bold"),
            cursor="hand2",
        )
        label.pack(side="left")
        label.bind("<Button-1>", lambda e: _navegar())
        self.nav_labels[chave] = label
        return label

    def _atualizar_nav(self, nome):
        ativo = mapear_tela_para_nav(nome)
        for chave, label in self.nav_labels.items():
            label.config(fg=ACCENT if chave == ativo else TEXTO_MUTED)
        if hasattr(self, "footer_status"):
            if self.usuario:
                self.footer_status.config(text=f"USUÁRIO: {self.usuario.upper()} ({self.permissao})  •  TELA: {nome}")
            else:
                self.footer_status.config(text=f"TELA: {nome}")

    def _aplicar_estilo(self):
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")

        style.configure("TFrame", background=BG)
        style.configure("TLabel", background=BG, foreground=TEXTO)
        style.configure("Muted.TLabel", background=BG, foreground=TEXTO_MUTED)
        style.configure("TButton", background=BG_CARD, foreground=TEXTO, borderwidth=1, padding=(12, 10))
        style.map(
            "TButton",
            background=[("active", BORDA), ("pressed", BORDA)],
            foreground=[("active", TEXTO)]
        )
        style.configure("Accent.TButton", background=ACCENT_SOFT, foreground=TEXTO, borderwidth=1, padding=(12, 10))
        style.map(
            "Accent.TButton",
            background=[("active", "#263b57"), ("pressed", "#263b57")],
            foreground=[("active", TEXTO)]
        )
        style.configure("TEntry", fieldbackground=BG_CARD, foreground=TEXTO, borderwidth=1, padding=(8, 8))
        style.map("TEntry", fieldbackground=[("focus", "#1c2230")], foreground=[("focus", TEXTO)])

    def mostrar(self, nome, empilhar=True):
        for f in self.frames.values():
            f.pack_forget()
        if empilhar and (not self.history or self.history[-1] != nome):
            self.history.append(nome)
        if nome in self.frames:
            self.frames[nome].pack(fill="both", expand=True)
            self._animar_troca_tela(self.frames[nome])
            self._atualizar_nav(nome)
            self._atualizar_cards_ativos(nome)
            salvar_estado_app(nome)

    def voltar(self):
        if len(self.history) > 1:
            self.history.pop()
            anterior = self.history[-1]
            self.mostrar(anterior, empilhar=False)

    def _animar_troca_tela(self, frame):
        try:
            frame.configure(bg=BG)
            frame.update_idletasks()
        except Exception:
            pass
        self._animar_fade()

    def _animar_fade(self, alvo=0.98, passos=6, intervalo=18):
        try:
            atual = float(self.root.attributes("-alpha"))
        except Exception:
            atual = 0.98

        if atual >= alvo:
            self.root.attributes("-alpha", alvo)
            return

        passo = (alvo - atual) / passos

        def _aplicar(idx):
            if idx >= passos:
                self.root.attributes("-alpha", alvo)
                return
            nova = atual + passo * (idx + 1)
            self.root.attributes("-alpha", nova)
            self.root.after(intervalo, lambda i=idx + 1: _aplicar(i))

        _aplicar(0)

    def _show_loading_state(self):
        self._loading_frame = tk.Frame(self.container, bg=BG)
        self._loading_frame.pack(pady=(24, 0))

        self._loading_label = tk.Label(
            self._loading_frame,
            text="Inicializando ambiente operacional...",
            bg=BG,
            fg=TEXTO_MUTED,
            font=("Arial", 10),
        )
        self._loading_label.pack(anchor="w")

        self._loading_bar = tk.Frame(self._loading_frame, bg=ACCENT_SOFT, height=8)
        self._loading_bar.pack(fill="x", pady=(10, 0), ipady=2)
        self._loading_fill = tk.Frame(self._loading_bar, bg=ACCENT, width=30, height=8)
        self._loading_fill.pack(side="left")
        self._loading_fill.pack_propagate(False)
        self._loading_progress = 0
        self._animar_loading()

    def _animar_loading(self):
        loading_fill = getattr(self, "_loading_fill", None)
        if loading_fill is None or not loading_fill.winfo_exists():
            return
        self._loading_progress = (self._loading_progress + 1) % 101
        largura = 20 + (self._loading_progress % 80)
        loading_fill.configure(width=largura)
        if getattr(self, "_loading_frame", None) is not None and self._loading_frame.winfo_exists():
            self.root.after(25, self._animar_loading)

    def _finalizar_inicializacao(self):
        if getattr(self, "_loading_frame", None) is not None:
            self._loading_frame.destroy()
            self._loading_frame = None
            self._loading_fill = None

        self.mostrar(FRAME_LOGIN, empilhar=False)

    def linha_divisoria(self, parent):
        tk.Frame(parent, bg=BORDA, height=1).pack(fill="x", padx=40, pady=(0, 20))

    def cabecalho(self, parent, caption, titulo):
        tk.Label(
            parent, text=espacar(caption), bg=BG, fg=ACCENT, font=FONT_CAPTION
        ).pack(pady=(40, 6))

        tk.Label(
            parent, text=titulo.upper(), bg=BG, fg=TEXTO, font=FONT_TITULO
        ).pack(pady=(0, 12))

        self.linha_divisoria(parent)

    def botao_voltar(self, parent):
        lbl = tk.Label(
            parent, text=espacar("← Voltar"), bg=BG, fg=ACCENT,
            font=FONT_CAPTION, cursor="hand2"
        )
        lbl.pack(anchor="w", padx=40, pady=(20, 0))
        lbl.bind("<Enter>", lambda e: lbl.config(fg=TEXTO))
        lbl.bind("<Leave>", lambda e: lbl.config(fg=TEXTO_MUTED))
        lbl.bind("<Button-1>", lambda e: self.voltar())
        return lbl

    def botao_flat(self, parent, texto, comando, largura=None):
        texto_exibicao = espacar(texto) if len(texto) <= 8 else texto.upper()
        btn = tk.Button(
            parent, text=texto_exibicao, command=comando,
            bg=ACCENT_SOFT, fg=TEXTO, activebackground=ACCENT, activeforeground="#ffffff",
            relief="solid", bd=1, font=("Arial", 8, "bold"), cursor="hand2",
            width=largura, highlightbackground=BORDA, highlightthickness=1,
            padx=10, pady=5
        )
        def _on_enter(e):
            btn.config(bg=ACCENT, fg="#ffffff", highlightbackground=ACCENT)
        def _on_leave(e):
            btn.config(bg=ACCENT_SOFT, fg=TEXTO, highlightbackground=BORDA)
        btn.bind("<Enter>", _on_enter)
        btn.bind("<Leave>", _on_leave)
        return btn

    def _aplicar_estado_card(self, card, lbl_titulo, lbl_num, ativo):
        card.configure(bg="#151515" if ativo else BG_CARD)
        lbl_titulo.configure(fg=ACCENT if ativo else TEXTO)
        lbl_num.configure(fg=ACCENT)

    def _atualizar_cards_ativos(self, nome):
        chave = mapear_tela_para_nav(nome)
        self._active_card_key = chave
        for card_key, widgets in self._cards.items():
            self._aplicar_estado_card(*widgets, ativo=(card_key == chave))

    def card_navegavel(self, parent, numero, titulo, comando, descricao=None, chave=None):
        card = tk.Frame(parent, bg=BG_CARD, cursor="hand2", highlightbackground=BORDA, highlightthickness=0, bd=0)
        card.pack(fill="x", padx=28, pady=(0, 8))

        linha = tk.Frame(card, bg=BG_CARD)
        linha.pack(fill="x", padx=14, pady=10)

        lbl_num = tk.Label(
            linha, text=f".{numero:02d}", bg=BG_CARD, fg=ACCENT, font=("Arial", 10, "bold")
        )
        lbl_num.pack(side="left", padx=(0, 10))

        lbl_titulo = tk.Label(
            linha, text=titulo.upper(), bg=BG_CARD, fg=TEXTO, font=FONT_CARD_TITULO
        )
        lbl_titulo.pack(side="left")

        if descricao:
            tk.Label(
                linha,
                text=descricao,
                bg=BG_CARD,
                fg=TEXTO_MUTED,
                font=("Arial", 8),
            ).pack(side="right")

        divisor = tk.Frame(card, bg=BORDA, height=1)
        divisor.pack(fill="x")

        if chave:
            self._cards[chave] = (card, lbl_titulo, lbl_num)

        def entrar(e=None):
            if self._active_card_key == chave:
                self._aplicar_estado_card(card, lbl_titulo, lbl_num, True)
            else:
                self._aplicar_estado_card(card, lbl_titulo, lbl_num, False)
                lbl_titulo.config(fg=ACCENT)
                lbl_num.config(fg=ACCENT)
                card.config(bg="#151515")

        def sair(e=None):
            self._aplicar_estado_card(card, lbl_titulo, lbl_num, self._active_card_key == chave)
            if self._active_card_key != chave:
                lbl_titulo.config(fg=TEXTO)
                lbl_num.config(fg=ACCENT)
                card.config(bg=BG_CARD)

        for widget in (card, linha, lbl_num, lbl_titulo):
            widget.bind("<Enter>", entrar)
            widget.bind("<Leave>", sair)
            widget.bind("<Button-1>", lambda e: comando())

        self._aplicar_estado_card(card, lbl_titulo, lbl_num, self._active_card_key == chave)
        return card

    def campo_entry(self, parent, mostrar=None):
        entry = tk.Entry(
            parent, bg=ENTRY_BG, fg=TEXTO, insertbackground=ACCENT,
            relief="solid", font=("Arial", 11), show=mostrar,
            highlightbackground=BORDA, highlightthickness=1, bd=0
        )
        def _on_focus_in(e):
            entry.config(highlightbackground=ACCENT, highlightthickness=2)
        def _on_focus_out(e):
            entry.config(highlightbackground=BORDA, highlightthickness=1)
        entry.bind("<FocusIn>", _on_focus_in)
        entry.bind("<FocusOut>", _on_focus_out)
        return entry

    def build_menu_automacao(self, parent, titulo, itens):
        self.botao_voltar(parent)
        self.cabecalho(parent, "Automação", titulo)
        for index, item in enumerate(itens, start=1):
            if isinstance(item, tuple) and len(item) == 3:
                label, comando, descricao = item
            else:
                label, comando = item
                descricao = None
            self.card_navegavel(parent, index, label, comando, descricao)

    # =========================================================================
    # LOGIN (Refatorado: Nome Completo + Código Engenharia, sem Senha)
    # =========================================================================

    def _build_login(self):
        if FRAME_LOGIN in self.frames:
            self.frames[FRAME_LOGIN].destroy()

        frame = tk.Frame(self.container, bg=BG)
        self.frames[FRAME_LOGIN] = frame

        painel = tk.Frame(frame, bg=BG_CARD, highlightbackground=BORDA, highlightthickness=0)
        painel.pack(fill="x", padx=28, pady=(38, 16))

        tk.Label(
            painel, text="FLEX • CLASSIFICAÇÃO FISCAL", bg=BG_CARD, fg=ACCENT, font=("Arial", 9, "bold")
        ).pack(pady=(24, 6))

        tk.Label(
            painel, text="LOGIN CORPORATIVO", bg=BG_CARD, fg=TEXTO, font=FONT_TITULO
        ).pack(pady=(0, 6))

        tk.Label(
            painel, text="Informe seu Nome Completo e Código de Engenharia para acessar.", bg=BG_CARD, fg=TEXTO_MUTED, font=FONT_SUBTITULO
        ).pack(pady=(0, 18))

        bloco = tk.Frame(painel, bg=BG_CARD)
        bloco.pack(padx=20, pady=8, fill="x")

        tk.Label(
            bloco, text=espacar("Nome Completo"), bg=BG_CARD, fg=TEXTO_MUTED, font=FONT_CAPTION
        ).pack(anchor="w", pady=(20, 6))

        self.campo_nome = self.campo_entry(bloco)
        self.campo_nome.pack(fill="x", ipady=8)

        tk.Label(
            bloco, text=espacar("Código Engenharia"), bg=BG_CARD, fg=TEXTO_MUTED, font=FONT_CAPTION
        ).pack(anchor="w", pady=(24, 6))

        self.campo_codigo = self.campo_entry(bloco, mostrar="*")
        self.campo_codigo.pack(fill="x", ipady=8)

        self.campo_nome.bind("<Return>", lambda e: self.entrar())
        self.campo_codigo.bind("<Return>", lambda e: self.entrar())

        self.botao_flat(bloco, "Entrar", self.entrar).pack(fill="x", pady=(24, 12))

        btn_help = tk.Button(
            bloco,
            text="❓ Não tem cadastro? Clique aqui para saber como solicitar",
            command=self._dialog_ajuda_cadastro,
            bg=BG_CARD,
            fg=ACCENT,
            activebackground=BG_CARD,
            activeforeground=ACCENT,
            bd=0,
            font=("Arial", 9, "underline"),
            cursor="hand2"
        )
        btn_help.pack(anchor="center", pady=(0, 16))

    def _dialog_ajuda_cadastro(self):
        janela = tk.Toplevel(self.root)
        janela.title("Como Solicitar Cadastro")
        janela.configure(bg=BG)
        janela.attributes("-topmost", True)
        janela.resizable(False, False)
        janela.transient(self.root)

        corpo = tk.Frame(janela, bg=BG, padx=24, pady=20)
        corpo.pack(fill="both", expand=True)

        tk.Label(
            corpo, text="🔐 Solicitação de Cadastro e Autorização", bg=BG, fg=ACCENT, font=("Arial", 12, "bold")
        ).pack(anchor="w", pady=(0, 10))

        msg = (
            "Para solicitar acesso ao sistema Flex-tax classification:\n\n"
            "1. Envie um e-mail explicando o motivo da sua solicitação.\n"
            "2. Destinatário: gabriell.girotto@flex.com\n"
            "3. Informe no e-mail o seu Nome Completo.\n\n"
            "Após a análise e aprovação, seu cadastro será liberado pelo Administrador."
        )

        tk.Label(
            corpo, text=msg, bg=BG, fg=TEXTO, font=("Arial", 9), justify="left"
        ).pack(anchor="w", pady=(0, 16))

        self.botao_flat(corpo, "Entendido", janela.destroy).pack(fill="x")

        janela.update_idletasks()
        largura, altura = janela.winfo_width(), janela.winfo_height()
        x = (janela.winfo_screenwidth() // 2) - (largura // 2)
        y = (janela.winfo_screenheight() // 2) - (altura // 2)
        janela.geometry(f"+{x}+{y}")
        janela.grab_set()

    def entrar(self, automatic=False):
        nome = self.campo_nome.get().strip()
        codigo = self.campo_codigo.get().strip()

        if not nome:
            messagebox.showwarning("Aviso", "Digite o Nome Completo.")
            return
        if not codigo:
            messagebox.showwarning("Aviso", "Digite o Código Engenharia.")
            return

        sucesso, mensagem_motivo, resultado = validar_usuario_detalhado(nome, codigo)

        if not sucesso or not resultado:
            db_atual = get_db_path()
            messagebox.showerror(
                "Acesso Negado",
                f"{mensagem_motivo}\n\n"
                f"• Banco de Dados em Uso:\n{db_atual}\n\n"
                f"💡 Verifique se o cadastro foi salvo exatamente neste arquivo de banco de dados ou se a conta foi importada."
            )
            return

        permissao, user_id = resultado
        self.usuario = nome
        self.usuario_id = user_id
        self.permissao = permissao

        logger.info("Login efetuado (usuário=%s, id=%s, permissão=%s)", nome, user_id, permissao)
        registrar_log_auditoria("LOGIN", nome, f"ID: {user_id} | Perfil: {permissao}")

        if hasattr(self, "user_badge"):
            self.user_badge.config(text=f"USUÁRIO: {nome.upper()} ({permissao})")

        self._recriar_itens_nav()
        self._build_menu()
        self.history = []
        self.mostrar(FRAME_MENU)

    # =========================================================================
    # MENU PRINCIPAL (Específico por Perfil)
    # =========================================================================

    def _build_menu(self):
        if FRAME_MENU in self.frames:
            self.frames[FRAME_MENU].destroy()

        frame = tk.Frame(self.container, bg=BG)
        self.frames[FRAME_MENU] = frame

        painel = tk.Frame(frame, bg=BG_CARD, highlightbackground=BORDA, highlightthickness=0)
        painel.pack(fill="x", padx=28, pady=(18, 12))
        painel.configure(borderwidth=0)

        tk.Label(
            painel,
            text="● FLEX-TAX CLASSIFICATION 1.0",
            bg=BG_CARD,
            fg=ACCENT,
            font=("Arial", 8, "bold"),
        ).pack(anchor="w", padx=16, pady=(14, 0))

        tk.Label(
            painel, text="PAINEL OPERACIONAL", bg=BG_CARD, fg=ACCENT, font=("Arial", 9, "bold")
        ).pack(anchor="w", padx=18, pady=(12, 4))

        tk.Label(
            painel, text=f"Bem-vindo, {self.usuario}", bg=BG_CARD, fg=TEXTO, font=("Arial", 16, "bold")
        ).pack(anchor="w", padx=18, pady=(0, 4))

        tk.Label(
            painel, text=f"Perfil de acesso: {self.permissao}", bg=BG_CARD, fg=TEXTO_MUTED, font=FONT_SUBTITULO
        ).pack(anchor="w", padx=18, pady=(0, 14))

        tk.Label(
            frame,
            text=datetime.now().strftime("%d/%m/%Y  •  %H:%M"),
            bg=BG,
            fg=TEXTO_MUTED,
            font=("Arial", 9),
        ).pack(anchor="w", padx=28, pady=(0, 6))

        tk.Label(
            frame,
            text=resumo_ultimo_registro(self._carregar_historico()),
            bg=BG,
            fg=TEXTO_MUTED,
            font=("Arial", 9),
        ).pack(anchor="w", padx=28, pady=(0, 10))

        self.cabecalho(frame, "Navegação", "Menu")

        idx = 1
        perm = str(self.permissao).upper()

        # Todos os perfis veem o Dashboard
        self.card_navegavel(frame, idx, "Dashboard", self._abrir_dashboard, "Visão geral operacional e métricas", chave="dashboard")
        idx += 1

        if perm in ("ADMIN", "ENGENHARIA"):
            self.card_navegavel(frame, idx, "Phoenix", self._abrir_menu_phoenix, "Fluxo principal e solicitações", chave="phoenix")
            idx += 1
            self.card_navegavel(frame, idx, "Pegasus", self._abrir_menu_pegasus, "Fluxo Pegasus e classificações", chave="pegasus")
            idx += 1
            self.card_navegavel(frame, idx, "Cost Request", self._executar_cost_request, "Automação de custo")
            idx += 1
            self.card_navegavel(frame, idx, "Minhas Credenciais", self._abrir_minhas_credenciais, "Gerenciar logins de automação", chave="credenciais")
            idx += 1

        if perm == "ADMIN":
            self.card_navegavel(frame, idx, "Administração", self._abrir_administracao, "Gerenciar usuários e permissões do sistema", chave="admin")
            idx += 1

        tk.Label(
            frame, text=f"Status: Conectado • Perfil: {perm}", bg=BG, fg=ACCENT,
            font=FONT_CAPTION
        ).pack(side="bottom", pady=16)

    # =========================================================================
    # GERENCIAMENTO DE CREDENCIAIS INDIVIDUAIS PARA AUTOMAÇÕES
    # =========================================================================

    def _obter_ou_pedir_credencial(self, tipo: str, forcar: bool = False) -> Optional[Dict[str, str]]:
        """Busca credencial no SQLite; se não existir ou se forcar=True, abre modal para solicitar."""
        cred = obter_credencial(self.usuario_id, tipo)
        if cred and cred.get("login") and cred.get("senha") and not forcar:
            return cred

        # Não encontrou credencial salvos ou solicitou alteração -> Abrir modal
        resultado = [None]

        janela = tk.Toplevel(self.root)
        janela.title(f"Credenciais {tipo}")
        janela.configure(bg=BG)
        janela.attributes("-topmost", True)
        janela.resizable(False, False)
        janela.transient(self.root)

        corpo = tk.Frame(janela, bg=BG, padx=24, pady=20)
        corpo.pack(fill="both", expand=True)

        tk.Label(
            corpo, text=f"Credenciais {tipo}", bg=BG, fg=TEXTO, font=("Arial", 14, "bold")
        ).pack(anchor="w", pady=(0, 6))

        tk.Label(
            corpo,
            text=f"Informe suas credenciais individuais para a automação {tipo}.\nSerão salvas de forma segura no SQLite.",
            bg=BG, fg=TEXTO_MUTED, font=("Arial", 9), justify="left"
        ).pack(anchor="w", pady=(0, 16))

        tk.Label(corpo, text="JAG / Usuário:", bg=BG, fg=TEXTO_MUTED, font=FONT_CAPTION).pack(anchor="w", pady=(4, 2))
        campo_jag = self.campo_entry(corpo)
        campo_jag.pack(fill="x", ipady=6, pady=(0, 10))

        tk.Label(corpo, text="Senha:", bg=BG, fg=TEXTO_MUTED, font=FONT_CAPTION).pack(anchor="w", pady=(4, 2))
        campo_senha = self.campo_entry(corpo, mostrar="*")
        campo_senha.pack(fill="x", ipady=6, pady=(0, 16))

        # Preencher campos com valores existentes se disponíveis
        if cred:
            if cred.get("login"):
                campo_jag.insert(0, cred.get("login"))
            if cred.get("senha"):
                campo_senha.insert(0, cred.get("senha"))

        def salvar():
            jag = campo_jag.get().strip()
            senha = campo_senha.get().strip()
            if not jag or not senha:
                messagebox.showwarning("Aviso", "Preencha o JAG e a Senha.", parent=janela)
                return
            salvar_credencial(self.usuario_id, tipo, jag, senha)
            resultado[0] = {"login": jag, "senha": senha}
            janela.destroy()

        self.botao_flat(corpo, "Salvar", salvar).pack(fill="x")

        janela.update_idletasks()
        largura, altura = janela.winfo_width(), janela.winfo_height()
        x = (janela.winfo_screenwidth() // 2) - (largura // 2)
        y = (janela.winfo_screenheight() // 2) - (altura // 2)
        janela.geometry(f"+{x}+{y}")
        campo_jag.focus_set()
        janela.grab_set()
        self.root.wait_window(janela)

        return resultado[0]

    def _executar_cost_request(self):
        if not self._validar_permissao("cost"):
            return
        cred = self._obter_ou_pedir_credencial("COST")
        if not cred:
            return
        executar_script("automocoes", "cost", "cost.py", cred_user=cred["login"], cred_pass=cred["senha"])

    def _carregar_historico(self):
        return carregar_historico()

    def _abrir_dashboard(self):
        self._build_dashboard()
        self.mostrar(FRAME_DASHBOARD)

    # =========================================================================
    # DASHBOARD & EXPORTAÇÕES EXCEL
    # =========================================================================

    def _build_dashboard(self):
        if FRAME_DASHBOARD in self.frames:
            self.frames[FRAME_DASHBOARD].destroy()

        frame = tk.Frame(self.container, bg=BG)
        self.frames[FRAME_DASHBOARD] = frame

        self.botao_voltar(frame)
        self.cabecalho(frame, "Visão geral", "Dashboard")

        painel = tk.Frame(frame, bg=BG_CARD, highlightbackground=BORDA, highlightthickness=0)
        painel.pack(fill="x", padx=28, pady=(0, 10))

        tk.Label(
            painel,
            text="Resumo operacional",
            bg=BG_CARD,
            fg=TEXTO,
            font=("Arial", 12, "bold"),
        ).pack(anchor="w", padx=16, pady=(16, 6))

        tk.Label(
            painel,
            text="● atualização em tempo real",
            bg=BG_CARD,
            fg="#67d28d",
            font=("Arial", 8, "bold"),
        ).pack(anchor="w", padx=16, pady=(0, 8))

        tk.Label(
            painel,
            text="Acompanhe solicitações, pendências e evolução do histórico em tempo real.",
            bg=BG_CARD,
            fg=TEXTO_MUTED,
            font=FONT_SUBTITULO,
            justify="left",
        ).pack(anchor="w", padx=16, pady=(0, 16))

        topo_barra = tk.Frame(frame, bg=BG)
        topo_barra.pack(fill="x", padx=28, pady=(0, 10))

        self.botao_flat(topo_barra, "Atualizar", self._abrir_dashboard).pack(side="left")

        # Botões de Exportação para Excel e Ações
        self.botao_flat(topo_barra, "Exportar Dashboard", self._exportar_dashboard).pack(side="left", padx=(6, 0))
        self.botao_flat(topo_barra, "Exportar Registros", self._exportar_registros).pack(side="left", padx=(4, 0))
        self.botao_flat(topo_barra, "Esvaziar Lixeira", self._esvaziar_lixeira_confirm).pack(side="left", padx=(4, 0))

        for texto, valor in (("Todos", "todos"), ("Em andamento", "on going"), ("Finalizadas", "finalizadas"), ("🗑️ Lixeira", "lixeira")):
            btn = tk.Button(
                topo_barra,
                text=texto,
                command=lambda v=valor: self._aplicar_filtro_dashboard(v),
                bg=BG_CARD,
                fg=TEXTO,
                activebackground=BORDA,
                activeforeground=TEXTO,
                relief="solid",
                bd=1,
                font=("Arial", 9, "bold"),
                padx=10,
                pady=6,
                cursor="hand2",
            )
            btn.pack(side="left", padx=(12, 4) if texto == "Todos" else 4)

        # Barra de Pesquisa
        busca_frame = tk.Frame(topo_barra, bg=BG)
        busca_frame.pack(side="left", padx=(16, 0))

        tk.Label(
            busca_frame, text="Buscar:", bg=BG, fg=TEXTO_MUTED, font=("Arial", 9, "bold")
        ).pack(side="left", padx=(0, 6))

        self.campo_busca = tk.Entry(
            busca_frame,
            bg=BG_CARD,
            fg=TEXTO,
            insertbackground=TEXTO,
            relief="solid",
            font=("Arial", 9),
            highlightbackground=BORDA,
            highlightthickness=1,
            bd=0,
            width=20,
        )
        self.campo_busca.pack(side="left", ipady=4)
        if hasattr(self, "_dashboard_search_query") and self._dashboard_search_query:
            self.campo_busca.insert(0, self._dashboard_search_query)
        self.campo_busca.bind("<KeyRelease>", lambda e: self._filtrar_busca_dashboard())

        historico = self._carregar_historico()
        self._dashboard_last_signature = self._dashboard_signature(historico)
        self._renderizar_dashboard_conteudo(frame, historico)
        self.root.after(2000, self._check_dashboard_refresh)

    def _filtrar_busca_dashboard(self):
        if hasattr(self, "campo_busca"):
            self._dashboard_search_query = self.campo_busca.get().strip().lower()
        if FRAME_DASHBOARD in self.frames:
            historico = self._carregar_historico()
            self._renderizar_dashboard_conteudo(self.frames[FRAME_DASHBOARD], historico)

    def _aplicar_filtro_dashboard(self, valor):
        self._dashboard_filter = valor
        if FRAME_DASHBOARD in self.frames:
            historico = self._carregar_historico()
            self._renderizar_dashboard_conteudo(self.frames[FRAME_DASHBOARD], historico)

    def _esvaziar_lixeira_confirm(self):
        if not self._validar_permissao("admin"):
            return
        res = messagebox.askyesno(
            "Esvaziar Lixeira",
            "Tem certeza que deseja EXCLUIR PERMANENTEMENTE todos os registros da lixeira?\n\nEsta ação não poderá ser desfeita.",
            parent=self.root
        )
        if res:
            qtd = esvaziar_lixeira()
            messagebox.showinfo("Lixeira Esvaziada", f"{qtd} registro(s) cancelado(s) foi/foram excluído(s) permanentemente do banco.")
            self._abrir_dashboard()

    def _renderizar_dashboard_conteudo(self, frame, historico):
        for widget in frame.winfo_children():
            if getattr(widget, "_dashboard_area", False):
                widget.destroy()

        counts = contar_por_status()
        total_metrica = counts["total"]
        em_andamento_metrica = counts["em_andamento"]
        finalizadas_metrica = counts["finalizadas"]
        canceladas_metrica = counts["canceladas"]

        stats = tk.Frame(frame, bg=BG)
        stats._dashboard_area = True
        stats.pack(fill="x", padx=28, pady=(0, 8))

        self._stat_card(stats, total_metrica, "Total de solicitações")
        self._stat_card(stats, em_andamento_metrica, "Em andamento")
        self._stat_card(stats, finalizadas_metrica, "Finalizadas")
        self._stat_card(stats, canceladas_metrica, "Canceladas")

        # Gráfico de Pizza
        grafico_frame = tk.Frame(frame, bg=BG)
        grafico_frame._dashboard_area = True
        grafico_frame.pack(pady=(10, 15))

        tk.Label(
            grafico_frame,
            text="VISÃO GERAL DAS SOLICITAÇÕES",
            bg=BG,
            fg=ACCENT,
            font=("Arial", 10, "bold")
        ).pack(pady=(0, 10))

        v_fin = finalizadas_metrica
        v_and = em_andamento_metrica
        v_can = canceladas_metrica

        valores_pie = [v_fin, v_and, v_can]
        cores_pie = ["#67d28d", "#ffbf69", "#ff6b6b"]

        if sum(valores_pie) == 0:
            valores_pie = [1, 0, 0]

        import matplotlib.pyplot as plt
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

        fig, ax = plt.subplots(figsize=(3.5, 3.5), facecolor=BG)
        ax.set_facecolor(BG)

        wedges, texts, autotexts = ax.pie(
            valores_pie,
            colors=cores_pie,
            startangle=90,
            autopct="%1.0f%%" if sum(valores_pie) > 0 else "",
            pctdistance=0.75,
            wedgeprops={"width": 0.40, "edgecolor": BG, "linewidth": 2},
            textprops={"color": "white", "fontsize": 9, "fontweight": "bold"}
        )

        ax.text(
            0, 0,
            f"{total_metrica}\nTOTAL",
            color="white", ha="center", va="center",
            fontsize=14, fontweight="bold"
        )
        ax.set_aspect("equal")

        canvas = FigureCanvasTkAgg(fig, master=grafico_frame)
        canvas.get_tk_widget().configure(bg=BG, highlightthickness=0, bd=0)
        canvas.draw()
        canvas.get_tk_widget().pack()
        plt.close(fig)  # libera memória da figura

        self.linha_divisoria(frame)

        tk.Label(
            frame,
            text=espacar("Registros"),
            bg=BG,
            fg=TEXTO_MUTED,
            font=FONT_CAPTION,
        ).pack(anchor="w", padx=28, pady=(0, 8))

        container = tk.Frame(frame, bg=BG)
        container._dashboard_area = True
        container.pack(fill="both", padx=28, pady=(0, 16))

        filtro = (self._dashboard_filter or "todos").strip().lower()
        if filtro == "on going":
            lista = [i for i in historico if str(i.get("status", "")).upper() == "ON GOING" and not i.get("cancelado")]
        elif filtro == "finalizadas":
            lista = [i for i in historico if str(i.get("status", "")).upper() not in ("ON GOING", "CANCELADO") and not i.get("cancelado")]
        elif filtro in ("lixeira", "canceladas"):
            lista = [i for i in historico if i.get("cancelado") or str(i.get("status", "")).upper() == "CANCELADO"]
        else:  # "todos" — exibe apenas solicitações ativas (não canceladas)
            lista = [i for i in historico if not i.get("cancelado") and str(i.get("status", "")).upper() != "CANCELADO"]

        query = getattr(self, "_dashboard_search_query", "").strip().lower()
        if query:
            lista = [
                item for item in lista
                if query in str(item.get("ticket", "")).lower()
                or query in str(item.get("pn", "")).lower()
                or query in str(item.get("part_number", "")).lower()
                or query in str(item.get("mpn", "")).lower()
                or query in str(item.get("description", "")).lower()
                or query in str(item.get("solicitante", "")).lower()
                or query in str(item.get("requisitante", "")).lower()
                or query in str(item.get("produto", "")).lower()
            ]

        if not lista:
            msg = "Nenhum registro encontrado para a busca informada." if query else "Nenhum registro encontrado para este filtro."
            tk.Label(container, text=msg, bg=BG, fg=TEXTO_MUTED, font=FONT_SUBTITULO).pack(anchor="w", pady=10)
            return

        recentes = list(reversed(lista))[:15]
        for item in recentes:
            self._linha_historico(container, item)

    def _dashboard_signature(self, historico):
        return tuple(
            (
                str(item.get("id", "")),
                str(item.get("linha", "")),
                str(item.get("status", "")),
                str(item.get("description", "")),
                str(item.get("pn") or item.get("part_number") or ""),
                str(item.get("ultima_alteracao", "")),
            )
            for item in historico
        )

    def _check_dashboard_refresh(self):
        if self.history and self.history[-1] != FRAME_DASHBOARD:
            return
        if FRAME_DASHBOARD not in self.frames:
            return

        db_path = get_db_path()
        try:
            current_mtime = os.path.getmtime(db_path) if os.path.isfile(db_path) else 0
        except Exception:
            current_mtime = 0

        last_mtime = getattr(self, "_dashboard_db_mtime", None)
        if last_mtime != current_mtime:
            self._dashboard_db_mtime = current_mtime
            historico = self._carregar_historico()
            signature = self._dashboard_signature(historico)
            if signature != getattr(self, "_dashboard_last_signature", None):
                self._dashboard_last_signature = signature
                self._renderizar_dashboard_conteudo(self.frames[FRAME_DASHBOARD], historico)

        self.root.after(2000, self._check_dashboard_refresh)

    def _stat_card(self, parent, valor, legenda):
        card = tk.Frame(parent, bg=BG_CARD, highlightbackground=BORDA, highlightthickness=0)
        card.pack(side="left", expand=True, fill="x", padx=(0, 6))

        tk.Label(
            card, text=str(valor), bg=BG_CARD, fg=TEXTO, font=("Arial", 18, "bold")
        ).pack(pady=(8, 0))

        tk.Label(
            card, text=espacar(legenda), bg=BG_CARD, fg=TEXTO_MUTED, font=FONT_CAPTION
        ).pack(pady=(2, 8))

    def _linha_historico(self, parent, item):
        perm = str(self.permissao or "").upper()

        card = tk.Frame(
            parent, bg=BG_CARD, cursor="hand2",
            highlightbackground=BORDA, highlightthickness=1, bd=0
        )
        card.pack(fill="x", padx=0, pady=6)

        corpo = tk.Frame(card, bg=BG_CARD, padx=14, pady=12)
        corpo.pack(fill="both", expand=True)

        topo = tk.Frame(corpo, bg=BG_CARD)
        topo.pack(fill="x", pady=(0, 6))

        btn_area = tk.Frame(topo, bg=BG_CARD)
        btn_area.pack(side="right", anchor="ne", padx=(10, 0))

        status_raw = str(item.get("status", "—")).strip().upper()
        cancelado = bool(item.get("cancelado")) or status_raw == "CANCELADO"
        status_on_going = (status_raw == "ON GOING" and not cancelado)

        if cancelado:
            badge_bg = "#3a1414"
            badge_fg = "#ff6b6b"
            badge_texto = "CANCELADA"
        elif status_on_going:
            badge_bg = "#3a2d14"
            badge_fg = "#ffbf69"
            badge_texto = "EM ANDAMENTO"
        else:
            badge_bg = "#163321"
            badge_fg = "#67d28d"
            badge_texto = status_raw or "FINALIZADA"

        badge = tk.Label(
            btn_area, text=badge_texto, bg=badge_bg, fg=badge_fg,
            font=("Arial", 8, "bold"), padx=8, pady=4
        )
        badge.pack(side="right", padx=(6, 0))

        # Ações baseadas no perfil
        if perm in ("ADMIN", "ENGENHARIA") and not cancelado:
            btn_editar = tk.Button(
                btn_area, text="Editar",
                command=lambda item=item: self._abrir_editar_registro(item),
                bg=BG_CARD, fg=ACCENT, activebackground=BORDA, activeforeground=TEXTO,
                relief="solid", bd=1, font=("Arial", 8, "bold"), padx=8, pady=3, cursor="hand2"
            )
            btn_editar.pack(side="right", padx=(6, 0))

        if perm == "ADMIN" and not cancelado:
            btn_excluir = tk.Button(
                btn_area, text="Excluir",
                command=lambda item=item: self._excluir_registro(item),
                bg=BG_CARD, fg="#F30808", activebackground=BORDA, activeforeground=TEXTO,
                relief="solid", bd=1, font=("Arial", 8, "bold"), padx=8, pady=3, cursor="hand2"
            )
            btn_excluir.pack(side="right", padx=(6, 0))

        if perm in ("ADMIN", "ENGENHARIA") and not cancelado:
            btn_atualizar_pn = tk.Button(
                btn_area, text="Atualizar Flex PN",
                command=lambda item=item, b=None: self._atualizar_pn_dashboard(item, b),
                bg=BG_CARD, fg=ACCENT, activebackground=BORDA, activeforeground=TEXTO,
                relief="solid", bd=1, font=("Arial", 8, "bold"), padx=8, pady=3, cursor="hand2"
            )
            btn_atualizar_pn.pack(side="right", padx=(6, 0))

            phoenix_fechado = bool(str(item.get("phoenix_fechamento") or "").strip())
            pegasus_iniciado = bool(str(item.get("pegasus_abertura") or "").strip())
            pegasus_finalizado = bool(str(item.get("pegasus_fechamento") or "").strip())
            custo_iniciado = bool(str(item.get("custo_abertura") or "").strip())
            custo_finalizado = bool(str(item.get("custo_fechamento") or "").strip())

            if not phoenix_fechado:
                tk.Button(
                    btn_area, text="Fechar Phoenix",
                    command=lambda item=item: self._finalizar_phoenix(item),
                    bg=BG_CARD, fg=ACCENT, relief="solid", bd=1, font=("Arial", 8, "bold"), padx=6, pady=2
                ).pack(side="left", padx=2)

            if phoenix_fechado and not pegasus_iniciado:
                tk.Button(
                    btn_area, text="Iniciar Pegasus",
                    command=lambda item=item: self._iniciar_pegasus(item),
                    bg=BG_CARD, fg=ACCENT, relief="solid", bd=1, font=("Arial", 8, "bold"), padx=6, pady=2
                ).pack(side="left", padx=2)

            if pegasus_iniciado and not pegasus_finalizado:
                tk.Button(
                    btn_area, text="Finalizar Pegasus",
                    command=lambda item=item: self._finalizar_pegasus(item),
                    bg=BG_CARD, fg=ACCENT, relief="solid", bd=1, font=("Arial", 8, "bold"), padx=6, pady=2
                ).pack(side="left", padx=2)

            if pegasus_finalizado and not custo_iniciado:
                tk.Button(
                    btn_area, text="Iniciar Custo",
                    command=lambda item=item: self._iniciar_custo(item),
                    bg=BG_CARD, fg=ACCENT, relief="solid", bd=1, font=("Arial", 8, "bold"), padx=6, pady=2
                ).pack(side="left", padx=2)

            if custo_iniciado and not custo_finalizado:
                tk.Button(
                    btn_area, text="Finalizar Custo",
                    command=lambda item=item: self._finalizar_custo(item),
                    bg=BG_CARD, fg=ACCENT, relief="solid", bd=1, font=("Arial", 8, "bold"), padx=6, pady=2
                ).pack(side="left", padx=2)

        ticket = str(item.get("ticket") or "").strip()
        descricao = str(item.get("description") or "Sem descrição").strip()
        titulo_texto = f"{ticket} - {descricao}" if ticket else descricao

        lbl_titulo = tk.Label(
            topo, text=titulo_texto, bg=BG_CARD, fg=TEXTO,
            font=("Arial", 10, "bold"), justify="left", anchor="w", wraplength=480
        )
        lbl_titulo.pack(side="left", fill="x", expand=True)

        div = tk.Frame(corpo, bg=BORDA, height=1)
        div.pack(fill="x", pady=(4, 10))

        info_frame = tk.Frame(corpo, bg=BG_CARD)
        info_frame.pack(fill="x")

        solicitante = str(item.get("solicitante") or "").strip() or "Não informado"
        requisitante = str(item.get("requisitante") or "").strip() or "Não informado"
        data_abert = item.get("data_abertura", "—")
        hora_abert = item.get("hora_abertura", "")
        data_str = f"{data_abert} {hora_abert}".strip()

        tk.Label(
            info_frame, text=f"Solicitante: {solicitante}   •   Requisitante: {requisitante}",
            bg=BG_CARD, fg=TEXTO_MUTED, font=("Arial", 9)
        ).pack(anchor="w", pady=(0, 2))

        tk.Label(
            info_frame, text=f"Data de Abertura: {data_str}",
            bg=BG_CARD, fg=TEXTO_MUTED, font=("Arial", 9)
        ).pack(anchor="w", pady=(0, 2))

        pn_val = str(item.get("pn") or item.get("part_number") or "").strip()
        pn_row = tk.Frame(info_frame, bg=BG_CARD)
        pn_row.pack(anchor="w", pady=(2, 2))

        tk.Label(pn_row, text="PN Capturado: ", bg=BG_CARD, fg=TEXTO_MUTED, font=("Arial", 9, "bold")).pack(side="left")
        tk.Label(pn_row, text=pn_val or "Não capturado", bg=BG_CARD, fg=ACCENT if pn_val else TEXTO_MUTED, font=("Arial", 9, "bold" if pn_val else "italic")).pack(side="left")

        criador = str(item.get("criado_por") or item.get("user") or "—").strip()
        criado_em = str(item.get("criado_em") or data_str).strip()
        ultima_alt = str(item.get("ultima_alteracao") or criado_em).strip()
        usr_alt = str(item.get("ultimo_usuario_alterou") or criador).strip()

        tk.Label(
            info_frame,
            text=f"Criado por: {criador} ({criado_em})   •   Última alteração: {ultima_alt} por {usr_alt}",
            bg=BG_CARD, fg="#8c8c8c", font=("Arial", 8)
        ).pack(anchor="w", pady=(2, 0))

        def mostrar_detalhes(e=None):
            hist_log = "\n".join(
                f"  • {h.get('data')} - {h.get('descricao')} ({h.get('usuario')})"
                for h in item.get("historico_alteracoes", [])
            ) or "  • Nenhum histórico registrado."

            messagebox.showinfo(
                "Detalhes do Registro",
                f"ID: {item.get('id', '—')}\n"
                f"Linha: {item.get('linha', '—')}\n"
                f"Ticket: {item.get('ticket') or '—'}\n\n"
                f"Description: {item.get('description', '—')}\n"
                f"Produto: {item.get('produto') or 'Não informado'}\n"
                f"MPN: {item.get('mpn') or 'Não informado'}\n\n"
                f"Status: {item.get('status', '—')}\n"
                f"PN Capturado: {pn_val or '—'}\n"
                f"Origem: {item.get('origem') or 'Não informado'}\n"
                f"Custo: {item.get('custo') or 'Não informado'}\n\n"
                f"Solicitante: {solicitante}\n"
                f"Requisitante: {requisitante}\n\n"
                f"--- AUDITORIA ---\n"
                f"Criado por: {criador} em {criado_em}\n"
                f"Última alteração: {ultima_alt} por {usr_alt}\n\n"
                f"--- HISTÓRICO DE ALTERAÇÕES ---\n{hist_log}"
            )

        card.bind("<Enter>", lambda e: card.configure(highlightbackground=ACCENT))
        card.bind("<Leave>", lambda e: card.configure(highlightbackground=BORDA))

        for widget in (corpo, topo, info_frame, lbl_titulo):
            widget.bind("<Button-1>", mostrar_detalhes)

    # Exportar Dashboard para Excel
    def _exportar_dashboard(self):
        if not self._validar_permissao("exportar"):
            return

        historico = self._carregar_historico()
        counts = contar_por_status()
        data_str = datetime.now().strftime("%d-%m-%Y")
        nome_arquivo = f"Phoenix_Dashboard_{data_str}.xlsx"

        caminho = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=nome_arquivo,
            filetypes=[("Excel Files", "*.xlsx")],
            title="Exportar Dashboard"
        )
        if not caminho:
            return

        try:
            wb = openpyxl.Workbook()
            # Aba Resumo
            ws_resumo = wb.active
            ws_resumo.title = "Resumo"

            header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
            header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")

            ws_resumo.append(["Métrica", "Quantidade"])
            ws_resumo.append(["Total de Solicitações", counts["total"]])
            ws_resumo.append(["Em Andamento", counts["em_andamento"]])
            ws_resumo.append(["Finalizadas", counts["finalizadas"]])
            ws_resumo.append(["Canceladas", counts["canceladas"]])

            for cell in ws_resumo[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center")

            for row in ws_resumo.iter_rows(min_row=2, max_col=2):
                row[0].font = Font(name="Arial", size=10, bold=True)
                row[1].font = Font(name="Arial", size=10)
                row[1].alignment = Alignment(horizontal="center")

            # Aba Indicadores
            ws_ind = wb.create_sheet(title="Indicadores")
            ws_ind.append(["Status", "Quantidade", "Percentual"])
            total = counts["total"] or 1
            ws_ind.append(["Em Andamento", counts["em_andamento"], f"{(counts['em_andamento']/total)*100:.1f}%"])
            ws_ind.append(["Finalizadas", counts["finalizadas"], f"{(counts['finalizadas']/total)*100:.1f}%"])
            ws_ind.append(["Canceladas", counts["canceladas"], f"{(counts['canceladas']/total)*100:.1f}%"])

            for cell in ws_ind[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center")

            # Auto ajustar largura
            for sheet in (ws_resumo, ws_ind):
                for col in sheet.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = get_column_letter(col[0].column)
                    sheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

            wb.save(caminho)
            messagebox.showinfo("Sucesso", f"Dashboard exportado com sucesso:\n{caminho}")
        except Exception as exc:
            logger.exception("Erro ao exportar dashboard")
            messagebox.showerror("Erro", f"Falha ao exportar dashboard: {exc}")

    # Exportar Registros para Excel
    def _exportar_registros(self):
        if not self._validar_permissao("exportar"):
            return

        # Perguntar se deseja exportar todos ou filtrados
        opcao = messagebox.askyesnocancel(
            "Exportar Registros",
            "Deseja exportar TODOS os registros?\n\nClique 'Sim' para todos ou 'Não' para exportar apenas os registros do filtro/busca atual."
        )
        if opcao is None:
            return

        historico = self._carregar_historico()

        if not opcao:  # Filtrados
            filtro = (self._dashboard_filter or "todos").strip().lower()
            if filtro == "on going":
                historico = [i for i in historico if str(i.get("status", "")).upper() == "ON GOING" and not i.get("cancelado")]
            elif filtro == "finalizadas":
                historico = [i for i in historico if str(i.get("status", "")).upper() not in ("ON GOING", "CANCELADO") and not i.get("cancelado")]
            elif filtro == "canceladas":
                historico = [i for i in historico if i.get("cancelado") or str(i.get("status", "")).upper() == "CANCELADO"]

            query = getattr(self, "_dashboard_search_query", "").strip().lower()
            if query:
                historico = [
                    item for item in historico
                    if query in str(item.get("ticket", "")).lower()
                    or query in str(item.get("pn", "")).lower()
                    or query in str(item.get("part_number", "")).lower()
                    or query in str(item.get("mpn", "")).lower()
                    or query in str(item.get("description", "")).lower()
                    or query in str(item.get("solicitante", "")).lower()
                ]

        data_str = datetime.now().strftime("%d-%m-%Y")
        nome_arquivo = f"Phoenix_Registros_{data_str}.xlsx"

        caminho = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=nome_arquivo,
            filetypes=[("Excel Files", "*.xlsx")],
            title="Exportar Registros"
        )
        if not caminho:
            return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Registros"

            colunas = [
                "Ticket", "Description", "Produto", "MPN", "PN", "Solicitante",
                "Data Abertura Phoenix", "Data Abertura Pegasus", "Data Abertura Cost",
                "Data Fechamento"
            ]
            ws.append(colunas)

            # Estilo do cabeçalho
            header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
            header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            thin_border = Border(
                left=Side(style='thin', color='D9D9D9'),
                right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'),
                bottom=Side(style='thin', color='D9D9D9')
            )

            for item in historico:
                ticket = str(item.get("ticket") or "")
                desc = str(item.get("description") or "")
                prod = str(item.get("produto") or "")
                mpn = str(item.get("mpn") or "")
                pn = str(item.get("pn") or item.get("part_number") or "")
                solic = str(item.get("solicitante") or "")
                dt_phx = str(item.get("phoenix_data_abertura") or item.get("data_abertura") or "")
                dt_peg = str(item.get("pegasus_data_abertura") or item.get("data_abertura") or "")
                dt_cost = str(item.get("custo_data_abertura") or item.get("data_abertura") or "")
                dt_fech = str(item.get("data_fechamento") or "")

                row_vals = [
                    ticket, desc, prod, mpn, pn, solic,
                    dt_phx, dt_peg, dt_cost, dt_fech
                ]
                ws.append(row_vals)

                row_idx = ws.max_row
                for col_idx in range(1, len(colunas) + 1):
                    ws.cell(row=row_idx, column=col_idx).border = thin_border

            # Auto ajuste de colunas
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 50)

            # Filtro automático e congelamento de primeira linha
            ws.auto_filter.ref = ws.dimensions
            ws.freeze_panes = "A2"

            wb.save(caminho)
            messagebox.showinfo("Sucesso", f"Registros exportados com sucesso ({len(historico)} linhas):\n{caminho}")
        except Exception as exc:
            logger.exception("Erro ao exportar registros")
            messagebox.showerror("Erro", f"Falha ao exportar registros: {exc}")

    def _atualizar_pn_dashboard(self, item, botao=None):
        if not self._validar_permissao("atualizar_pn"):
            return
        ticket = str(item.get("ticket") or "").strip()
        if not ticket:
            messagebox.showwarning("Atualizar Flex PN", "Este registro não possui ticket.")
            return

        if botao is not None and botao.winfo_exists():
            botao.configure(state="disabled", text="Buscando...")

        def _worker():
            from automocoes.phoenix.atualizar_pn import buscar_pn_por_ticket
            resultado = buscar_pn_por_ticket(ticket)
            self.root.after(0, lambda: self._finalizar_atualizacao_pn(resultado, botao))

        threading.Thread(target=_worker, daemon=True).start()

    def _finalizar_atualizacao_pn(self, resultado, botao=None):
        if botao is not None:
            try:
                if botao.winfo_exists():
                    botao.configure(state="normal", text="Atualizar Flex PN")
            except Exception:
                pass

        messagebox.showinfo("Atualizar Flex PN", resultado.get("mensagem") or "Busca concluída.")

        if FRAME_DASHBOARD in self.frames:
            historico = self._carregar_historico()
            self._dashboard_last_signature = self._dashboard_signature(historico)
            self._renderizar_dashboard_conteudo(self.frames[FRAME_DASHBOARD], historico)

    def _abrir_editar_registro(self, item):
        if not self._validar_permissao("editar"):
            return

        if str(item.get("status", "")).upper() == "ENCERRADO":
            messagebox.showwarning("Registro Encerrado", "Este registro está encerrado e não pode ser editado.")
            return

        janela = tk.Toplevel(self.root)
        janela.title("Editar informações do produto")
        janela.configure(bg=BG)
        janela.attributes("-topmost", True)
        janela.resizable(False, False)
        janela.transient(self.root)

        corpo = tk.Frame(janela, bg=BG, padx=16, pady=16)
        corpo.pack(fill="both", expand=True)

        tk.Label(
            corpo,
            text=str(item.get("description") or "—").upper(),
            bg=BG, fg=TEXTO, font=("Arial", 11, "bold"),
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 14))

        campos_def = [
            ("Produto", "produto"),
            ("Solicitante", "solicitante"),
            ("Requisitante", "requisitante"),
            ("PN (Part Number)", "pn"),
            ("MPN", "mpn"),
            ("Origem", "origem"),
            ("Custo", "custo"),
        ]

        entries = {}
        for r_idx, (label_txt, campo_key) in enumerate(campos_def, start=1):
            tk.Label(corpo, text=espacar(label_txt), bg=BG, fg=TEXTO_MUTED, font=FONT_CAPTION).grid(row=r_idx, column=0, sticky="w", pady=(0, 6))
            e = self.campo_entry(corpo)
            e.grid(row=r_idx, column=1, sticky="ew", pady=(0, 6), padx=(10, 0), ipady=4)
            val = str(item.get(campo_key) or "")
            if not val and campo_key == "origem":
                val = "IMPORTADO"
            e.insert(0, val)
            entries[campo_key] = e

        var_custo_fin = tk.BooleanVar(value=bool(item.get("custo_finalizado")))
        tk.Label(corpo, text=espacar("Custo Finalizado"), bg=BG, fg=TEXTO_MUTED, font=FONT_CAPTION).grid(row=8, column=0, sticky="w", pady=(0, 6))
        chk_custo = tk.Checkbutton(corpo, variable=var_custo_fin, bg=BG, fg=TEXTO, activebackground=BG, selectcolor=BG_CARD, relief="flat")
        chk_custo.grid(row=8, column=1, sticky="w", pady=(0, 6), padx=(10, 0))

        corpo.grid_columnconfigure(1, weight=1)

        botoes = tk.Frame(corpo, bg=BG)
        botoes.grid(row=9, column=0, columnspan=2, pady=(16, 0))

        def salvar():
            novos = {key: e.get().strip() for key, e in entries.items()}
            novos["part_number"] = novos.get("pn", "")
            novos["custo_finalizado"] = var_custo_fin.get()

            atualizar_campos_registro(item.get("linha"), novos, usuario_alteracao=self.usuario)
            item.update(novos)
            janela.destroy()
            if FRAME_DASHBOARD in self.frames:
                historico = self._carregar_historico()
                self._dashboard_last_signature = self._dashboard_signature(historico)
                self._renderizar_dashboard_conteudo(self.frames[FRAME_DASHBOARD], historico)

        self.botao_flat(botoes, "Salvar", salvar, largura=12).pack(side="left", padx=6)
        self.botao_flat(botoes, "Cancelar", janela.destroy, largura=12).pack(side="left", padx=6)

        janela.update_idletasks()
        largura, altura = janela.winfo_width(), janela.winfo_height()
        x = (janela.winfo_screenwidth() // 2) - (largura // 2)
        y = (janela.winfo_screenheight() // 2) - (altura // 2)
        janela.geometry(f"+{x}+{y}")
        janela.grab_set()

    # =========================================================================
    # MENUS PHOENIX & PEGASUS
    # =========================================================================

    def _abrir_menu_phoenix(self):
        if not self._validar_permissao("phoenix"):
            return
        self._build_menu_phoenix()
        self.mostrar(FRAME_MENU_PHOENIX)

    def _executar_phoenix(self, arg=None):
        if not self._validar_permissao("phoenix"):
            return
        cred = self._obter_ou_pedir_credencial("PHOENIX")
        if not cred:
            return
        executar_script("automocoes", "phoenix", "phoenix.py", arg=arg, cred_user=cred["login"], cred_pass=cred["senha"])

    def _build_menu_phoenix(self):
        if FRAME_MENU_PHOENIX in self.frames:
            self.frames[FRAME_MENU_PHOENIX].destroy()

        frame = tk.Frame(self.container, bg=BG)
        self.frames[FRAME_MENU_PHOENIX] = frame

        self.build_menu_automacao(
            frame,
            "Phoenix",
            [
                ("Home Phoenix", lambda: self._executar_phoenix(arg="home"), "Retorna ao fluxo inicial"),
                ("Nova solicitação", lambda: self._executar_phoenix(), "Inicia uma nova solicitação"),
                ("Atualizar Flex PN", self._abrir_atualizar_pn_phoenix, "Busca e atualiza o PN"),
                ("Importar solicitações existentes", lambda: self._executar_phoenix(arg="importar"), "Traz tickets já criados"),
            ],
        )

    def _abrir_menu_pegasus(self):
        if not self._validar_permissao("pegasus"):
            return
        self._build_menu_pegasus()
        self.mostrar(FRAME_MENU_PEGASUS)

    def _executar_pegasus(self, arg=None):
        if not self._validar_permissao("pegasus"):
            return
        cred = self._obter_ou_pedir_credencial("PEGASUS")
        if not cred:
            return
        executar_script("automocoes", "pegasus", "pegasus.py", arg=arg, cred_user=cred["login"], cred_pass=cred["senha"])

    def _build_menu_pegasus(self):
        if FRAME_MENU_PEGASUS in self.frames:
            self.frames[FRAME_MENU_PEGASUS].destroy()

        frame = tk.Frame(self.container, bg=BG)
        self.frames[FRAME_MENU_PEGASUS] = frame

        self.build_menu_automacao(
            frame,
            "Pegasus",
            [
                ("Home Pegasus", lambda: self._executar_pegasus(arg="home"), "Retorna ao fluxo inicial"),
                ("Nova solicitação", lambda: self._executar_pegasus(), "Inicia uma nova solicitação"),
            ],
        )

    def _abrir_atualizar_pn_phoenix(self):
        if not self._validar_permissao("atualizar_pn"):
            return
        self._build_atualizar_pn_phoenix()
        self.mostrar(FRAME_ATUALIZAR_PN)

    def _build_atualizar_pn_phoenix(self):
        if FRAME_ATUALIZAR_PN in self.frames:
            self.frames[FRAME_ATUALIZAR_PN].destroy()

        frame = tk.Frame(self.container, bg=BG)
        self.frames[FRAME_ATUALIZAR_PN] = frame

        self.botao_voltar(frame)
        self.cabecalho(frame, "Phoenix", "Buscar PN")

        bloco = tk.Frame(frame, bg=BG)
        bloco.pack(padx=60, fill="x")

        tk.Label(
            bloco, text=espacar("Linha da planilha"), bg=BG, fg=TEXTO_MUTED, font=FONT_CAPTION
        ).pack(anchor="w", pady=(10, 6))

        campo_linha = self.campo_entry(bloco)
        campo_linha.pack(fill="x", ipady=6)

        def buscar():
            linha = campo_linha.get().strip()
            item = encontrar_por_linha(linha)
            if item is None:
                messagebox.showwarning("Phoenix", "Linha não encontrada.")
                return
            messagebox.showinfo(
                "Phoenix",
                f"Linha: {item.get('linha', '—')}\n"
                f"Ticket: {item.get('ticket') or '—'}\n"
                f"Description: {item.get('description', '—')}\n"
                f"Status: {item.get('status', '—')}\n"
                f"Data: {item.get('data_abertura', '—')}"
            )

        def rodar_automacao():
            linha = campo_linha.get().strip()
            if not linha:
                messagebox.showwarning("Phoenix", "Digite a linha.")
                return
            item = encontrar_por_linha(linha)
            if item is None:
                messagebox.showwarning("Phoenix", "Linha não encontrada no histórico.")
                return

            cred = self._obter_ou_pedir_credencial("PHOENIX")
            if not cred:
                return
            executar_script("automocoes", "phoenix", "atualizar_pn.py", arg=linha, cred_user=cred["login"], cred_pass=cred["senha"])
            messagebox.showinfo("Phoenix", "Automação iniciada em segundo plano.")

        self.botao_flat(bloco, "Buscar Registro", buscar).pack(fill="x", pady=(14, 6))
        self.botao_flat(bloco, "Rodar Automação (PN)", rodar_automacao).pack(fill="x", pady=(0, 0))

    # =========================================================================
    # TELA MINHAS CREDENCIAIS
    # =========================================================================

    def _abrir_minhas_credenciais(self):
        if not self._validar_permissao("credenciais"):
            return
        self._build_minhas_credenciais()
        self.mostrar(FRAME_CREDENCIAIS)

    def _build_minhas_credenciais(self):
        if FRAME_CREDENCIAIS in self.frames:
            self.frames[FRAME_CREDENCIAIS].destroy()

        frame = tk.Frame(self.container, bg=BG)
        self.frames[FRAME_CREDENCIAIS] = frame

        self.botao_voltar(frame)
        self.cabecalho(frame, "Segurança", "Minhas Credenciais")

        painel = tk.Frame(frame, bg=BG_CARD, highlightbackground=BORDA, highlightthickness=0)
        painel.pack(fill="x", padx=28, pady=(0, 20))

        tk.Label(
            painel, text="Gerenciador de Credenciais Individuais", bg=BG_CARD, fg=TEXTO, font=("Arial", 12, "bold")
        ).pack(anchor="w", padx=16, pady=(16, 6))

        tk.Label(
            painel,
            text="Cadastre suas credenciais para automações Phoenix, Pegasus e Cost Request.\nEstas informações pertencem apenas ao seu perfil e não são compartilhadas com administradores.",
            bg=BG_CARD, fg=TEXTO_MUTED, font=FONT_SUBTITULO, justify="left"
        ).pack(anchor="w", padx=16, pady=(0, 16))

        tipos = [("PHOENIX", "Automação Phoenix (Portal)"), ("PEGASUS", "Automação Pegasus"), ("COST", "Automação Cost Request")]

        for tipo, desc in tipos:
            cred = obter_credencial(self.usuario_id, tipo) or {}
            login_val = cred.get("login") or "Não configurado"
            has_pass = bool(cred.get("senha"))

            card = tk.Frame(frame, bg=BG_CARD, highlightbackground=BORDA, highlightthickness=1)
            card.pack(fill="x", padx=28, pady=6)

            corpo = tk.Frame(card, bg=BG_CARD, padx=16, pady=12)
            corpo.pack(fill="x")

            lbl_t = tk.Label(corpo, text=f"CREDENCIAIS {tipo}", bg=BG_CARD, fg=ACCENT, font=("Arial", 10, "bold"))
            lbl_t.pack(anchor="w")

            lbl_d = tk.Label(corpo, text=desc, bg=BG_CARD, fg=TEXTO_MUTED, font=("Arial", 8))
            lbl_d.pack(anchor="w", pady=(0, 8))

            status_str = f"Login: {login_val}   •   Senha: {'••••••••' if has_pass else 'Não cadastrada'}"
            tk.Label(corpo, text=status_str, bg=BG_CARD, fg=TEXTO, font=("Arial", 9)).pack(anchor="w", pady=(0, 8))

            def _editar_cred(t=tipo):
                self._obter_ou_pedir_credencial(t, forcar=True)
                self._build_minhas_credenciais()
                self.mostrar(FRAME_CREDENCIAIS)

            btn = self.botao_flat(corpo, "Alterar Credenciais", _editar_cred, largura=18)
            btn.pack(anchor="w")

    # =========================================================================
    # TELA ÁREA ADMINISTRATIVA (Apenas ADMIN)
    # =========================================================================

    def _abrir_administracao(self):
        if not self._validar_permissao("admin"):
            return
        self._build_administracao()
        self.mostrar(FRAME_ADMIN)

    def _build_administracao(self):
        if not self._validar_permissao("admin"):
            return

        if FRAME_ADMIN in self.frames:
            self.frames[FRAME_ADMIN].destroy()

        frame = tk.Frame(self.container, bg=BG)
        self.frames[FRAME_ADMIN] = frame

        self.botao_voltar(frame)
        self.cabecalho(frame, "Administração", "Gerenciamento de Usuários")

        stats_db = obter_estatisticas_banco()

        # Painel Visual de Métricas do Banco de Dados
        db_panel = tk.Frame(frame, bg=BG)
        db_panel.pack(fill="x", padx=28, pady=(0, 12))

        # Card 1: Status do Banco
        c1 = tk.Frame(db_panel, bg=BG_CARD, highlightbackground=BORDA, highlightthickness=1)
        c1.pack(side="left", expand=True, fill="x", padx=(0, 6))
        c1_in = tk.Frame(c1, bg=BG_CARD, padx=12, pady=10)
        c1_in.pack(fill="x")
        tk.Label(c1_in, text=f"🗄️ {stats_db['nome_arquivo']}", bg=BG_CARD, fg=ACCENT, font=("Arial", 10, "bold")).pack(anchor="w")
        tk.Label(c1_in, text=f"Tamanho: {stats_db['tamanho_mb']} MB   •   Modo: {stats_db['modo_journal']}", bg=BG_CARD, fg=TEXTO, font=("Arial", 8)).pack(anchor="w", pady=(2, 0))
        tk.Label(c1_in, text=f"Última alteração: {stats_db['data_modificacao']}", bg=BG_CARD, fg=TEXTO_MUTED, font=("Arial", 7)).pack(anchor="w")

        # Card 2: Registros e Usuários
        c2 = tk.Frame(db_panel, bg=BG_CARD, highlightbackground=BORDA, highlightthickness=1)
        c2.pack(side="left", expand=True, fill="x", padx=3)
        c2_in = tk.Frame(c2, bg=BG_CARD, padx=12, pady=10)
        c2_in.pack(fill="x")
        tk.Label(c2_in, text=f"📊 Registros: {stats_db['total_solicitacoes']} solicitações", bg=BG_CARD, fg=ACCENT, font=("Arial", 10, "bold")).pack(anchor="w")
        tk.Label(c2_in, text=f"Usuários Cadastrados: {stats_db['total_usuarios']}   •   Credenciais: {stats_db['total_credenciais']}", bg=BG_CARD, fg=TEXTO, font=("Arial", 8)).pack(anchor="w", pady=(2, 0))
        tk.Label(c2_in, text="Permissão de Escrita: Administrador (Máquina Principal)", bg=BG_CARD, fg="#67d28d", font=("Arial", 7, "bold")).pack(anchor="w")

        # Card 3: Backups Diários
        c3 = tk.Frame(db_panel, bg=BG_CARD, highlightbackground=BORDA, highlightthickness=1)
        c3.pack(side="left", expand=True, fill="x", padx=(6, 0))
        c3_in = tk.Frame(c3, bg=BG_CARD, padx=12, pady=10)
        c3_in.pack(fill="x")
        tk.Label(c3_in, text=f"💾 Backups Diários: {stats_db['total_backups']} salvos", bg=BG_CARD, fg=ACCENT, font=("Arial", 10, "bold")).pack(anchor="w")
        tk.Label(c3_in, text="Backup diário automático gerado no fim do dia", bg=BG_CARD, fg=TEXTO, font=("Arial", 8)).pack(anchor="w", pady=(2, 0))
        tk.Label(c3_in, text=f"Pasta: /backups", bg=BG_CARD, fg=TEXTO_MUTED, font=("Arial", 7)).pack(anchor="w")

        top_bar = tk.Frame(frame, bg=BG)
        top_bar.pack(fill="x", padx=28, pady=(4, 16))

        self.botao_flat(top_bar, "+ Adicionar Usuário", self._dialog_adicionar_usuario).pack(side="left")
        self.botao_flat(top_bar, "Criar Backup Agora", self._executar_backup_manual).pack(side="left", padx=(4, 0))
        self.botao_flat(top_bar, "Alterar Banco (Rede)", self._alterar_local_banco).pack(side="left", padx=(4, 0))
        self.botao_flat(top_bar, "Exportar Usuários", self._exportar_usuarios).pack(side="left", padx=(4, 0))
        self.botao_flat(top_bar, "Atualizar Lista", self._abrir_administracao).pack(side="left", padx=(4, 0))

        usuarios = listar_usuarios()

        container = tk.Frame(frame, bg=BG)
        container.pack(fill="both", padx=28, pady=(0, 16))

        for u in usuarios:
            card = tk.Frame(container, bg=BG_CARD, highlightbackground=BORDA, highlightthickness=1)
            card.pack(fill="x", pady=4)

            corpo = tk.Frame(card, bg=BG_CARD, padx=14, pady=10)
            corpo.pack(fill="x")

            col_left = tk.Frame(corpo, bg=BG_CARD)
            col_left.pack(side="left", fill="x", expand=True)

            status_str = "ATIVO" if u["ativo"] else "DESATIVADO"
            status_col = "#67d28d" if u["ativo"] else "#ff6b6b"

            tk.Label(
                col_left, text=f"{u['nome']}  •  Código: {u['codigo']}",
                bg=BG_CARD, fg=TEXTO, font=("Arial", 10, "bold")
            ).pack(anchor="w")

            tk.Label(
                col_left, text=f"Permissão: {u['permissao']}   •   Status: {status_str}",
                bg=BG_CARD, fg=status_col, font=("Arial", 9)
            ).pack(anchor="w", pady=(2, 0))

            col_right = tk.Frame(corpo, bg=BG_CARD)
            col_right.pack(side="right")

            u_id = u["id"]

            btn_edit = tk.Button(
                col_right, text="Editar",
                command=lambda uid=u_id: self._dialog_editar_usuario(uid),
                bg=BG_CARD, fg=ACCENT, relief="solid", bd=1, font=("Arial", 8, "bold"), padx=6, pady=2, cursor="hand2"
            )
            btn_edit.pack(side="left", padx=3)

            if u["ativo"]:
                btn_toggle = tk.Button(
                    col_right, text="Desativar",
                    command=lambda uid=u_id: self._toggle_status_usuario(uid, True),
                    bg=BG_CARD, fg="#ff6b6b", relief="solid", bd=1, font=("Arial", 8, "bold"), padx=6, pady=2, cursor="hand2"
                )
            else:
                btn_toggle = tk.Button(
                    col_right, text="Reativar",
                    command=lambda uid=u_id: self._toggle_status_usuario(uid, False),
                    bg=BG_CARD, fg="#67d28d", relief="solid", bd=1, font=("Arial", 8, "bold"), padx=6, pady=2, cursor="hand2"
                )
            btn_toggle.pack(side="left", padx=3)

            btn_del = tk.Button(
                col_right, text="Excluir",
                command=lambda uid=u_id, unome=u["nome"]: self._excluir_usuario_confirm(uid, unome),
                bg=BG_CARD, fg="#ff4d4d", relief="solid", bd=1, font=("Arial", 8, "bold"), padx=6, pady=2, cursor="hand2"
            )
            btn_del.pack(side="left", padx=3)

    def _excluir_usuario_confirm(self, user_id: int, nome: str):
        if not self._validar_permissao("admin"):
            return
        if user_id == self.usuario_id:
            messagebox.showwarning("Aviso", "Você não pode excluir seu próprio usuário atualmente conectado.")
            return
        res = messagebox.askyesno(
            "Confirmar Exclusão",
            f"Tem certeza que deseja excluir permanentemente o usuário '{nome}'?",
            parent=self.root
        )
        if res:
            if excluir_usuario(user_id):
                messagebox.showinfo("Sucesso", f"Usuário '{nome}' excluído com sucesso.")
                self._abrir_administracao()
            else:
                messagebox.showerror("Erro", f"Falha ao excluir o usuário '{nome}'.")

    def _executar_backup_manual(self):
        if not self._validar_permissao("admin"):
            return
        dest = fazer_backup("manual_interface")
        if dest:
            verificar_e_criar_backup_diario()
            messagebox.showinfo("Backup Criado", f"Backup corporativo gerado com sucesso em:\n{dest}")
            self._abrir_administracao()
        else:
            messagebox.showerror("Erro", "Não foi possível criar o backup manual.")

    def _alterar_local_banco(self):
        if not self._validar_permissao("admin"):
            return
        novo_caminho = filedialog.asksaveasfilename(
            title="Selecionar / Definir Local do Banco de Dados (.db)",
            initialfile="phoenix_tool.db",
            filetypes=[("Banco SQLite", "*.db"), ("Todos os Arquivos", "*.*")]
        )
        if novo_caminho:
            try:
                set_db_path(novo_caminho)
                inicializar_banco()
                messagebox.showinfo("Sucesso", f"Caminho do banco de dados alterado com sucesso para:\n{novo_caminho}")
                self._abrir_administracao()
            except Exception as exc:
                messagebox.showerror("Erro", f"Não foi possível alterar o banco de dados:\n{exc}")

    def _exportar_usuarios(self):
        if not self._validar_permissao("admin"):
            return
        caminho = filedialog.asksaveasfilename(
            title="Exportar Cadastro de Usuários",
            initialfile="usuarios_phoenix_tool.json",
            defaultextension=".json",
            filetypes=[("Arquivo JSON", "*.json")]
        )
        if caminho:
            try:
                qtd = exportar_usuarios_json(caminho)
                messagebox.showinfo("Sucesso", f"{qtd} usuário(s) exportado(s) com sucesso para:\n{caminho}")
            except Exception as exc:
                messagebox.showerror("Erro", f"Erro ao exportar usuários:\n{exc}")

    def _importar_usuarios(self):
        if not self._validar_permissao("admin"):
            return
        caminho = filedialog.askopenfilename(
            title="Importar Cadastro de Usuários",
            filetypes=[("Arquivo JSON", "*.json")]
        )
        if caminho:
            try:
                imp, exis = importar_usuarios_json(caminho)
                messagebox.showinfo("Sucesso", f"Importação concluída!\n\nImportados com sucesso: {imp}\nJá existentes/ignorados: {exis}")
                self._abrir_administracao()
            except Exception as exc:
                messagebox.showerror("Erro", f"Erro ao importar usuários:\n{exc}")

    def _dialog_adicionar_usuario(self):
        if not self._validar_permissao("admin"):
            return

        janela = tk.Toplevel(self.root)
        janela.title("Adicionar Usuário")
        janela.configure(bg=BG)
        janela.attributes("-topmost", True)
        janela.resizable(False, False)

        corpo = tk.Frame(janela, bg=BG, padx=20, pady=20)
        corpo.pack(fill="both", expand=True)

        db_path_atual = get_db_path()

        tk.Label(corpo, text="Novo Usuário", bg=BG, fg=TEXTO, font=("Arial", 12, "bold")).pack(anchor="w", pady=(0, 2))
        tk.Label(corpo, text=f"Salvando no Banco: {os.path.basename(db_path_atual)}", bg=BG, fg=ACCENT, font=("Arial", 8, "italic")).pack(anchor="w", pady=(0, 10))

        tk.Label(corpo, text="Nome Completo:", bg=BG, fg=TEXTO_MUTED, font=FONT_CAPTION).pack(anchor="w", pady=(4, 2))
        e_nome = self.campo_entry(corpo)
        e_nome.pack(fill="x", ipady=6, pady=(0, 8))

        tk.Label(corpo, text="Código Engenharia:", bg=BG, fg=TEXTO_MUTED, font=FONT_CAPTION).pack(anchor="w", pady=(4, 2))
        e_cod = self.campo_entry(corpo)
        e_cod.pack(fill="x", ipady=6, pady=(0, 8))

        tk.Label(corpo, text="Permissão:", bg=BG, fg=TEXTO_MUTED, font=FONT_CAPTION).pack(anchor="w", pady=(4, 2))
        cb_perm = ttk.Combobox(corpo, values=["ADMIN", "ENGENHARIA", "VISITANTE"], state="readonly", font=("Arial", 10))
        cb_perm.set("ENGENHARIA")
        cb_perm.pack(fill="x", ipady=4, pady=(0, 16))

        def salvar():
            nome = e_nome.get().strip()
            cod = e_cod.get().strip()
            perm = cb_perm.get().strip()
            if not nome or not cod:
                messagebox.showwarning("Aviso", "Preencha Nome e Código.", parent=janela)
                return
            res = adicionar_usuario(nome, cod, perm)
            if res:
                messagebox.showinfo("Sucesso", "Usuário cadastrado com sucesso.", parent=janela)
                janela.destroy()
                self._abrir_administracao()
            else:
                messagebox.showerror("Erro", "Não foi possível cadastrar (nome pode já existir).", parent=janela)

        self.botao_flat(corpo, "Salvar", salvar).pack(fill="x")

        janela.update_idletasks()
        janela.geometry(f"+{(janela.winfo_screenwidth()//2)-150}+{(janela.winfo_screenheight()//2)-180}")
        janela.grab_set()

    def _dialog_editar_usuario(self, user_id: int):
        if not self._validar_permissao("admin"):
            return

        u = obter_usuario_por_id(user_id)
        if not u:
            return

        janela = tk.Toplevel(self.root)
        janela.title("Editar Usuário")
        janela.configure(bg=BG)
        janela.attributes("-topmost", True)
        janela.resizable(False, False)

        corpo = tk.Frame(janela, bg=BG, padx=20, pady=20)
        corpo.pack(fill="both", expand=True)

        tk.Label(corpo, text=f"Editar {u['nome']}", bg=BG, fg=TEXTO, font=("Arial", 12, "bold")).pack(anchor="w", pady=(0, 10))

        tk.Label(corpo, text="Nome Completo:", bg=BG, fg=TEXTO_MUTED, font=FONT_CAPTION).pack(anchor="w", pady=(4, 2))
        e_nome = self.campo_entry(corpo)
        e_nome.pack(fill="x", ipady=6, pady=(0, 8))
        e_nome.insert(0, u["nome"])

        tk.Label(corpo, text="Código Engenharia:", bg=BG, fg=TEXTO_MUTED, font=FONT_CAPTION).pack(anchor="w", pady=(4, 2))
        e_cod = self.campo_entry(corpo)
        e_cod.pack(fill="x", ipady=6, pady=(0, 8))
        e_cod.insert(0, u["codigo"])

        tk.Label(corpo, text="Permissão:", bg=BG, fg=TEXTO_MUTED, font=FONT_CAPTION).pack(anchor="w", pady=(4, 2))
        cb_perm = ttk.Combobox(corpo, values=["ADMIN", "ENGENHARIA", "VISITANTE"], state="readonly", font=("Arial", 10))
        cb_perm.set(u["permissao"])
        cb_perm.pack(fill="x", ipady=4, pady=(0, 16))

        def salvar():
            nome = e_nome.get().strip()
            cod = e_cod.get().strip()
            perm = cb_perm.get().strip()
            if not nome or not cod:
                messagebox.showwarning("Aviso", "Preencha Nome e Código.", parent=janela)
                return
            ok = editar_usuario(user_id, nome, cod, perm)
            if ok:
                messagebox.showinfo("Sucesso", "Usuário atualizado com sucesso.", parent=janela)
                janela.destroy()
                self._abrir_administracao()
            else:
                messagebox.showerror("Erro", "Falha ao editar usuário.", parent=janela)

        self.botao_flat(corpo, "Salvar Alterações", salvar).pack(fill="x")

        janela.update_idletasks()
        janela.geometry(f"+{(janela.winfo_screenwidth()//2)-150}+{(janela.winfo_screenheight()//2)-180}")
        janela.grab_set()

    def _toggle_status_usuario(self, user_id: int, desativar: bool):
        if not self._validar_permissao("admin"):
            return

        if user_id == self.usuario_id:
            messagebox.showwarning("Aviso", "Você não pode desativar seu próprio usuário.")
            return

        u = obter_usuario_por_id(user_id)
        if not u:
            return

        acao_str = "desativar" if desativar else "reativar"
        if messagebox.askyesno("Confirmar", f"Deseja realmente {acao_str} o usuário '{u['nome']}'?"):
            if desativar:
                desativar_usuario(user_id)
            else:
                reativar_usuario(user_id)
            self._abrir_administracao()


def _tratar_excecao_global(exc_type, exc_value, exc_traceback):
    """Tratamento global de exceções não capturadas para evitar fechamento abrupoto do executável corporativo."""
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    logger.critical("Exceção não tratada capturada globalmente:", exc_info=(exc_type, exc_value, exc_traceback))
    try:
        messagebox.showerror(
            "Erro Inesperado no Sistema",
            f"Ocorreu um erro não esperado na aplicação:\n\n{exc_value}\n\n"
            "O detalhe do erro foi gravado no log corporativo em /logs."
        )
    except Exception:
        pass


if __name__ == "__main__":
    # Handler global para exceções do Python e do Tkinter
    sys.excepthook = _tratar_excecao_global

    root = tk.Tk()
    root.report_callback_exception = lambda et, ev, tb: _tratar_excecao_global(et, ev, tb)

    # Executar backup de inicialização do sistema
    try:
        fazer_backup("startup_main")
    except Exception as _b_err:
        logger.warning("Não foi possível criar backup inicial: %s", _b_err)

    app = PhoenixTool(root)
    root.mainloop()
