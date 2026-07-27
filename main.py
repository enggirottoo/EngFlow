import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime
from typing import Any, Dict, List, Optional
import threading
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from core.runner import abrir_planilha, executar_script, logger
from services.storage import (
    atualizar_campos_registro,
    carregar_config,
    carregar_estado_app,
    carregar_historico,
    carregar_login,
    encontrar_por_linha,
    salvar_estado_app,
    salvar_login,
)
from ui import theme
from ui.theme import (
    ACCENT,
    ACCENT_SOFT,
    BG,
    BG_CARD,
    BORDA,
    FONT_BOTAO,
    FONT_CAPTION,
    FONT_CARD_TITULO,
    FONT_SUBTITULO,
    FONT_TITULO,
    FOOTER_BG,
    HEADER_BG,
    TEXTO,
    TEXTO_MUTED,
    espacar,
)

FRAME_LOGIN = "frame_login"
FRAME_MENU = "frame_menu"
FRAME_DASHBOARD = "frame_dashboard"
FRAME_MENU_PHOENIX = "frame_menu_phoenix"
FRAME_MENU_PEGASUS = "frame_menu_pegasus"
FRAME_ATUALIZAR_PLANILHA_PEGASUS = "frame_atualizar_planilha_pegasus"
FRAME_ATUALIZAR_PN = "frame_atualizar_pn"

APP_VERSION = "1.0.0"


def mapear_tela_para_nav(nome: str) -> Optional[str]:
    mapa = {
        FRAME_DASHBOARD: "dashboard",
        FRAME_MENU_PHOENIX: "phoenix",
        FRAME_MENU_PEGASUS: "pegasus",
        FRAME_ATUALIZAR_PN: "phoenix",
        FRAME_ATUALIZAR_PLANILHA_PEGASUS: "pegasus",
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
    elif status:
        status_texto = "Finalizado"
    else:
        status_texto = "Sem status"

    return f"Último registro: {descricao} • Linha {linha} • {status_texto}"


class PhoenixTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Phoenix Tool")
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
        """Copia as cores atuais de ui.theme para as constantes usadas em main.py."""
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

    def _montar_estrutura_chrome(self):
        self._aplicar_estilo()

        self.header = tk.Frame(self.root, bg=HEADER_BG, height=52, highlightbackground=BORDA, highlightthickness=0)
        self.header.pack(fill="x", side="top")
        self.header.pack_propagate(False)

        tk.Label(self.header, text="FLEX • CLASSIFICAÇÃO FISCAL", bg=HEADER_BG, fg=ACCENT, font=("Arial", 10, "bold")).pack(side="left", padx=18)
        texto_badge = f"USUÁRIO: {self.usuario.upper()}" if self.usuario else "USUÁRIO: NÃO LOGADO"
        self.user_badge = tk.Label(self.header, text=texto_badge, bg=HEADER_BG, fg=TEXTO_MUTED, font=("Arial", 8))
        self.user_badge.pack(side="left", padx=(12, 0))
        tk.Label(self.header, text="STATUS: ONLINE", bg=HEADER_BG, fg=TEXTO_MUTED, font=("Arial", 9)).pack(side="right", padx=18)

        self.footer = tk.Frame(self.root, bg=FOOTER_BG, height=30, highlightbackground=BORDA, highlightthickness=0)
        self.footer.pack(fill="x", side="bottom")
        self.footer.pack_propagate(False)

        nav_left = tk.Frame(self.footer, bg=FOOTER_BG)
        nav_left.pack(side="left", padx=18)
        self.nav_labels = {}
        self._criar_item_nav(nav_left, "DASHBOARD", "dashboard")
        tk.Label(nav_left, text="  •  ", bg=FOOTER_BG, fg=TEXTO_MUTED, font=("Arial", 8)).pack(side="left")
        self._criar_item_nav(nav_left, "PHOENIX", "phoenix")
        tk.Label(nav_left, text="  •  ", bg=FOOTER_BG, fg=TEXTO_MUTED, font=("Arial", 8)).pack(side="left")
        self._criar_item_nav(nav_left, "PEGASUS", "pegasus")

        nav_right = tk.Frame(self.footer, bg=FOOTER_BG)
        nav_right.pack(side="right", padx=18)
        self.footer_status = tk.Label(nav_right, text="PRONTO", bg=FOOTER_BG, fg=TEXTO_MUTED, font=("Arial", 8))
        self.footer_status.pack(side="right")
        tk.Label(nav_right, text="  •  ", bg=FOOTER_BG, fg=TEXTO_MUTED, font=("Arial", 8)).pack(side="right")
        tk.Label(nav_right, text="MINIMAL", bg=FOOTER_BG, fg=TEXTO_MUTED, font=("Arial", 8)).pack(side="right")

        self.signature = tk.Label(
            self.root,
            text="Desenvolvido por Gabriell Girotto",
            bg=BG,
            fg="#6f6f6f",
            font=("Arial", 8),
        )
        self.signature.place(relx=0.5, rely=0.985, anchor="s")

        area_rolagem = tk.Frame(self.root, bg=BG)
        area_rolagem.pack(fill="both", expand=True)

        self.main_canvas = tk.Canvas(area_rolagem, bg=BG, highlightthickness=0, bd=0)
        self.vscrollbar = tk.Scrollbar(area_rolagem, orient="vertical", command=self.main_canvas.yview)
        self.hscrollbar = tk.Scrollbar(area_rolagem, orient="horizontal", command=self.main_canvas.xview)
        self.main_canvas.configure(yscrollcommand=self.vscrollbar.set, xscrollcommand=self.hscrollbar.set)

        self.vscrollbar.pack(side="right", fill="y")
        self.hscrollbar.pack(side="bottom", fill="x")
        self.main_canvas.pack(side="left", fill="both", expand=True)

        self.container = tk.Frame(self.main_canvas, bg=BG)
        self.container.grid_columnconfigure(0, weight=1)
        self._container_window = self.main_canvas.create_window((0, 0), window=self.container, anchor="nw")

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

    def _build_title_bar(self, root):
        bar = tk.Frame(root, bg=HEADER_BG, height=32, highlightbackground=BORDA, highlightthickness=0)
        bar.pack(fill="x", side="top")
        bar.pack_propagate(False)

        tk.Label(bar, text="", bg=HEADER_BG).pack(side="left", expand=True)

        btn_close = tk.Label(bar, text="✕", bg=HEADER_BG, fg=TEXTO_MUTED, cursor="hand2", font=("Arial", 10, "bold"))
        btn_close.pack(side="right", padx=(0, 10))
        btn_close.bind("<Button-1>", lambda e: root.destroy())

        btn_min = tk.Label(bar, text="—", bg=HEADER_BG, fg=TEXTO_MUTED, cursor="hand2", font=("Arial", 10, "bold"))
        btn_min.pack(side="right", padx=(0, 8))
        btn_min.bind("<Button-1>", lambda e: root.iconify())

        btn_full = tk.Label(bar, text="▣", bg=HEADER_BG, fg=TEXTO_MUTED, cursor="hand2", font=("Arial", 9, "bold"))
        btn_full.pack(side="right", padx=(0, 8))
        btn_full.bind("<Button-1>", lambda e: self._alternar_tela_cheia())

        tk.Label(bar, text="Phoenix Tool", bg=HEADER_BG, fg=TEXTO_MUTED, font=("Arial", 8)).pack(side="right", padx=(0, 12))

    def _iniciar_arraste(self, event):
        self._drag_offset = (event.x_root - self.root.winfo_x(), event.y_root - self.root.winfo_y())

    def _arrastar_janela(self, event):
        if self._drag_offset is None:
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
        label = tk.Label(
            parent,
            text=texto,
            bg=FOOTER_BG,
            fg=TEXTO_MUTED,
            font=("Arial", 8, "bold"),
            cursor="hand2",
        )
        label.pack(side="left")
        self.nav_labels[chave] = label
        return label

    def _atualizar_nav(self, nome):
        mapa = {
            FRAME_LOGIN: None,
            FRAME_MENU: "dashboard",
            FRAME_DASHBOARD: "dashboard",
            FRAME_MENU_PHOENIX: "phoenix",
            FRAME_ATUALIZAR_PN: "phoenix",
            FRAME_MENU_PEGASUS: "pegasus",
            FRAME_ATUALIZAR_PLANILHA_PEGASUS: "pegasus",
        }
        ativo = mapa.get(nome)
        for chave, label in self.nav_labels.items():
            label.config(fg=ACCENT if chave == ativo else TEXTO_MUTED)
        if hasattr(self, "footer_status"):
            if self.usuario:
                self.footer_status.config(text=f"USUÁRIO: {self.usuario.upper()}  •  TELA: {nome}")
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
        estado = carregar_estado_app()
        tela = str(estado.get("last_screen", "")).strip()
        if tela in self.frames:
            self.mostrar(tela, empilhar=False)
        else:
            self.mostrar(FRAME_LOGIN)
        self.root.after(300, self._try_auto_login)
        self._iniciar_verificacao_atualizacao()

    def _normalizar_versao(self, valor: str) -> List[int]:
        texto = str(valor or "").strip().lower().lstrip("v")
        partes = []
        for item in texto.split("."):
            digitos = "".join(ch for ch in item if ch.isdigit())
            partes.append(int(digitos) if digitos else 0)
        while partes and partes[-1] == 0:
            partes.pop()
        return partes or [0]

    def _versao_mais_nova(self, remota: str, atual: str) -> bool:
        v_remota = self._normalizar_versao(remota)
        v_atual = self._normalizar_versao(atual)
        tamanho = max(len(v_remota), len(v_atual))
        v_remota += [0] * (tamanho - len(v_remota))
        v_atual += [0] * (tamanho - len(v_atual))
        return v_remota > v_atual

    def _iniciar_verificacao_atualizacao(self) -> None:
        cfg = carregar_config()
        url = str(cfg.get("update_check_url", "")).strip()
        if not url:
            return

        def _worker() -> None:
            try:
                req = Request(url, headers={"User-Agent": "PhoenixTool/1.0"})
                with urlopen(req, timeout=8) as resp:
                    bruto = resp.read().decode("utf-8", errors="ignore")
                import json

                payload = json.loads(bruto)
                if not isinstance(payload, dict):
                    return

                versao_remota = str(payload.get("version", "")).strip()
                if not versao_remota:
                    return

                if not self._versao_mais_nova(versao_remota, APP_VERSION):
                    return

                download_url = str(payload.get("download_url", "")).strip()
                notas = str(payload.get("notes", "")).strip()

                def _avisar() -> None:
                    texto = (
                        f"Nova versão disponível: {versao_remota}\n"
                        f"Versão atual: {APP_VERSION}"
                    )
                    if notas:
                        texto += f"\n\nNotas:\n{notas}"
                    if download_url:
                        texto += f"\n\nDownload:\n{download_url}"
                    messagebox.showinfo("Atualização disponível", texto)

                self.root.after(0, _avisar)
            except (HTTPError, URLError, TimeoutError, ValueError):
                return
            except Exception:
                logger.exception("Falha na verificação automática de atualização")

        threading.Thread(target=_worker, daemon=True).start()

   

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
        btn = tk.Button(
            parent, text=espacar(texto), command=comando,
            bg=ACCENT_SOFT, fg=TEXTO, activebackground="#3c2f18", activeforeground=TEXTO,
            relief="solid", bd=1, font=FONT_BOTAO, cursor="hand2",
            width=largura, highlightbackground=BORDA, highlightthickness=1,
            padx=12, pady=10
        )
        btn.bind("<Enter>", lambda e: btn.config(bg="#2f2616"))
        btn.bind("<Leave>", lambda e: btn.config(bg=ACCENT_SOFT))
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
            parent, bg=BG_CARD, fg=TEXTO, insertbackground=TEXTO,
            relief="solid", font=("Arial", 11), show=mostrar,
            highlightbackground=BORDA, highlightthickness=0, bd=0
        )
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

   

    def _build_login(self):
        frame = tk.Frame(self.container, bg=BG)
        self.frames[FRAME_LOGIN] = frame

        painel = tk.Frame(frame, bg=BG_CARD, highlightbackground=BORDA, highlightthickness=0)
        painel.pack(fill="x", padx=28, pady=(38, 16))

        tk.Label(
            painel, text="FLEX • CLASSIFICAÇÃO FISCAL", bg=BG_CARD, fg=ACCENT, font=("Arial", 9, "bold")
        ).pack(pady=(24, 6))

        tk.Label(
            painel, text="LOGIN", bg=BG_CARD, fg=TEXTO, font=FONT_TITULO
        ).pack(pady=(0, 6))

        tk.Label(
            painel, text="Acesso seguro para automações Phoenix, Pegasus e planilha.", bg=BG_CARD, fg=TEXTO_MUTED, font=FONT_SUBTITULO
        ).pack(pady=(0, 18))

        bloco = tk.Frame(painel, bg=BG_CARD)
        bloco.pack(padx=20, pady=8, fill="x")

        tk.Label(
            bloco, text=espacar("Usuário"), bg=BG, fg=TEXTO_MUTED, font=FONT_CAPTION
        ).pack(anchor="w", pady=(20, 6))

        self.campo_usuario = self.campo_entry(bloco)
        self.campo_usuario.pack(fill="x", ipady=8)

        tk.Label(
            bloco, text=espacar("Senha"), bg=BG, fg=TEXTO_MUTED, font=FONT_CAPTION
        ).pack(anchor="w", pady=(24, 6))

        self.campo_senha = self.campo_entry(bloco, mostrar="•")
        self.campo_senha.pack(fill="x", ipady=8)

        dados_login = carregar_login()
        self.lembrar_senha_var = tk.BooleanVar(value=dados_login.get("remember", True))
        chk_lembrar = tk.Checkbutton(
            bloco,
            text=espacar("Lembrar minha senha"),
            variable=self.lembrar_senha_var,
            bg=BG_CARD,
            fg=TEXTO_MUTED,
            selectcolor=BG_CARD,
            activebackground=BG_CARD,
            activeforeground=TEXTO,
            font=FONT_CAPTION,
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )
        chk_lembrar.pack(anchor="w", pady=(14, 0))

        self.campo_usuario.bind("<Return>", lambda e: self.entrar())
        self.campo_senha.bind("<Return>", lambda e: self.entrar())

        self.botao_flat(bloco, "Entrar", self.entrar).pack(fill="x", pady=(24, 0))

        try:
            self.campo_usuario.insert(0, dados_login.get("user", ""))
            self.campo_senha.insert(0, dados_login.get("password", ""))
        except Exception:
            logger.exception("Falha ao pré-preencher os campos de login")

    def _try_auto_login(self):
        dados = carregar_login()
        usuario = (dados.get("user") or "").strip()
        senha = (dados.get("password") or "").strip()
        if usuario and senha and dados.get("remember", True):
            self.campo_usuario.delete(0, tk.END)
            self.campo_senha.delete(0, tk.END)
            self.campo_usuario.insert(0, usuario)
            self.campo_senha.insert(0, senha)
            self.entrar(automatic=True)

    def entrar(self, automatic=False):
        usuario = self.campo_usuario.get().strip()
        senha = self.campo_senha.get().strip()

        if usuario == "":
            if not automatic:
                messagebox.showwarning("Aviso", "Digite o usuário.")
            return
        if senha == "":
            if not automatic:
                messagebox.showwarning("Aviso", "Digite a senha.")
            return

        lembrar = bool(getattr(self, "lembrar_senha_var", None) and self.lembrar_senha_var.get())
        salvar_login(usuario, senha, lembrar=lembrar)
        logger.info("Login efetuado (usuário=%s, lembrar=%s)", usuario, lembrar)
        self.usuario = usuario
        if hasattr(self, "user_badge"):
            self.user_badge.config(text=f"USUÁRIO: {usuario.upper()}")
        self._build_menu()
        self.history = []
        self.mostrar(FRAME_MENU)

  

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
            text="● SISTEMA OPERACIONAL",
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
            painel, text="Acesse as automações e acompanhe o histórico em um único lugar.", bg=BG_CARD, fg=TEXTO_MUTED, font=FONT_SUBTITULO
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

        self.card_navegavel(frame, 1, "Dashboard", self._abrir_dashboard, "Visão geral operacional", chave="dashboard")
        self.card_navegavel(frame, 2, "Phoenix", self._abrir_menu_phoenix, "Fluxo principal", chave="phoenix")
        self.card_navegavel(frame, 3, "Pegasus", self._abrir_menu_pegasus, "Atualizações e planilha", chave="pegasus")
        self.card_navegavel(frame, 4, "Cost Request", lambda: executar_script("automocoes", "cost", "cost.py"), "Abrir automação")
        self.card_navegavel(frame, 5, "Abrir planilha", abrir_planilha, "Acesse a base")

        tk.Label(
            frame,
            text="Atalhos rápidos",
            bg=BG,
            fg=TEXTO_MUTED,
            font=("Arial", 9, "bold"),
        ).pack(anchor="w", padx=28, pady=(10, 6))

        tk.Label(
            frame, text=espacar("Status: conectado"), bg=BG, fg=ACCENT,
            font=FONT_CAPTION
        ).pack(side="bottom", pady=16)

    



    def _carregar_historico(self):
        return carregar_historico()

    def _abrir_dashboard(self):
        self._build_dashboard()
        self.mostrar(FRAME_DASHBOARD)

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

        tk.Label(
            frame,
            text="Últimos registros e métricas operacionais.",
            bg=BG,
            fg=TEXTO_MUTED,
            font=FONT_SUBTITULO,
            justify="left",
        ).pack(anchor="w", padx=28, pady=(0, 10))

        topo_barra = tk.Frame(frame, bg=BG)
        topo_barra.pack(fill="x", padx=28, pady=(0, 10))

        self.botao_flat(topo_barra, "Atualizar", self._abrir_dashboard, largura=12).pack(side="left")
        self.botao_flat(topo_barra, "Planilha", abrir_planilha, largura=12).pack(side="left", padx=(10, 0))

        tk.Label(
            topo_barra,
            text="● atualização automática",
            bg=BG,
            fg="#67d28d",
            font=("Arial", 8, "bold"),
        ).pack(side="left", padx=(12, 0))

        for texto, valor in (("Todos", "todos"), ("Em andamento", "on going"), ("Finalizadas", "finalizadas")):
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
            btn.pack(side="left", padx=(12, 6) if texto == "Todos" else 6)

        historico = self._carregar_historico()
        self._dashboard_last_signature = self._dashboard_signature(historico)
        self._renderizar_dashboard_conteudo(frame, historico)
        self.root.after(2000, self._check_dashboard_refresh)

    def _aplicar_filtro_dashboard(self, valor):
        self._dashboard_filter = valor
        if FRAME_DASHBOARD in self.frames:
            historico = self._carregar_historico()
            self._renderizar_dashboard_conteudo(self.frames[FRAME_DASHBOARD], historico)

    def _renderizar_dashboard_conteudo(self, frame, historico):
        for widget in frame.winfo_children():
            if getattr(widget, "_dashboard_area", False):
                widget.destroy()

        stats = tk.Frame(frame, bg=BG)
        stats._dashboard_area = True
        stats.pack(fill="x", padx=28, pady=(0, 8))

        total = len(historico)
        em_andamento = sum(1 for i in historico if str(i.get("status", "")).upper() == "ON GOING")
        finalizadas = sum(1 for i in historico if str(i.get("status", "")).upper() not in ("ON GOING", ""))

        self._stat_card(stats, total, "Total")
        self._stat_card(stats, em_andamento, "Em andamento")
        self._stat_card(stats, finalizadas, "Finalizadas")

        self.linha_divisoria(frame)

        tk.Label(
            frame,
            text=espacar("Últimas atualizações"),
            bg=BG,
            fg=TEXTO_MUTED,
            font=FONT_CAPTION,
        ).pack(anchor="w", padx=28, pady=(0, 8))

        container = tk.Frame(frame, bg=BG)
        container._dashboard_area = True
        container.pack(fill="both", padx=28, pady=(0, 16))

        if not historico:
            tk.Label(container, text="Nenhum registro encontrado.", bg=BG, fg=TEXTO_MUTED, font=FONT_SUBTITULO).pack(anchor="w", pady=10)
            return

        filtro = (self._dashboard_filter or "todos").strip().lower()
        if filtro == "on going":
            lista = [item for item in historico if str(item.get("status", "")).upper() == "ON GOING"]
        elif filtro == "finalizadas":
            lista = [item for item in historico if str(item.get("status", "")).upper() not in ("ON GOING", "")]
        else:
            lista = list(historico)

        if not lista:
            tk.Label(container, text="Nenhum registro encontrado para este filtro.", bg=BG, fg=TEXTO_MUTED, font=FONT_SUBTITULO).pack(anchor="w", pady=10)
            return

        recentes = list(reversed(lista))[:8]
        for item in recentes:
            self._linha_historico(container, item)

    def _dashboard_signature(self, historico):
        return tuple(
            (
                str(item.get("id", "")),
                str(item.get("linha", "")),
                str(item.get("status", "")),
                str(item.get("description", "")),
                str(item.get("produto", "")),
                str(item.get("solicitante", "")),
                str(item.get("requisitante", "")),
            )
            for item in historico
        )

    def _check_dashboard_refresh(self):
        if self.history and self.history[-1] != FRAME_DASHBOARD:
            return
        if FRAME_DASHBOARD not in self.frames:
            return

        historico = self._carregar_historico()
        signature = self._dashboard_signature(historico)
        if signature != self._dashboard_last_signature:
            self._dashboard_last_signature = signature
            self._renderizar_dashboard_conteudo(self.frames[FRAME_DASHBOARD], historico)

        self.root.after(2000, self._check_dashboard_refresh)

    def _atualizar_dashboard_apos_registro(self):
        if FRAME_DASHBOARD in self.frames:
            self._build_dashboard()
            self.mostrar(FRAME_DASHBOARD)

    def _stat_card(self, parent, valor, legenda):
        card = tk.Frame(parent, bg=BG_CARD, highlightbackground=BORDA, highlightthickness=0)
        card.pack(side="left", expand=True, fill="x", padx=(0, 6))

        tk.Label(
            card, text=str(valor), bg=BG_CARD, fg=TEXTO, font=("Arial", 20, "bold")
        ).pack(pady=(10, 0))

        tk.Label(
            card, text=espacar(legenda), bg=BG_CARD, fg=TEXTO_MUTED, font=FONT_CAPTION
        ).pack(pady=(2, 10))

    def _linha_historico(self, parent, item):
        linha = tk.Frame(parent, bg=BG_CARD, cursor="hand2", highlightbackground=BORDA, highlightthickness=0)
        linha.pack(fill="x", padx=0, pady=6)

        topo = tk.Frame(linha, bg=BG_CARD)
        topo.pack(fill="x")

        status = str(item.get("status", "—")).upper()
        destaque = ACCENT if status == "ON GOING" else TEXTO_MUTED

        tk.Label(
            topo, text=str(item.get("description", "—")).upper(),
            bg=BG_CARD, fg=TEXTO, font=("Arial", 11, "bold")
        ).pack(side="left", padx=12, pady=10)

        tk.Label(
            topo, text=espacar(status),
            bg=BG_CARD, fg=destaque, font=FONT_CAPTION
        ).pack(side="right", padx=12, pady=10)

        if status == "ON GOING":
            tk.Label(
                topo,
                text="PENDENTE",
                bg=BG_CARD,
                fg="#ffbf69",
                font=("Arial", 8, "bold"),
            ).pack(side="right", padx=(0, 8), pady=10)

        btn_editar = tk.Button(
            topo,
            text="Editar",
            command=lambda item=item: self._abrir_editar_registro(item),
            bg=BG_CARD,
            fg=ACCENT,
            activebackground=BORDA,
            activeforeground=TEXTO,
            relief="solid",
            bd=1,
            font=("Arial", 8, "bold"),
            padx=10,
            pady=4,
            cursor="hand2",
        )
        btn_editar.pack(side="right", padx=(0, 8), pady=10)

        tk.Label(
            linha,
            text=f"Linha {item.get('linha', '—')}   ·   {item.get('data_abertura', '—')}"
                 f"   ·   {item.get('user', '—')}",
            bg=BG_CARD, fg=TEXTO_MUTED, font=("Arial", 9)
        ).pack(anchor="w", padx=12, pady=(0, 2))

        solicitante = str(item.get("solicitante") or "").strip() or "Não informado"
        requisitante = str(item.get("requisitante") or "").strip() or "Não informado"
        produto = str(item.get("produto") or "").strip() or "Não informado"
        tk.Label(
            linha,
            text=f"Produto: {produto}",
            bg=BG_CARD, fg=TEXTO_MUTED, font=("Arial", 9)
        ).pack(anchor="w", padx=12, pady=(0, 2))
        tk.Label(
            linha,
            text=f"Solicitante: {solicitante}   ·   Requisitante: {requisitante}",
            bg=BG_CARD, fg=TEXTO_MUTED, font=("Arial", 9)
        ).pack(anchor="w", padx=12, pady=(0, 10))

        tk.Frame(linha, bg=BORDA, height=1).pack(fill="x", padx=12)

        def mostrar_detalhes(e=None):
            messagebox.showinfo(
                "Detalhes",
                f"ID: {item.get('id', '—')}\n\n"
                f"Description: {item.get('description', '—')}\n\n"
                f"Status: {item.get('status', '—')}\n\n"
                f"Data abertura: {item.get('data_abertura', '—')}\n\n"
                f"PN: {item.get('pn') or '—'}\n\n"
                f"Produto: {item.get('produto') or 'Não informado'}\n\n"
                f"Solicitante: {item.get('solicitante') or 'Não informado'}\n\n"
                f"Requisitante: {item.get('requisitante') or 'Não informado'}\n\n"
                f"Usuário: {item.get('user', '—')}"
            )

        for widget in (linha, topo):
            widget.bind("<Button-1>", mostrar_detalhes)

    def _abrir_editar_registro(self, item):
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

        tk.Label(
            corpo, text=espacar("Produto"), bg=BG, fg=TEXTO_MUTED, font=FONT_CAPTION
        ).grid(row=1, column=0, sticky="w", pady=(0, 6))
        campo_produto = self.campo_entry(corpo)
        campo_produto.grid(row=1, column=1, sticky="ew", pady=(0, 6), padx=(10, 0), ipady=4)
        campo_produto.insert(0, str(item.get("produto") or ""))

        tk.Label(
            corpo, text=espacar("Solicitante"), bg=BG, fg=TEXTO_MUTED, font=FONT_CAPTION
        ).grid(row=2, column=0, sticky="w", pady=(0, 6))
        campo_solicitante = self.campo_entry(corpo)
        campo_solicitante.grid(row=2, column=1, sticky="ew", pady=(0, 6), padx=(10, 0), ipady=4)
        campo_solicitante.insert(0, str(item.get("solicitante") or ""))

        tk.Label(
            corpo, text=espacar("Requisitante"), bg=BG, fg=TEXTO_MUTED, font=FONT_CAPTION
        ).grid(row=3, column=0, sticky="w", pady=(0, 6))
        campo_requisitante = self.campo_entry(corpo)
        campo_requisitante.grid(row=3, column=1, sticky="ew", pady=(0, 6), padx=(10, 0), ipady=4)
        campo_requisitante.insert(0, str(item.get("requisitante") or ""))

        corpo.grid_columnconfigure(1, weight=1)

        botoes = tk.Frame(corpo, bg=BG)
        botoes.grid(row=4, column=0, columnspan=2, pady=(16, 0))

        def salvar():
            novo_produto = campo_produto.get().strip()
            novo_solicitante = campo_solicitante.get().strip()
            novo_requisitante = campo_requisitante.get().strip()
            atualizar_campos_registro(
                item.get("linha"),
                {
                    "produto": novo_produto,
                    "solicitante": novo_solicitante,
                    "requisitante": novo_requisitante,
                },
            )
            item["produto"] = novo_produto
            item["solicitante"] = novo_solicitante
            item["requisitante"] = novo_requisitante
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

        campo_solicitante.focus_set()
        janela.grab_set()

   

    def _abrir_menu_phoenix(self):
        self._build_menu_phoenix()
        self.mostrar(FRAME_MENU_PHOENIX)

    def _build_menu_phoenix(self):
        if FRAME_MENU_PHOENIX in self.frames:
            self.frames[FRAME_MENU_PHOENIX].destroy()

        frame = tk.Frame(self.container, bg=BG)
        self.frames[FRAME_MENU_PHOENIX] = frame

        self.build_menu_automacao(
            frame,
            "Phoenix",
            [
                ("Home Phoenix", lambda: executar_script("automocoes", "phoenix", "phoenix.py", arg="home"), "Retorna ao fluxo inicial"),
                ("Nova solicitação", lambda: executar_script("automocoes", "phoenix", "phoenix.py"), "Inicia uma nova solicitação"),
                ("Atualizar PN Phoenix", self._abrir_atualizar_pn_phoenix, "Busca e atualiza o PN"),
            ],
        )



    def _abrir_menu_pegasus(self):
        self._build_menu_pegasus()
        self.mostrar(FRAME_MENU_PEGASUS)

    def _build_menu_pegasus(self):
        if FRAME_MENU_PEGASUS in self.frames:
            self.frames[FRAME_MENU_PEGASUS].destroy()

        frame = tk.Frame(self.container, bg=BG)
        self.frames[FRAME_MENU_PEGASUS] = frame

        self.build_menu_automacao(
            frame,
            "Pegasus",
            [
                ("Home Pegasus", lambda: executar_script("automocoes", "pegasus", "pegasus.py", arg="home"), "Retorna ao fluxo inicial"),
                ("Nova solicitação", lambda: executar_script("automocoes", "pegasus", "pegasus.py"), "Inicia uma nova solicitação"),
                ("Atualizar planilha", self._abrir_atualizar_planilha_pegasus, "Atualiza a planilha com a description"),
            ],
        )

    def _abrir_atualizar_planilha_pegasus(self):
        self._build_atualizar_planilha_pegasus()
        self.mostrar(FRAME_ATUALIZAR_PLANILHA_PEGASUS)

    def _build_atualizar_planilha_pegasus(self):
        if FRAME_ATUALIZAR_PLANILHA_PEGASUS in self.frames:
            self.frames[FRAME_ATUALIZAR_PLANILHA_PEGASUS].destroy()

        frame = tk.Frame(self.container, bg=BG)
        self.frames[FRAME_ATUALIZAR_PLANILHA_PEGASUS] = frame

        self.botao_voltar(frame)
        self.cabecalho(frame, "Pegasus", "Atualizar planilha")

        bloco = tk.Frame(frame, bg=BG)
        bloco.pack(padx=60, fill="x")

        tk.Label(
            bloco, text=espacar("Linha da planilha"), bg=BG, fg=TEXTO_MUTED, font=FONT_CAPTION
        ).pack(anchor="w", pady=(10, 6))

        campo_linha = self.campo_entry(bloco)
        campo_linha.pack(fill="x", ipady=6)
        campo_linha.bind("<Return>", lambda event: buscar())

        item_selecionado = None

        def limpar():
            nonlocal item_selecionado
            campo_linha.delete(0, tk.END)
            item_selecionado = None
            mensagem.config(text="Ainda não há registro selecionado.")
            status_label.config(text="Campo limpo.")

        def buscar():
            nonlocal item_selecionado
            linha = campo_linha.get().strip()
            if not linha:
                status_label.config(text="Digite uma linha antes de buscar.")
                messagebox.showwarning("Pegasus", "Digite a linha da planilha.")
                return

            item = encontrar_por_linha(linha)
            if item is None:
                item_selecionado = None
                mensagem.config(text="Nenhum registro encontrado.")
                status_label.config(text="Linha não encontrada no histórico.")
                messagebox.showwarning("Pegasus", "Linha não encontrada no histórico.")
                return

            item_selecionado = item
            descricao = item.get("description") or "—"
            mensagem.config(text=descricao)
            status_label.config(text=f"Registro encontrado para a linha {linha}.")
            messagebox.showinfo("Pegasus", f"Registro encontrado para a linha {linha}.")

        def atualizar():
            nonlocal item_selecionado
            if item_selecionado is None:
                status_label.config(text="Busque um registro antes de atualizar.")
                messagebox.showwarning("Pegasus", "Busque um registro primeiro.")
                return

            descricao = item_selecionado.get("description") or ""
            if not descricao:
                status_label.config(text="Não há description para atualizar.")
                messagebox.showwarning("Pegasus", "Esta solicitação não possui description para atualizar a planilha.")
                return

            confirmar = messagebox.askyesno(
                "Pegasus",
                f"Deseja atualizar a planilha com a description abaixo?\n\n{descricao}"
            )
            if not confirmar:
                return

            status_label.config(text="Atualização iniciada.")
            executar_script("automocoes", "planilha", "planilha.py", arg=descricao)
            messagebox.showinfo("Pegasus", "Automação iniciada em segundo plano.")

        botoes = tk.Frame(bloco, bg=BG)
        botoes.pack(fill="x", pady=(14, 6))
        self.botao_flat(botoes, "Buscar", buscar, largura=12).pack(side="left")
        self.botao_flat(botoes, "Limpar", limpar, largura=12).pack(side="left", padx=(10, 0))

        tk.Label(
            bloco, text=espacar("Description"), bg=BG, fg=TEXTO_MUTED, font=FONT_CAPTION
        ).pack(anchor="w", pady=(16, 6))

        mensagem = tk.Label(
            bloco,
            text="Ainda não há registro selecionado.",
            bg=BG,
            fg=TEXTO,
            font=("Arial", 10),
            wraplength=300,
            justify="left"
        )
        mensagem.pack(anchor="w", pady=(0, 8))

        status_label = tk.Label(
            bloco,
            text="Pronto para buscar.",
            bg=BG,
            fg=TEXTO_MUTED,
            font=("Arial", 9)
        )
        status_label.pack(anchor="w", pady=(0, 16))

        self.botao_flat(bloco, "Atualizar planilha", atualizar).pack(fill="x", pady=(0, 0))

   

    def _abrir_atualizar_pn_phoenix(self):
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

        def _carregar_registro(linha):
            return encontrar_por_linha(linha)

        def buscar():
            linha = campo_linha.get().strip()
            try:
                item = _carregar_registro(linha)
                if item is None:
                    messagebox.showwarning("Phoenix", "Linha não encontrada.")
                    return
                messagebox.showinfo(
                    "Phoenix",
                    f"Linha: {item.get('id', '—')}\n\n"
                    f"Ticket: {item.get('ticket') or '—'}\n\n"
                    f"Description: {item.get('description', '—')}\n\n"
                    f"Status: {item.get('status', '—')}\n\n"
                    f"Data: {item.get('data_abertura', '—')}\n\n"
                    f"Usuário: {item.get('user', '—')}"
                )
            except Exception as e:
                messagebox.showerror("Erro", str(e))

        def rodar_automacao():
            linha = campo_linha.get().strip()
            if not linha:
                messagebox.showwarning("Phoenix", "Digite a linha da planilha.")
                return

            try:
                item = _carregar_registro(linha)
            except Exception as e:
                messagebox.showerror("Erro", str(e))
                return

            if item is None:
                messagebox.showwarning("Phoenix", "Linha não encontrada no histórico.")
                return

            status_atual = str(item.get("status", "")).upper()
            if status_atual != "ON GOING":
                messagebox.showinfo(
                    "Phoenix",
                    f"Esta linha já está com status '{item.get('status')}'. "
                    "Nada a fazer."
                )
                return

            confirmar = messagebox.askyesno(
                "Phoenix",
                "Isso vai abrir o navegador e buscar automaticamente o PN "
                f"desta solicitação (linha {linha}).\n\n"
                "Acompanhe o console para ver o progresso. Deseja continuar?"
            )
            if not confirmar:
                return

            executar_script("automocoes", "phoenix", "atualizar_pn.py", arg=linha)
            messagebox.showinfo(
                "Phoenix",
                "Automação iniciada em segundo plano.\n"
                "Acompanhe o console do script para ver o andamento."
            )

        self.botao_flat(bloco, "Buscar PN", buscar).pack(fill="x", pady=(14, 6))
        self.botao_flat(bloco, "Rodar automação (PN)", rodar_automacao).pack(fill="x", pady=(0, 0))



if __name__ == "__main__":
    root = tk.Tk()
    app = PhoenixTool(root)
    root.mainloop()
