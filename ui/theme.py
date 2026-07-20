"""Tema visual (cores/fontes) e pequenos helpers de apresentação usados pela
GUI do Phoenix Tool.

Mantido separado de `main.py` para isolar constantes de estilo da lógica de
telas/estado da aplicação.
"""

from __future__ import annotations

BG = "#050505"
BG_CARD = "#111111"
BORDA = "#2d2d2d"
TEXTO = "#f5f5f5"
TEXTO_MUTED = "#8d8d8d"
TEXTO_NUM = "#5f5f5f"
ACCENT = "#ffffff"
ACCENT_SOFT = "#1a1a1a"
HEADER_BG = "#0b0b0b"
FOOTER_BG = "#0b0b0b"

FONT_TITULO = ("Arial", 28, "bold")
FONT_SUBTITULO = ("Arial", 11)
FONT_CAPTION = ("Arial", 9)
FONT_CARD_TITULO = ("Arial", 13, "bold")
FONT_BOTAO = ("Arial", 10, "bold")


def espacar(texto: str) -> str:
    """Simula letter-spacing, deixando legendas pequenas com respiro entre letras."""
    return " ".join(list(texto.upper()))
