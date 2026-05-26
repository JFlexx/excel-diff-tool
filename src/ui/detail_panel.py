from __future__ import annotations

import html
from difflib import SequenceMatcher
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from src.core.diff_engine import format_value, is_formula
from src.ui.styles import detect_theme


class DetailPanel(QWidget):
    """Bottom panel showing the selected cell side-by-side with character-level diff."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._build_ui()
        self.set_idle()

    def _build_ui(self) -> None:
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        bar = QFrame()
        bar.setFrameShape(QFrame.Shape.StyledPanel)
        bl = QHBoxLayout(bar)
        bl.setContentsMargins(8, 4, 8, 4)
        self.title = QLabel("Detalle")
        bold = QFont()
        bold.setBold(True)
        self.title.setFont(bold)
        bl.addWidget(self.title)
        bl.addStretch(1)
        outer.addWidget(bar)

        body = QHBoxLayout()
        body.setContentsMargins(6, 4, 6, 6)
        body.setSpacing(6)

        self.left_header = QLabel("")
        self.left_browser = self._make_browser()
        self.right_header = QLabel("")
        self.right_browser = self._make_browser()

        left_col = QVBoxLayout()
        left_col.setSpacing(2)
        left_col.addWidget(self.left_header)
        left_col.addWidget(self.left_browser, 1)

        right_col = QVBoxLayout()
        right_col.setSpacing(2)
        right_col.addWidget(self.right_header)
        right_col.addWidget(self.right_browser, 1)

        body.addLayout(left_col, 1)
        body.addLayout(right_col, 1)
        outer.addLayout(body, 1)

    def _make_browser(self) -> QTextBrowser:
        b = QTextBrowser()
        b.setOpenLinks(False)
        b.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse | Qt.TextInteractionFlag.TextSelectableByKeyboard)
        mono = QFont("Consolas")
        mono.setStyleHint(QFont.StyleHint.Monospace)
        mono.setPointSize(10)
        b.setFont(mono)
        return b

    # ---------- Public API ----------

    def set_idle(self) -> None:
        self.title.setText("Detalle  ·  selecciona una celda para ver el contenido completo")
        self.left_header.setText("")
        self.right_header.setText("")
        self.left_browser.clear()
        self.right_browser.clear()

    def show_cell(
        self,
        coord: str,
        status: str,
        left_value: Any,
        right_value: Any,
        left_name: str,
        right_name: str,
    ) -> None:
        status_label = {
            "modified": "Modificada",
            "only_left": f"Solo en {left_name}",
            "only_right": f"Solo en {right_name}",
            "equal": "Igual",
        }.get(status, "")
        self.title.setText(f"Detalle: celda <b>{coord}</b>  ·  {status_label}")
        self.left_header.setText(self._header_html(left_name))
        self.right_header.setText(self._header_html(right_name))

        theme = detect_theme()
        l_text = self._fmt(left_value)
        r_text = self._fmt(right_value)

        if status == "modified" and isinstance(left_value, str) and isinstance(right_value, str):
            l_html, r_html = self._char_diff_html(l_text, r_text, theme)
        else:
            l_html = self._plain_html(left_value, l_text)
            r_html = self._plain_html(right_value, r_text)

        self.left_browser.setHtml(self._wrap_html(l_html))
        self.right_browser.setHtml(self._wrap_html(r_html))

    # ---------- Rendering ----------

    @staticmethod
    def _header_html(name: str) -> str:
        return f"<span style='font-weight:600'>{html.escape(name)}</span>"

    @staticmethod
    def _fmt(value: Any) -> str:
        if value is None:
            return ""
        if is_formula(value):
            return value
        return str(format_value(value))

    @staticmethod
    def _plain_html(value: Any, text: str) -> str:
        if value is None:
            return "<i style='opacity:0.6'>(vacía)</i>"
        return html.escape(text).replace("\n", "<br>").replace(" ", "&nbsp;")

    @staticmethod
    def _wrap_html(inner: str) -> str:
        return f"<div style='font-family:Consolas,monospace;font-size:10pt;white-space:pre-wrap'>{inner}</div>"

    @staticmethod
    def _char_diff_html(left: str, right: str, theme) -> tuple[str, str]:
        del_bg = theme.char_del.name()
        ins_bg = theme.char_ins.name()
        sm = SequenceMatcher(None, left, right, autojunk=False)
        l_parts: list[str] = []
        r_parts: list[str] = []
        for op, i1, i2, j1, j2 in sm.get_opcodes():
            l_seg = html.escape(left[i1:i2]).replace(" ", "&nbsp;")
            r_seg = html.escape(right[j1:j2]).replace(" ", "&nbsp;")
            if op == "equal":
                l_parts.append(l_seg)
                r_parts.append(r_seg)
            elif op == "replace":
                l_parts.append(
                    f"<span style='background:{del_bg};color:white;text-decoration:line-through'>{l_seg}</span>"
                )
                r_parts.append(
                    f"<span style='background:{ins_bg};color:white'>{r_seg}</span>"
                )
            elif op == "delete":
                l_parts.append(
                    f"<span style='background:{del_bg};color:white;text-decoration:line-through'>{l_seg}</span>"
                )
            elif op == "insert":
                r_parts.append(
                    f"<span style='background:{ins_bg};color:white'>{r_seg}</span>"
                )
        return "".join(l_parts), "".join(r_parts)
