"""Smoke test for the detail panel and its char-level diff rendering."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from PySide6.QtWidgets import QApplication

from src.ui.detail_panel import DetailPanel
from src.ui.styles import detect_theme


def main() -> int:
    app = QApplication.instance() or QApplication(sys.argv)
    p = DetailPanel()
    p.show()

    # Modified strings: should render char-level diff
    p.show_cell(
        "B2", "modified",
        "Queso Manchego",
        "Queso Manchego DOP",
        "antiguo.xlsx", "nuevo.xlsx",
    )
    left = p.left_browser.toHtml()
    right = p.right_browser.toHtml()
    theme = detect_theme()
    del_color = theme.char_del.name().lower()
    ins_color = theme.char_ins.name().lower()
    assert ins_color in right.lower(), "Insertion color not found in right HTML"
    assert "Manchego" in left
    assert "DOP" in right
    print("[OK] Char-diff modified string renders insertion highlight.")

    # Only-left
    p.show_cell("A4", "only_left", "Tomate Rama", None, "antiguo.xlsx", "nuevo.xlsx")
    assert "Tomate Rama" in p.left_browser.toPlainText()
    assert "vacía" in p.right_browser.toPlainText().lower() or "(vac" in p.right_browser.toPlainText().lower()
    print("[OK] Only-left cell shows empty marker on the right.")

    # Only-right
    p.show_cell("A11", "only_right", None, "Aceite Oliva", "antiguo.xlsx", "nuevo.xlsx")
    assert "Aceite Oliva" in p.right_browser.toPlainText()
    print("[OK] Only-right cell shows value on the right.")

    # Number (not string): should NOT do char-diff
    p.show_cell("D2", "modified", 1.20, 1.35, "antiguo.xlsx", "nuevo.xlsx")
    # We expect plain text, no highlight bg color in the html that matches our delete color
    # (very loose check: just ensure both numbers are present)
    assert "1.2" in p.left_browser.toPlainText() or "1,2" in p.left_browser.toPlainText()
    assert "1.35" in p.right_browser.toPlainText() or "1,35" in p.right_browser.toPlainText()
    print("[OK] Numeric modified cell renders both values.")

    # Idle
    p.set_idle()
    assert "selecciona" in p.title.text().lower()
    print("[OK] Idle state restores title.")

    print("\n[OK] DetailPanel smoke test passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
