from __future__ import annotations

import sys
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication, QMessageBox

from src.ui.main_window import MainWindow


def _install_excepthook() -> None:
    """Surface uncaught Python exceptions instead of letting Qt swallow them silently."""
    def hook(exc_type, exc, tb):
        msg = "".join(traceback.format_exception(exc_type, exc, tb))
        print(msg, file=sys.stderr, flush=True)
        try:
            QMessageBox.critical(None, "Error inesperado", msg)
        except Exception:
            pass
    sys.excepthook = hook


def _apply_system_theme(app: QApplication) -> None:
    """Follow the OS color scheme. Qt's default Windows style ignores dark mode,
    so we force Fusion + a matching palette when the system reports dark."""
    app.setStyle("Fusion")
    try:
        scheme = app.styleHints().colorScheme()
        is_dark = scheme == Qt.ColorScheme.Dark
    except Exception:
        is_dark = False
    if not is_dark:
        return
    p = QPalette()
    p.setColor(QPalette.ColorRole.Window, QColor("#2B2B2B"))
    p.setColor(QPalette.ColorRole.WindowText, QColor("#E6E6E6"))
    p.setColor(QPalette.ColorRole.Base, QColor("#1E1E1E"))
    p.setColor(QPalette.ColorRole.AlternateBase, QColor("#262626"))
    p.setColor(QPalette.ColorRole.Text, QColor("#E6E6E6"))
    p.setColor(QPalette.ColorRole.Button, QColor("#3A3A3A"))
    p.setColor(QPalette.ColorRole.ButtonText, QColor("#E6E6E6"))
    p.setColor(QPalette.ColorRole.Highlight, QColor("#3A6AA8"))
    p.setColor(QPalette.ColorRole.HighlightedText, QColor("#FFFFFF"))
    p.setColor(QPalette.ColorRole.ToolTipBase, QColor("#3A3A3A"))
    p.setColor(QPalette.ColorRole.ToolTipText, QColor("#E6E6E6"))
    p.setColor(QPalette.ColorRole.PlaceholderText, QColor("#888888"))
    p.setColor(QPalette.ColorRole.BrightText, QColor("#FF7A88"))
    p.setColor(QPalette.ColorRole.Link, QColor("#7CA9E3"))
    p.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Text, QColor("#777777"))
    p.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.ButtonText, QColor("#777777"))
    app.setPalette(p)


def main() -> int:
    _install_excepthook()
    app = QApplication(sys.argv)
    app.setApplicationName("Excel Diff Tool")
    _apply_system_theme(app)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
