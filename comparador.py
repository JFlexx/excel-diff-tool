from __future__ import annotations

import sys
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

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


def main() -> int:
    _install_excepthook()
    app = QApplication(sys.argv)
    app.setApplicationName("Comparador Excel CAF")
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
