"""Simulate the full UI flow: set paths, run compare, materialize every sheet."""
from __future__ import annotations

import sys
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from PySide6.QtCore import QEventLoop, QTimer
from PySide6.QtWidgets import QApplication

from src.ui.main_window import MainWindow, TabPanel


def main() -> int:
    app = QApplication.instance() or QApplication(sys.argv)
    w = MainWindow()
    w.show()

    w.left_path.setText(str(ROOT / "ejemplos" / "antiguo.xlsx"))
    w.right_path.setText(str(ROOT / "ejemplos" / "nuevo.xlsx"))
    w._start_compare()

    # Wait for the worker to finish
    loop = QEventLoop()
    def done():
        if w.diff is not None:
            loop.quit()
        else:
            QTimer.singleShot(50, done)
    QTimer.singleShot(50, done)
    loop.exec()

    print(f"Diff ready, sheets: {list(w.diff.sheets.keys())}")

    # Walk every tab, materializing it
    for i in range(w.tabs.count()):
        w.tabs.setCurrentIndex(i)
        QApplication.processEvents()
        panel = w.tabs.widget(i)
        assert isinstance(panel, TabPanel)
        assert panel.view is not None, f"Tab {i} not materialized"
        sheet_name = panel.sheet_diff.name
        diffs = panel.sheet_diff.diff_count
        print(f"  Tab {i} '{sheet_name}': materialized OK, diffs={diffs}")

    print("[OK] All tabs materialized without errors.")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        traceback.print_exc()
        sys.exit(1)
