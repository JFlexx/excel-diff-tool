"""Smoke test for the live side-by-side comparator: load, copy across, verify diff updates."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import openpyxl
from PySide6.QtWidgets import QApplication

from src.core.diff_engine import DiffStatus, compute_workbook_diff, recompute_cell_status
from src.ui.diff_view import DiffView


def main() -> int:
    app = QApplication.instance() or QApplication(sys.argv)

    left_path = str(ROOT / "ejemplos" / "antiguo.xlsx")
    right_path = str(ROOT / "ejemplos" / "nuevo.xlsx")
    wb_l = openpyxl.load_workbook(left_path, data_only=False)
    wb_r = openpyxl.load_workbook(right_path, data_only=False)

    diff = compute_workbook_diff(wb_l, wb_r, left_path, right_path)
    sd = diff.sheets["Productos"]
    print(f"Productos diffs before: {sd.diff_count}")
    assert sd.diff_count == 31

    ws_l = wb_l["Productos"]
    ws_r = wb_r["Productos"]

    # D2 (Manzana price) differs: 1.2 in old, 1.35 in new
    assert ws_l["D2"].value == 1.2
    assert ws_r["D2"].value == 1.35
    assert (2, 4) in sd.diffs

    view = DiffView(ws_l, ws_r, sd)

    # Simulate selecting D2 and copying right -> left
    view.left_table.setCurrentCell(1, 3)  # 0-indexed → D2
    view.copy_right_to_left()
    assert ws_l["D2"].value == 1.35, f"left D2 should be 1.35, got {ws_l['D2'].value}"
    assert (2, 4) not in sd.diffs, "D2 should no longer be a diff"
    print(f"After copying D2 right->left, diffs: {sd.diff_count}")
    assert sd.diff_count == 30

    # E2 (Manzana stock): old=150, new=180
    view.left_table.setCurrentCell(1, 4)  # E2
    view.copy_left_to_right()
    assert ws_r["E2"].value == 150, f"right E2 should be 150, got {ws_r['E2'].value}"
    assert (2, 5) not in sd.diffs
    print(f"After copying E2 left->right, diffs: {sd.diff_count}")
    assert sd.diff_count == 29

    # Simulate inline edit on left B2 (Producto name was equal: "Manzana Golden")
    item = view.left_table.item(1, 1)
    item.setText("Manzana Royal Gala")
    view._on_left_edited(item)
    assert ws_l["B2"].value == "Manzana Royal Gala"
    assert sd.diffs.get((2, 2)) == DiffStatus.MODIFIED
    print(f"After inline editing B2 on left, B2 status: {sd.diffs.get((2,2))}")

    # Now align right B2 to match left → should clear that diff
    view.left_table.setCurrentCell(1, 1)  # B2
    view.copy_left_to_right()
    assert ws_r["B2"].value == "Manzana Royal Gala"
    assert (2, 2) not in sd.diffs
    print(f"Final diffs: {sd.diff_count}")

    print("[OK] All live-edit smoke tests passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
