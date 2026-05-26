from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, date, time
from enum import Enum
from typing import Any, Iterable

import openpyxl
from openpyxl.utils import get_column_letter


class DiffStatus(str, Enum):
    EQUAL = "equal"
    MODIFIED = "modified"
    ONLY_LEFT = "only_left"
    ONLY_RIGHT = "only_right"


class SheetPresence(str, Enum):
    BOTH = "both"
    ONLY_LEFT = "only_left"
    ONLY_RIGHT = "only_right"


def is_formula(v: Any) -> bool:
    return isinstance(v, str) and v.startswith("=")


def values_equal(a: Any, b: Any) -> bool:
    a_f = is_formula(a)
    b_f = is_formula(b)
    if a_f != b_f:
        return False
    if a_f:
        return a == b
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool) and not isinstance(b, bool):
        try:
            return float(a) == float(b)
        except (TypeError, ValueError):
            return False
    return a == b


def is_empty(v: Any) -> bool:
    return v is None or (isinstance(v, str) and v == "")


def cell_status(left: Any, right: Any) -> DiffStatus:
    l_empty = is_empty(left)
    r_empty = is_empty(right)
    if l_empty and r_empty:
        return DiffStatus.EQUAL
    if l_empty:
        return DiffStatus.ONLY_RIGHT
    if r_empty:
        return DiffStatus.ONLY_LEFT
    if values_equal(left, right):
        return DiffStatus.EQUAL
    return DiffStatus.MODIFIED


def format_value(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "VERDADERO" if v else "FALSO"
    if isinstance(v, datetime):
        if v.hour or v.minute or v.second:
            return v.strftime("%Y-%m-%d %H:%M:%S")
        return v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, time):
        return v.strftime("%H:%M:%S")
    if isinstance(v, float):
        if v.is_integer():
            return str(int(v))
        return f"{v:g}"
    return str(v)


@dataclass
class SheetDiff:
    name: str
    presence: SheetPresence
    diffs: dict[tuple[int, int], DiffStatus] = field(default_factory=dict)
    max_row: int = 1
    max_col: int = 1

    @property
    def diff_count(self) -> int:
        return len(self.diffs)


@dataclass
class WorkbookDiff:
    left_path: str
    right_path: str
    sheets: dict[str, SheetDiff] = field(default_factory=dict)

    @property
    def total_diffs(self) -> int:
        return sum(s.diff_count for s in self.sheets.values())


def _sheet_bounds(ws) -> tuple[int, int]:
    """Real content bounds: highest row/col with a non-empty value.

    Some Excel files inflate max_row/max_column with phantom formatting,
    which would otherwise force us to render millions of empty cells.
    """
    reported_r = max(ws.max_row or 0, 0)
    reported_c = max(ws.max_column or 0, 0)
    if reported_r == 0 or reported_c == 0:
        return 1, 1
    max_r = 0
    max_c = 0
    for row_idx, row in enumerate(
        ws.iter_rows(min_row=1, max_row=reported_r, min_col=1, max_col=reported_c, values_only=True),
        start=1,
    ):
        for col_idx, val in enumerate(row, start=1):
            if val is not None and val != "":
                if row_idx > max_r:
                    max_r = row_idx
                if col_idx > max_c:
                    max_c = col_idx
    return max(max_r, 1), max(max_c, 1)


def _diff_sheet_both(name: str, ws_l, ws_r) -> SheetDiff:
    mr_l, mc_l = _sheet_bounds(ws_l)
    mr_r, mc_r = _sheet_bounds(ws_r)
    max_row = max(mr_l, mr_r)
    max_col = max(mc_l, mc_r)
    diffs: dict[tuple[int, int], DiffStatus] = {}
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            lv = ws_l.cell(r, c).value if r <= mr_l and c <= mc_l else None
            rv = ws_r.cell(r, c).value if r <= mr_r and c <= mc_r else None
            st = cell_status(lv, rv)
            if st != DiffStatus.EQUAL:
                diffs[(r, c)] = st
    return SheetDiff(name=name, presence=SheetPresence.BOTH, diffs=diffs, max_row=max_row, max_col=max_col)


def _diff_sheet_one_side(name: str, ws, side: str) -> SheetDiff:
    mr, mc = _sheet_bounds(ws)
    diffs: dict[tuple[int, int], DiffStatus] = {}
    status = DiffStatus.ONLY_LEFT if side == "left" else DiffStatus.ONLY_RIGHT
    for r in range(1, mr + 1):
        for c in range(1, mc + 1):
            v = ws.cell(r, c).value
            if not is_empty(v):
                diffs[(r, c)] = status
    presence = SheetPresence.ONLY_LEFT if side == "left" else SheetPresence.ONLY_RIGHT
    return SheetDiff(name=name, presence=presence, diffs=diffs, max_row=mr, max_col=mc)


def compute_workbook_diff(wb_left, wb_right, left_path: str, right_path: str) -> WorkbookDiff:
    left_names = list(wb_left.sheetnames)
    right_names = list(wb_right.sheetnames)
    seen: set[str] = set()
    ordered: list[str] = []
    for n in left_names + right_names:
        if n not in seen:
            ordered.append(n)
            seen.add(n)

    sheets: dict[str, SheetDiff] = {}
    for name in ordered:
        in_l = name in wb_left.sheetnames
        in_r = name in wb_right.sheetnames
        if in_l and in_r:
            sheets[name] = _diff_sheet_both(name, wb_left[name], wb_right[name])
        elif in_l:
            sheets[name] = _diff_sheet_one_side(name, wb_left[name], "left")
        else:
            sheets[name] = _diff_sheet_one_side(name, wb_right[name], "right")
    return WorkbookDiff(left_path=left_path, right_path=right_path, sheets=sheets)


def recompute_cell_status(ws_left, ws_right, row: int, col: int) -> DiffStatus:
    """Re-evaluate the diff status of a single cell after an edit."""
    lv = ws_left.cell(row, col).value if ws_left is not None else None
    rv = ws_right.cell(row, col).value if ws_right is not None else None
    return cell_status(lv, rv)


def cell_coord(row: int, col: int) -> str:
    return f"{get_column_letter(col)}{row}"


def sorted_diff_keys(sd: SheetDiff) -> list[tuple[int, int]]:
    return sorted(sd.diffs.keys())
