from __future__ import annotations

from dataclasses import dataclass

from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication


@dataclass(frozen=True)
class Theme:
    name: str
    # diff backgrounds
    modified: QColor       # both sides have content, differs
    present: QColor        # cell with a value where the other side has none
    absent: QColor         # the gap cell on the other side (visual "missing" indicator)
    equal: QColor          # matches neighborhood base
    # text accents for char-level diff
    char_del: QColor
    char_ins: QColor
    # header / accent
    header_bg: QColor
    header_fg: QColor


LIGHT = Theme(
    name="light",
    modified=QColor("#FFF4CE"),
    present=QColor("#CCE5FF"),
    absent=QColor("#ECECEC"),
    equal=QColor("#FFFFFF"),
    char_del=QColor("#C0142C"),
    char_ins=QColor("#1A7A30"),
    header_bg=QColor("#2D5C8C"),
    header_fg=QColor("#FFFFFF"),
)


DARK = Theme(
    name="dark",
    modified=QColor("#5C4A1A"),
    present=QColor("#1F3A5F"),
    absent=QColor("#3A3A3A"),
    equal=QColor("#2B2B2B"),
    char_del=QColor("#FF7A88"),
    char_ins=QColor("#82D896"),
    header_bg=QColor("#3A6AA8"),
    header_fg=QColor("#FFFFFF"),
)


_active: Theme | None = None


def detect_theme() -> Theme:
    """Pick light or dark theme based on the active QPalette window color."""
    global _active
    if _active is not None:
        return _active
    app = QApplication.instance()
    if app is None:
        return LIGHT
    win = app.palette().color(QPalette.ColorRole.Window)
    _active = DARK if win.lightness() < 128 else LIGHT
    return _active


def cell_color(status_value: str, side: str) -> QColor:
    """Side-aware background color for a diff cell.

    side: 'left' or 'right' — when a cell exists only on one side, the side
    with content shows `present`, the other shows `absent`.
    """
    t = detect_theme()
    if status_value == "equal":
        return t.equal
    if status_value == "modified":
        return t.modified
    if status_value == "only_left":
        return t.present if side == "left" else t.absent
    if status_value == "only_right":
        return t.absent if side == "left" else t.present
    return t.equal


STATUS_LABEL = {
    "modified": "Modificada",
    "only_left": "Solo en A",
    "only_right": "Solo en B",
    "equal": "Igual",
}
