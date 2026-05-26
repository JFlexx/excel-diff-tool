from PySide6.QtGui import QColor

COLOR_MODIFIED = QColor("#FFF4CE")      # amber: cells with different values
COLOR_ONLY_LEFT = QColor("#FADBD8")     # soft red: missing on the new side
COLOR_ONLY_RIGHT = QColor("#D5F5E3")    # soft green: missing on the old side
COLOR_EQUAL = QColor("#FFFFFF")         # white: equal
COLOR_SELECTED_DIFF = QColor("#B8DAFF") # blue: highlight for the cell being focused

STATUS_LABEL = {
    "modified": "Modificada",
    "only_left": "Solo en antiguo",
    "only_right": "Solo en nuevo",
    "equal": "Igual",
}

STATUS_COLOR = {
    "modified": COLOR_MODIFIED,
    "only_left": COLOR_ONLY_LEFT,
    "only_right": COLOR_ONLY_RIGHT,
    "equal": COLOR_EQUAL,
}
