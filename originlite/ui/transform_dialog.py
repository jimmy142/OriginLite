# originlite/ui/transform_dialog.py
from __future__ import annotations
from typing import List
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QLineEdit, QLabel, QDialogButtonBox
)


def _excel_col_name(n: int) -> str:
    s = ""
    n = int(n)
    while True:
        n, r = divmod(n, 26)
        s = chr(ord('A') + r) + s
        if n == 0:
            return s
        n -= 1


class TransformDialog(QDialog):
    def __init__(self, parent=None, columns: List[str] | None = None):
        super().__init__(parent)
        self.setWindowTitle("Add Column from Expression")

        self.expr_edit = QLineEdit()
        self.expr_edit.setPlaceholderText("e.g. 2*A - 0.5*B, log10(C), where(A>0, A, 0)")

        self.name_edit = QLineEdit()
        # Pre-fill with the next Excel-like letter if we know how many columns exist
        if columns:
            # infer next index by current count
            next_name = _excel_col_name(len(columns))
            self.name_edit.setText(next_name)
        else:
            self.name_edit.setPlaceholderText("New column name, e.g. D")

        help_lbl = QLabel(
            "Use letters A, B, C… for columns.\n"
            "Functions: sin, cos, tan, exp, log, log10, sqrt, abs, min, max, where, np.*"
        )
        help_lbl.setWordWrap(True)

        form = QFormLayout()
        form.addRow("Expression", self.expr_edit)
        form.addRow("New column name", self.name_edit)

        lay = QVBoxLayout(self)
        lay.addLayout(form)
        lay.addWidget(help_lbl)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)

    def get_values(self):
        return self.expr_edit.text().strip(), self.name_edit.text().strip()
