from __future__ import annotations

from typing import List, Tuple, Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QComboBox, QListWidget, QListWidgetItem,
    QDialogButtonBox, QLineEdit, QAbstractItemView
)


class AddTraceDialog(QDialog):
    def __init__(self, parent, worksheets: List[Tuple[str, object]]):
        super().__init__(parent)
        self.setWindowTitle("Add Trace")
        self._worksheets = worksheets
        lay = QVBoxLayout(self)
        form = QFormLayout()
        self.ws_combo = QComboBox()
        for title, ws in worksheets:
            self.ws_combo.addItem(title, userData=ws)
        self.x_combo = QComboBox()
        self.y_list = QListWidget(); self.y_list.setSelectionMode(QAbstractItemView.MultiSelection)
        self.label_edit = QLineEdit(); self.label_edit.setPlaceholderText("Label (optional)")
        form.addRow("Worksheet", self.ws_combo)
        form.addRow("X column", self.x_combo)
        form.addRow("Y column(s)", self.y_list)
        form.addRow("Label", self.label_edit)
        lay.addLayout(form)
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept); btns.rejected.connect(self.reject)
        lay.addWidget(btns)
        self.ws_combo.currentIndexChanged.connect(self._populate_columns)
        self._populate_columns(0)

    def _populate_columns(self, idx: int):
        ws = self.ws_combo.currentData()
        self.x_combo.clear(); self.y_list.clear()
        if not ws:
            return
        headers = ws.headers
        self.x_combo.addItems(headers)
        for j, h in enumerate(headers):
            item = QListWidgetItem(h)
            item.setData(Qt.UserRole, j)
            self.y_list.addItem(item)

    def values(self) -> Optional[tuple]:
        ws = self.ws_combo.currentData()
        if ws is None:
            return None
        xi = self.x_combo.currentIndex()
        yjs = [self.y_list.item(i).data(Qt.UserRole) for i in range(self.y_list.count()) if self.y_list.item(i).isSelected()]
        label = self.label_edit.text().strip()
        return ws, xi, yjs, label
