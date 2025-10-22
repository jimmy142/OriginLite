# originlite/data/table_model.py
from __future__ import annotations

from typing import List, Optional
import numpy as np
from PySide6.QtCore import QAbstractTableModel, Qt, QModelIndex


class NumpyTableModel(QAbstractTableModel):
    """
    2D numpy array model with editable numeric cells.
    - Empty string ↔ NaN
    - Displays NaN as blank
    - Emits dataChanged on edits
    """

    def __init__(self, data: np.ndarray, headers: Optional[List[str]] = None):
        super().__init__()
        self._data = data  # shape (rows, cols), float with NaNs for empty
        self._headers = headers or [f"C{j+1}" for j in range(self._data.shape[1])]

    # ----- public update API -----
    def update_all(self, data: np.ndarray, headers: Optional[List[str]] = None):
        self.beginResetModel()
        self._data = data
        if headers is not None:
            self._headers = list(headers)
        else:
            # keep length in sync
            if len(self._headers) != self._data.shape[1]:
                self._headers = [f"C{j+1}" for j in range(self._data.shape[1])]
        self.endResetModel()

    # ----- required overrides -----
    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else (self._data.shape[0] if self._data is not None else 0)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else (self._data.shape[1] if self._data is not None else 0)

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
        if not index.isValid():
            return None
        r, c = index.row(), index.column()
        v = self._data[r, c]

        if role in (Qt.DisplayRole, Qt.EditRole):
            if np.isnan(v):
                return ""  # show empty for NaN
            # Keep it readable but precise
            return f"{v:g}"
        if role == Qt.TextAlignmentRole:
            return int(Qt.AlignVCenter | Qt.AlignRight)
        return None

    def setData(self, index: QModelIndex, value, role: int = Qt.EditRole) -> bool:
        if role != Qt.EditRole or not index.isValid():
            return False
        r, c = index.row(), index.column()
        text = str(value).strip()

        if text == "":
            new_val = np.nan
        else:
            try:
                new_val = float(text)
            except Exception:
                # reject invalid text
                return False

        if np.isnan(self._data[r, c]) and np.isnan(new_val):
            # no real change
            pass
        else:
            self._data[r, c] = new_val

        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
        return True

    def flags(self, index: QModelIndex):
        if not index.isValid():
            return Qt.NoItemFlags
        return Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            if 0 <= section < len(self._headers):
                return self._headers[section]
            return f"C{section+1}"
        else:
            return str(section + 1)

    # Keep headers in sync when caller updates them
    def set_headers(self, headers: list[str]):
        self._headers = list(headers)
        # Trigger a header repaint
        self.headerDataChanged.emit(Qt.Horizontal, 0, max(0, len(headers) - 1))
