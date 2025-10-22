# originlite/ui/data_dock.py
from __future__ import annotations

from typing import Callable, Optional
import numpy as np

from PySide6.QtWidgets import (
    QDockWidget, QWidget, QVBoxLayout, QToolBar, QTableView, QMessageBox
)
from PySide6.QtGui import QAction

from ..data.table_model import NumpyTableModel


class DataDock(QDockWidget):
    """
    Spreadsheet-like view of the current dataset.
    Exposes callbacks so the main window can handle actions (add expr, delete col).
    """
    def __init__(self, parent=None):
        super().__init__("Data Table", parent)
        self.setObjectName("DataDock")

        self._table_view = QTableView()
        self._model: Optional[NumpyTableModel] = None

        container = QWidget()
        lay = QVBoxLayout(container)
        lay.setContentsMargins(0, 0, 0, 0)

        self.tb = QToolBar("Data")
        lay.addWidget(self.tb)
        lay.addWidget(self._table_view)
        self.setWidget(container)

        # actions
        self.action_add = QAction("Add Column (Expr)…", self)
        self.action_delete = QAction("Delete Column", self)
        self.tb.addAction(self.action_add)
        self.tb.addAction(self.action_delete)

        # callbacks (set by MainWindow)
        self._on_add_expr: Optional[Callable[[], None]] = None
        self._on_delete_col: Optional[Callable[[int], None]] = None

        # wire
        self.action_add.triggered.connect(self._trigger_add_expr)
        self.action_delete.triggered.connect(self._delete_selected_column)

    # ---- public api ----
    def set_callbacks(
        self,
        on_add_expr: Optional[Callable[[], None]] = None,
        on_delete_column: Optional[Callable[[int], None]] = None,
    ) -> None:
        self._on_add_expr = on_add_expr
        self._on_delete_col = on_delete_column

    def bind(self, data: np.ndarray, headers: list[str]) -> None:
        if self._model is None:
            self._model = NumpyTableModel(data, headers)
            self._table_view.setModel(self._model)
        else:
            self._model.update_all(data, headers)

    def current_column(self) -> Optional[int]:
        idx = self._table_view.currentIndex()
        return idx.column() if idx.isValid() else None

    # ---- internals ----
    def _trigger_add_expr(self) -> None:
        if self._on_add_expr is not None:
            self._on_add_expr()

    def _delete_selected_column(self) -> None:
        if self._on_delete_col is None:
            return
        col = self.current_column()
        if col is None:
            QMessageBox.information(
                self, "Delete Column", "Select a column to delete (click a cell)."
            )
            return
        self._on_delete_col(col)
