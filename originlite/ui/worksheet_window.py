# originlite/ui/worksheet_window.py
from __future__ import annotations

from typing import Optional, List, Dict
import os
import csv
import re
import numpy as np

from PySide6.QtCore import Qt, QModelIndex, Signal, QPoint
from PySide6.QtGui import QFont, QPalette, QAction, QKeySequence, QGuiApplication
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QTableView, QAbstractItemView, QFileDialog,
    QMessageBox, QHeaderView, QStyledItemDelegate, QStyleOptionViewItem, QLineEdit, QMenu, QMdiSubWindow
)

from ..data.table_model import NumpyTableModel
from PySide6.QtCore import QAbstractTableModel


def _find_subwindow(widget) -> QMdiSubWindow | None:
    w = widget.parentWidget()
    while w is not None and not isinstance(w, QMdiSubWindow):
        w = w.parentWidget()
    return w


class NamesRowModel(QAbstractTableModel):
    dataChanged = Signal(QModelIndex, QModelIndex)

    def __init__(self, names: List[str]):
        super().__init__()
        self._names = names

    def update_names(self, names: List[str]):
        self.beginResetModel()
        self._names = names
        self.endResetModel()

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 1

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return len(self._names)

    def data(self, index: QModelIndex, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        j = index.column()
        if role in (Qt.DisplayRole, Qt.EditRole):
            return self._names[j]
        if role == Qt.TextAlignmentRole:
            return int(Qt.AlignVCenter | Qt.AlignLeft)
        return None

    def setData(self, index: QModelIndex, value, role=Qt.EditRole):
        if not index.isValid() or role != Qt.EditRole:
            return False
        j = index.column()
        self._names[j] = (value or "").strip()
        self.dataChanged.emit(index, index)
        return True

    def flags(self, index: QModelIndex):
        if not index.isValid():
            return Qt.NoItemFlags
        return Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable


class PlaceholderDelegate(QStyledItemDelegate):
    def __init__(self, placeholder="Name"):
        super().__init__()
        self.placeholder = placeholder

    def paint(self, painter, option: QStyleOptionViewItem, index: QModelIndex):
        value = index.data(Qt.DisplayRole)
        if not value:
            opt = QStyleOptionViewItem(option)
            self.initStyleOption(opt, index)
            painter.save()
            painter.setPen(opt.palette.color(QPalette.Disabled, QPalette.Text))
            rect = opt.rect.adjusted(6, 0, -6, 0)
            painter.drawText(rect, int(Qt.AlignVCenter | Qt.AlignLeft), self.placeholder)
            painter.restore()
        else:
            super().paint(painter, option, index)

    def createEditor(self, parent, option, index):
        return QLineEdit(parent)


class WorksheetWindow(QMainWindow):
    """
    Origin-like worksheet window with no menus/toolbars; pure grid.
    - Top "Names" strip (editable cells) for Long Names per column.
    - Equal column widths (based on column A) at init/resize.
    - Clipboard: Copy / Cut / Paste (Excel-compatible, TSV/CSV)
    - Close prompts to Save(minimize)/Delete/Cancel.
    - Emits sheet_changed on any data/model structural change for live plots.
    """

    # 🔔 external consumers (plots) listen to this for live updates
    sheet_changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Book*")
        self.resize(820, 520)

        self._model: Optional[NumpyTableModel] = None
        self.data: Optional[np.ndarray] = None
        self.headers: List[str] = []
        self._base_headers: List[str] = []
        self.long_names: List[str] = []

        # Roles
        self.x_col: Optional[int] = None
        self.y_cols: set[int] = set()
        self.z_col: Optional[int] = None

        self._last_sel_model = None
        self._sync_guard = False

        self.names_view = QTableView()
        self.table = QTableView()

        self.names_model = NamesRowModel(self.long_names)
        self.names_view.setModel(self.names_model)
        self.names_view.setItemDelegate(PlaceholderDelegate("Name"))
        self.names_view.verticalHeader().setVisible(False)
        self.names_view.horizontalHeader().setVisible(False)
        self.names_view.setFixedHeight(32)
        self.names_view.setSelectionMode(QAbstractItemView.NoSelection)
        self.names_view.setFocusPolicy(Qt.ClickFocus)
        self.names_view.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
        self.names_view.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.names_view.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)

        self.table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.verticalHeader().setVisible(True)
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.horizontalHeader().setSectionsMovable(False)
        self.table.setSortingEnabled(False)
        self.table.setAlternatingRowColors(False)
        self.table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._on_context_menu)
        self._make_clipboard_actions()

        self.names_view.horizontalScrollBar().valueChanged.connect(
            self.table.horizontalScrollBar().setValue
        )
        self.table.horizontalScrollBar().valueChanged.connect(
            self.names_view.horizontalScrollBar().setValue
        )

        self.table.horizontalHeader().sectionResized.connect(self._sync_from_table_resize)
        self.names_view.horizontalHeader().sectionResized.connect(self._sync_from_names_resize)

        # Names row edits also count as "sheet changes" (affect labels/legends)
        self.names_model.dataChanged.connect(lambda *_: self.sheet_changed.emit())

        container = QWidget()
        lay = QVBoxLayout(container)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        lay.addWidget(self.names_view)
        lay.addWidget(self.table)
        self.setCentralWidget(container)

        self.create_empty(rows=1000, cols=2)

    # ----- Close prompt -----
    def closeEvent(self, event):
        box = QMessageBox(self)
        box.setWindowTitle("Close Worksheet")
        box.setText("Do you want to save this worksheet or delete it?")
        save_btn = box.addButton("Save (minimize)", QMessageBox.AcceptRole)
        delete_btn = box.addButton("Delete", QMessageBox.DestructiveRole)
        cancel_btn = box.addButton(QMessageBox.Cancel)
        box.setDefaultButton(save_btn)
        box.exec()

        clicked = box.clickedButton()
        if clicked is cancel_btn:
            event.ignore()
            return
        if clicked is save_btn:
            event.ignore()
            sub = _find_subwindow(self)
            if sub:
                sub.showMinimized()
            return

        super().closeEvent(event)

    # ---------------- public API used by Workspace ----------------
    def create_empty(self, rows: int, cols: int) -> None:
        self._base_headers = [self._excel_col_name(i) for i in range(cols)]
        self.headers = list(self._base_headers)
        self.long_names = ["" for _ in range(cols)]
        self.names_model.update_names(self.long_names)
        self.data = np.full((rows, cols), np.nan, dtype=float)
        self._set_model(self.data, self.headers, structural=True)
        self.clear_roles()
        self._equalize_columns_from_col0()
        self.sheet_changed.emit()

    def ensure_size(self, rows: int, cols: int) -> None:
        cur_r = self.data.shape[0] if self.data is not None else 0
        cur_c = self.data.shape[1] if self.data is not None else 0
        new_r = max(cur_r, rows)
        new_c = max(cur_c, cols)

        if self.data is None:
            new_data = np.full((new_r, new_c), np.nan, dtype=float)
        else:
            new_data = np.full((new_r, new_c), np.nan, dtype=float)
            new_data[:cur_r, :cur_c] = self.data

        if len(self._base_headers) < new_c:
            self._base_headers.extend(self._excel_col_name(i) for i in range(len(self._base_headers), new_c))
        if len(self.long_names) < new_c:
            self.long_names.extend("" for _ in range(new_c - len(self.long_names)))
            self.names_model.update_names(self.long_names)

        self.data = new_data
        self._apply_role_labels()
        self._equalize_columns_from_col0()
        self.sheet_changed.emit()

    def overlay_csv(self, path: str) -> None:
        arr = self._read_csv_numeric(path)
        m, n = arr.shape
        self.ensure_size(max(1000, m), max(2, n))
        self.data[:m, :n] = arr
        self._apply_role_labels()
        self.setWindowTitle(f"Book [{os.path.basename(path)}]")
        self._emit_data_changed()
        self.sheet_changed.emit()

    def save_csv(self, path: str) -> None:
        with open(path, "w", newline="") as f:
            writer = csv.writer(f)
            header_row = [self.long_names[j] if self.long_names[j] else self._base_headers[j]
                          for j in range(len(self._base_headers))]
            writer.writerow(header_row)
            out = np.where(np.isnan(self.data), "", self.data).tolist()
            writer.writerows(out)

    def add_column(self, values: np.ndarray, name: Optional[str] = None) -> int:
        values = np.asarray(values, dtype=float).reshape(-1)
        if values.shape[0] != self.data.shape[0]:
            raise ValueError("Length of new column does not match row count")
        prev_cols = self.data.shape[1]
        self.ensure_size(self.data.shape[0], prev_cols + 1)
        j = prev_cols
        self.data[:, j] = values
        if name and name.strip():
            self._base_headers[j] = self._excel_col_name(j)
            self.long_names[j] = name.strip()
            self.names_model.update_names(self.long_names)
        self._apply_role_labels()
        self._emit_data_changed()
        self.sheet_changed.emit()
        return j

    def delete_columns(self, cols: List[int]) -> None:
        if not cols:
            return
        keep = [k for k in range(self.data.shape[1]) if k not in cols]
        if not keep:
            QMessageBox.warning(self, "Delete", "Cannot delete all columns.")
            return
        self.data = self.data[:, keep]
        self._base_headers = [self._base_headers[k] for k in keep]
        self.long_names = [self.long_names[k] for k in keep]
        self.names_model.update_names(self.long_names)
        self._reindex_roles_after_delete(set(cols))
        self._apply_role_labels()
        self._equalize_columns_from_col0()
        self._emit_data_changed()
        self.sheet_changed.emit()

    def rename_column(self, j: int, new_long_name: str) -> None:
        if 0 <= j < len(self.long_names):
            self.long_names[j] = new_long_name.strip()
            self.names_model.update_names(self.long_names)
            self._apply_role_labels()
            self.sheet_changed.emit()

    def get_selected_columns(self) -> List[int]:
        sel_model = self.table.selectionModel()
        if not sel_model:
            return []
        return sorted({ix.column() for ix in sel_model.selectedColumns()})

    def locals_map(self) -> Dict[str, np.ndarray]:
        return {self._excel_col_name(i): self.data[:, i] for i in range(self.data.shape[1])}

    # ----- roles -----
    def clear_roles(self) -> None:
        self.x_col = None
        self.y_cols = set()
        self.z_col = None
        self._apply_role_labels()

    def set_role_x(self, j: int) -> None:
        if 0 <= j < self.data.shape[1]:
            self.x_col = j
            self._apply_role_labels()
            self.sheet_changed.emit()

    def add_role_y(self, cols: List[int]) -> None:
        changed = False
        for j in cols:
            if 0 <= j < self.data.shape[1]:
                if j not in self.y_cols:
                    self.y_cols.add(j)
                    changed = True
        if changed:
            self._apply_role_labels()
            self.sheet_changed.emit()

    def set_role_z(self, j: int) -> None:
        if 0 <= j < self.data.shape[1]:
            self.z_col = j
            self._apply_role_labels()
            self.sheet_changed.emit()

    # ---------------- internals ----------------

    def _set_model(self, data: np.ndarray, headers: List[str], structural: bool = False) -> None:
        """Bind model and forward its change signals to sheet_changed."""
        if self._model is None:
            self._model = NumpyTableModel(data, headers)
            # Forward model signals
            try:
                self._model.dataChanged.connect(lambda *_: self.sheet_changed.emit())
                self._model.layoutChanged.connect(lambda *_: self.sheet_changed.emit())
                self._model.modelReset.connect(lambda *_: self.sheet_changed.emit())
            except Exception:
                pass
            self.table.setModel(self._model)
        else:
            self._model.update_all(data, headers)
            if structural:
                # a reset-equivalent
                try:
                    self._model.modelReset.emit()
                except Exception:
                    pass

    def _apply_role_labels(self) -> None:
        display = []
        for i, base in enumerate(self._base_headers):
            tags = []
            if self.x_col == i:
                tags.append("X")
            if i in self.y_cols:
                tags.append("Y")
            if self.z_col == i:
                tags.append("Z")
            display.append(f"{base} [{' '.join(tags)}]" if tags else base)
        self.headers = display
        self._set_model(self.data, self.headers)
        if len(self.long_names) != len(self.headers):
            self.long_names = (self.long_names + [""] * len(self.headers))[:len(self.headers)]
            self.names_model.update_names(self.long_names)

    def _reindex_roles_after_delete(self, deleted: set[int]) -> None:
        mapping: Dict[int, Optional[int]] = {}
        old_count = len(self._base_headers) + len(deleted)
        new_i = 0
        for i in range(old_count):
            if i in deleted:
                mapping[i] = None
            else:
                mapping[i] = new_i
                new_i += 1
        self.x_col = mapping.get(self.x_col) if self.x_col is not None else None
        self.z_col = mapping.get(self.z_col) if self.z_col is not None else None
        self.y_cols = {mapping[i] for i in self.y_cols if mapping.get(i) is not None}

    def _equalize_columns_from_col0(self):
        if self.table.model() is None:
            return
        header = self.table.horizontalHeader()
        names_header = self.names_view.horizontalHeader()
        w0 = header.sectionSize(0) if header.count() > 0 else 100
        header.setDefaultSectionSize(w0)
        names_header.setDefaultSectionSize(w0)
        for j in range(header.count()):
            header.resizeSection(j, w0)
            names_header.resizeSection(j, w0)

    def _sync_from_table_resize(self, logicalIndex: int, oldSize: int, newSize: int):
        if self._sync_guard:
            return
        self._sync_guard = True
        try:
            self.names_view.horizontalHeader().resizeSection(logicalIndex, newSize)
        finally:
            self._sync_guard = False

    def _sync_from_names_resize(self, logicalIndex: int, oldSize: int, newSize: int):
        if self._sync_guard:
            return
        self._sync_guard = True
        try:
            self.table.horizontalHeader().resizeSection(logicalIndex, newSize)
        finally:
            self._sync_guard = False

    # --- Clipboard / context menu ---
    def _make_clipboard_actions(self):
        self.act_copy = QAction("Copy", self); self.act_copy.setShortcut(QKeySequence.Copy)
        self.act_copy.triggered.connect(self._copy_selection)

        self.act_cut = QAction("Cut", self); self.act_cut.setShortcut(QKeySequence.Cut)
        self.act_cut.triggered.connect(self._cut_selection)

        self.act_paste = QAction("Paste", self); self.act_paste.setShortcut(QKeySequence.Paste)
        self.act_paste.triggered.connect(self._paste_from_clipboard)

        self.table.addAction(self.act_copy)
        self.table.addAction(self.act_cut)
        self.table.addAction(self.act_paste)

    def _on_context_menu(self, pos: QPoint):
        m = QMenu(self)
        m.addAction(self.act_copy)
        m.addAction(self.act_cut)
        m.addAction(self.act_paste)
        m.exec(self.table.viewport().mapToGlobal(pos))

    def _selected_rect(self) -> Optional[tuple[int, int, int, int]]:
        sel = self.table.selectionModel()
        if not sel or not sel.hasSelection():
            return None
        rows = [ix.row() for ix in sel.selectedIndexes()]
        cols = [ix.column() for ix in sel.selectedIndexes()]
        return min(rows), min(cols), max(rows), max(cols)

    def _copy_selection(self):
        rect = self._selected_rect()
        if rect is None:
            return
        r0, c0, r1, c1 = rect
        block = self.data[r0:r1+1, c0:c1+1]
        lines = []
        for row in block:
            vals = [("" if np.isnan(v) else f"{v:g}") for v in row]
            lines.append("\t".join(vals))
        QGuiApplication.clipboard().setText("\n".join(lines))

    def _cut_selection(self):
        rect = self._selected_rect()
        if rect is None:
            return
        self._copy_selection()
        r0, c0, r1, c1 = rect
        self.data[r0:r1+1, c0:c1+1] = np.nan
        self._set_model(self.data, self.headers)
        self._emit_data_changed()
        self.sheet_changed.emit()

    def _paste_from_clipboard(self):
        cb_text = QGuiApplication.clipboard().text()
        if not cb_text:
            return
        rows_text = [line for line in re.split(r"\r?\n", cb_text) if line != ""]
        if not rows_text:
            return
        matrix: List[List[float]] = []
        for line in rows_text:
            if "\t" in line:
                parts = line.split("\t")
            else:
                parts = re.split(r"[;,]", line)
            row_f = []
            for x in parts:
                xs = x.strip()
                if xs == "":
                    row_f.append(np.nan)
                else:
                    try:
                        row_f.append(float(xs))
                    except Exception:
                        row_f.append(np.nan)
            matrix.append(row_f)

        rect = self._selected_rect()
        r0 = rect[0] if rect else 0
        c0 = rect[1] if rect else 0

        nrows = len(matrix)
        ncols = max(len(r) for r in matrix)
        self.ensure_size(max(self.data.shape[0], r0 + nrows), max(self.data.shape[1], c0 + ncols))

        for i, row in enumerate(matrix):
            for j, val in enumerate(row):
                self.data[r0 + i, c0 + j] = val
        self._set_model(self.data, self.headers)
        self._emit_data_changed()
        self.sheet_changed.emit()

    def _emit_data_changed(self):
        """Lightweight nudge so views/plots refresh even if update_all didn't emit."""
        try:
            tl = self._model.index(0, 0)
            br = self._model.index(max(0, self.data.shape[0]-1), max(0, self.data.shape[1]-1))
            self._model.dataChanged.emit(tl, br)
        except Exception:
            pass

    @staticmethod
    def _read_csv_numeric(path: str) -> np.ndarray:
        with open(path, "r", newline="") as f:
            import csv as _csv
            sample = f.read(4096); f.seek(0)
            try:
                dialect = _csv.Sniffer().sniff(sample)
            except Exception:
                dialect = _csv.excel
            reader = _csv.reader(f, dialect)
            rows = list(reader)
        if not rows:
            raise ValueError("Empty file")
        def is_num(s: str) -> bool:
            try:
                float(s); return True
            except Exception:
                return False
        first = rows[0]
        numeric_first = sum(is_num(x) for x in first) / max(1, len(first)) >= 0.6
        data_rows = rows if numeric_first else rows[1:]
        parsed, widths = [], []
        for r in data_rows:
            try:
                parsed.append([float(x) for x in r])
            except Exception:
                parsed.append([float(x) if x.strip() != "" else np.nan for x in r])
            widths.append(len(parsed[-1]))
        if not parsed:
            raise ValueError("No data rows found")
        w = min(widths)
        A = np.asarray([row[:w] for row in parsed], dtype=float)
        keep = ~np.all(np.isnan(A), axis=1)
        A = A[keep]
        if A.size == 0:
            raise ValueError("No numeric cells after cleaning")
        return A

    @staticmethod
    def _excel_col_name(n: int) -> str:
        s = ""
        n = int(n)
        while True:
            n, r = divmod(n, 26)
            s = chr(ord('A') + r) + s
            if n == 0:
                return s
            n -= 1
