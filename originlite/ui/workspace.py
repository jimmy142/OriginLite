# originlite/ui/workspace.py
from __future__ import annotations

from typing import Optional, List, Dict, Tuple
import os
import io
import json
import zipfile
from datetime import datetime
from uuid import uuid4

import numpy as np

from PySide6.QtCore import Qt, QPoint, QRect, QObject, QTimer
try:
    from PySide6.QtCore import QSignalBlocker
except ImportError:
    class QSignalBlocker:
        def __init__(self, obj):
            self._obj = obj
            self._old = None
        def __enter__(self):
            try:
                self._old = self._obj.signalsBlocked()
                self._obj.blockSignals(True)
            except Exception:
                pass
        def __exit__(self, exc_type, exc, tb):
            try:
                self._obj.blockSignals(False if self._old is None else self._old)
            except Exception:
                pass

from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QMainWindow, QMdiArea, QMdiSubWindow, QToolBar, QFileDialog, QMessageBox, QMenuBar,
    QDockWidget, QTreeWidget, QTreeWidgetItem, QAbstractItemView, QMenu, QInputDialog
)

from .worksheet_window import WorksheetWindow
from .plot_window import PlotWindow
from .transform_dialog import TransformDialog
from ..data.eval import eval_expression

APP_VERSION = "0.1.0"
PROJECT_EXT = "olite"  # .olite is a zip container


class Workspace(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("OriginLite — Workspace")
        self.resize(1500, 900)

        self._shutting_down = False  # 🔒 used to bypass prompts and ignores during app exit

        self.mdi = QMdiArea()
        self.mdi.setViewMode(QMdiArea.SubWindowView)
        self.setCentralWidget(self.mdi)
        self.mdi.subWindowActivated.connect(self._on_subwindow_changed)

        self.project_path: Optional[str] = None

        self._uid_to_sub: Dict[str, QMdiSubWindow] = {}
        self._uid_to_item: Dict[str, QTreeWidgetItem] = {}

        self._explorer_guard = False

        self._build_explorer()
        self._build_menu_and_toolbar()

        self.new_worksheet()

    # ---------- Project Explorer ----------
    def _build_explorer(self):
        self.explorer_dock = QDockWidget("Project Explorer", self)
        self.explorer_dock.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        # Keep explorer always visible (no close)
        self.explorer_dock.setFeatures(self.explorer_dock.features() & ~QDockWidget.DockWidgetClosable)
        self.addDockWidget(Qt.LeftDockWidgetArea, self.explorer_dock)

        self.explorer = QTreeWidget()
        self.explorer.setHeaderHidden(True)
        self.explorer.setEditTriggers(QAbstractItemView.EditKeyPressed | QAbstractItemView.SelectedClicked)
        self.explorer.setContextMenuPolicy(Qt.CustomContextMenu)
        self.explorer.customContextMenuRequested.connect(self._explorer_context_menu)
        self.explorer.itemChanged.connect(self._explorer_item_changed)
        self.explorer.itemSelectionChanged.connect(self._explorer_selection_changed)
        self.explorer_dock.setWidget(self.explorer)

        self.root_ws = QTreeWidgetItem(self.explorer, ["Worksheets"])
        self.root_graphs = QTreeWidgetItem(self.explorer, ["Graphs"])
        self.root_ws.setExpanded(True)
        self.root_graphs.setExpanded(True)

    def _explorer_context_menu(self, pos: QPoint):
        item = self.explorer.itemAt(pos)
        if item is None or item in (self.root_ws, self.root_graphs):
            return
        m = QMenu(self)
        m.addAction("Rename", lambda: self.explorer.editItem(item, 0))
        m.addSeparator()
        m.addAction("Delete", lambda: self._explorer_delete_item(item))
        m.exec(self.explorer.viewport().mapToGlobal(pos))

    def _explorer_delete_item(self, item: QTreeWidgetItem):
        uid = item.data(0, Qt.UserRole)
        if not uid:
            return
        sub = self._uid_to_sub.get(uid)
        if sub:
            sub.close()  # child handles prompt unless shutting down
            QTimer.singleShot(0, self._purge_orphans)
        else:
            with QSignalBlocker(self.explorer):
                parent = item.parent() or self.explorer.invisibleRootItem()
                parent.removeChild(item)
            self._uid_to_item.pop(uid, None)
            self._uid_to_sub.pop(uid, None)

    def _explorer_item_changed(self, item: QTreeWidgetItem, col: int):
        if col != 0:
            return
        if item in (self.root_ws, self.root_graphs):
            item.setText(0, "Worksheets" if item is self.root_ws else "Graphs")
            return

        new_name = item.text(0).strip()
        uid = item.data(0, Qt.UserRole)

        if not new_name:
            sub = self._uid_to_sub.get(uid)
            if sub:
                item.setText(0, sub.windowTitle())
            return

        parent = item.parent()
        sibling_names = {parent.child(i).text(0).strip()
                         for i in range(parent.childCount()) if parent.child(i) is not item}
        if new_name in sibling_names:
            QMessageBox.warning(self, "Rename", f"An item named “{new_name}” already exists here.")
            sub = self._uid_to_sub.get(uid)
            if sub:
                item.setText(0, sub.windowTitle())
            return

        sub = self._uid_to_sub.get(uid)
        if sub:
            sub.setWindowTitle(new_name)

    def _explorer_selection_changed(self):
        if self._explorer_guard or self._shutting_down:
            return
        self._explorer_guard = True
        try:
            items = self.explorer.selectedItems()
            if not items:
                return
            item = items[0]
            uid = item.data(0, Qt.UserRole)
            sub = self._uid_to_sub.get(uid)

            if (sub is None) or (sub not in self.mdi.subWindowList()):
                with QSignalBlocker(self.explorer):
                    parent = item.parent() or self.explorer.invisibleRootItem()
                    parent.removeChild(item)
                self._uid_to_item.pop(uid, None)
                self._uid_to_sub.pop(uid, None)
                return

            self.mdi.setActiveSubWindow(sub)
        finally:
            self._explorer_guard = False

    def _explorer_add_subwindow(self, sub: QMdiSubWindow):
        uid = str(uuid4())
        sub.setProperty("olite_uid", uid)
        self._uid_to_sub[uid] = sub

        parent = self.root_ws if isinstance(sub.widget(), WorksheetWindow) else self.root_graphs
        label = self._unique_name_for_parent(parent, sub.windowTitle())
        sub.setWindowTitle(label)

        item = QTreeWidgetItem(parent, [label])
        item.setFlags(item.flags() | Qt.ItemIsEditable | Qt.ItemIsSelectable | Qt.ItemIsEnabled)
        item.setData(0, Qt.UserRole, uid)
        parent.setExpanded(True)
        self._uid_to_item[uid] = item

        with QSignalBlocker(self.explorer):
            self.explorer.setCurrentItem(item)

        sub.destroyed.connect(self._on_sub_destroyed)

    def _on_sub_destroyed(self, obj: QObject = None):
        sub = self.sender()
        if not isinstance(sub, QMdiSubWindow):
            return
        uid = sub.property("olite_uid")
        uid = str(uid) if uid is not None else None
        if not uid:
            return
        item = self._uid_to_item.pop(uid, None)
        if item:
            with QSignalBlocker(self.explorer):
                parent = item.parent() or self.explorer.invisibleRootItem()
                parent.removeChild(item)
        self._uid_to_sub.pop(uid, None)
        self.explorer.viewport().update()
        QTimer.singleShot(0, self._purge_orphans)

    def _purge_orphans(self):
        for uid, item in list(self._uid_to_item.items()):
            sub = self._uid_to_sub.get(uid)
            if (sub is None) or (sub not in self.mdi.subWindowList()):
                with QSignalBlocker(self.explorer):
                    parent = item.parent() or self.explorer.invisibleRootItem()
                    parent.removeChild(item)
                self._uid_to_item.pop(uid, None)
                self._uid_to_sub.pop(uid, None)
        self.explorer.viewport().update()

    def _unique_name_for_parent(self, parent: QTreeWidgetItem, base: str) -> str:
        existing = {parent.child(i).text(0).strip() for i in range(parent.childCount())}
        if base not in existing:
            return base
        n = 2
        while True:
            cand = f"{base} ({n})"
            if cand not in existing:
                return cand
            n += 1

    # ---------- Menu/Toolbar ----------
    def _build_menu_and_toolbar(self) -> None:
        self.menubar: QMenuBar = self.menuBar()
        self.toolbar = QToolBar("Main")
        self.addToolBar(self.toolbar)

        # Project
        self.menu_project = self.menubar.addMenu("Project")
        self.act_proj_new = QAction("New Project", self, triggered=self._proj_new)
        self.act_proj_open = QAction("Open Project…", self, triggered=self._proj_open)
        self.act_proj_save = QAction("Save Project", self, triggered=self._proj_save)
        self.act_proj_save_as = QAction("Save Project As…", self, triggered=self._proj_save_as)
        for a in (self.act_proj_new, self.act_proj_open, self.act_proj_save, self.act_proj_save_as):
            self.menu_project.addAction(a)

        # Worksheet
        self.menu_ws = self.menubar.addMenu("Worksheet")
        self.act_ws_new = QAction("New Worksheet", self, triggered=self.new_worksheet)
        self.act_ws_close = QAction("Close Active Worksheet", self, triggered=self._ws_close_active)
        self.menu_ws.addAction(self.act_ws_new)
        self.menu_ws.addAction(self.act_ws_close)

        # Data
        self.menu_data = self.menubar.addMenu("Data")
        self.act_data_import = QAction("Import Data (CSV)…", self, triggered=self._ws_import_csv)
        self.act_data_export = QAction("Export Data (CSV)…", self, triggered=self._ws_export_csv)
        self.act_add_expr = QAction("Add Column (Expression)…", self, triggered=self._ws_add_expr)
        self.act_add_empty_col = QAction("Add Empty Column…", self, triggered=self._ws_add_empty_column)
        self.act_del_col = QAction("Delete Selected Column(s)", self, triggered=self._ws_delete_cols)
        for a in (self.act_data_import, self.act_data_export, self.act_add_expr, self.act_add_empty_col, self.act_del_col):
            self.menu_data.addAction(a)

        self.toolbar.addAction(self.act_ws_new)
        self.toolbar.addSeparator()
        self.toolbar.addAction(self.act_data_import)
        self.toolbar.addAction(self.act_data_export)

        # Manipulation
        self.menu_manip = self.menubar.addMenu("Manipulation")
        self.act_arith = QAction("Column Arithmetic…", self, triggered=self._ws_column_arith)
        self.act_stats = QAction("Column Statistics", self, triggered=self._ws_column_stats)
        self.menu_manip.addAction(self.act_arith)
        self.menu_manip.addAction(self.act_stats)

        # Roles
        self.menu_roles = self.menubar.addMenu("Roles")
        self.act_role_x = QAction("Set as X (single)", self, triggered=self._ws_set_x)
        self.act_role_y = QAction("Set as Y (multi)", self, triggered=self._ws_set_y)
        self.act_role_z = QAction("Set as Z (single)", self, triggered=self._ws_set_z)
        self.act_role_clear = QAction("Clear Roles", self, triggered=self._ws_clear_roles)
        for a in (self.act_role_x, self.act_role_y, self.act_role_z, self.act_role_clear):
            self.menu_roles.addAction(a)

        # Plot
        self.menu_plot = self.menubar.addMenu("Plot")
        self.act_plot_line = QAction("Line", self, triggered=self._ws_plot_line)
        self.menu_plot.addAction(self.act_plot_line)

        self._refresh_action_states()

    # ---------- Close / Shutdown ----------
    def closeEvent(self, event):
        """
        Intercept main window close to perform clean shutdown (no prompts).
        """
        self.shutdown()
        event.accept()

    def shutdown(self):
        """
        Cleanly detach all signals, close subwindows without prompts, and
        release Matplotlib figures so the Python process can exit.
        """
        if self._shutting_down:
            return
        self._shutting_down = True

        # Stop reacting to MDI activation changes during teardown
        try:
            self.mdi.subWindowActivated.disconnect(self._on_subwindow_changed)
        except Exception:
            pass

        # Close all subwindows, telling children we're quitting
        for sub in list(self.mdi.subWindowList()):
            w = sub.widget()
            # If child has a special cleanup, run it (no dialogs)
            if hasattr(w, "cleanup_for_quit"):
                try:
                    w.cleanup_for_quit()
                except Exception:
                    pass
            # Ensure no delete-on-close surprises while we iterate
            sub.setAttribute(Qt.WA_DeleteOnClose, True)
            sub.close()

        # Explorer: block signals and clear mappings
        with QSignalBlocker(self.explorer):
            self.root_ws.takeChildren()
            self.root_graphs.takeChildren()
        self._uid_to_item.clear()
        self._uid_to_sub.clear()

    # ---------- Subwindow helpers ----------
    def new_worksheet(self):
        if self._shutting_down:
            return
        sub = QMdiSubWindow()
        win = WorksheetWindow(self)
        sub.setWidget(win)
        sub.setAttribute(Qt.WA_DeleteOnClose, True)
        sub.setWindowTitle("Book*")
        self.mdi.addSubWindow(sub)
        sub.resize(820, 520)
        sub.show()
        self._explorer_add_subwindow(sub)
        self._refresh_action_states()

    def _ws_close_active(self):
        sub = self._active_ws_sub()
        if sub:
            sub.close()
            QTimer.singleShot(0, self._purge_orphans)
        self._refresh_action_states()

    def _active_ws_sub(self) -> Optional[QMdiSubWindow]:
        sub = self.mdi.activeSubWindow()
        return sub if (sub and isinstance(sub.widget(), WorksheetWindow)) else None

    def _active_ws(self) -> Optional[WorksheetWindow]:
        sub = self._active_ws_sub()
        return sub.widget() if sub else None

    def _on_subwindow_changed(self, sub):
        if self._explorer_guard or self._shutting_down:
            return
        self._explorer_guard = True
        try:
            if sub:
                uid = sub.property("olite_uid")
                uid = str(uid) if uid is not None else None
                if uid:
                    item = self._uid_to_item.get(uid)
                    if item:
                        with QSignalBlocker(self.explorer):
                            self.explorer.setCurrentItem(item)
        finally:
            self._explorer_guard = False

        QTimer.singleShot(0, self._purge_orphans)
        self._refresh_action_states()

    def _refresh_action_states(self):
        ws = self._active_ws()
        has_ws = ws is not None
        for a in (
            self.act_ws_close, self.act_data_import, self.act_data_export,
            self.act_add_expr, self.act_add_empty_col, self.act_del_col, self.act_arith, self.act_stats,
            self.act_role_x, self.act_role_y, self.act_role_z, self.act_role_clear,
            self.act_plot_line, self.act_proj_save
        ):
            a.setEnabled(has_ws and not self._shutting_down)

    # ---------- Project ops ----------
    def _proj_new(self):
        if not self._confirm_discard_changes():
            return
        self._close_all_subwindows()
        self.project_path = None
        self.new_worksheet()

    def _proj_open(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Project", os.getcwd(), f"OriginLite Project (*.{PROJECT_EXT})"
        )
        if not path:
            return
        try:
            self._load_project(path)
            self.project_path = path
            self.statusBar().showMessage(f"Opened project: {os.path.basename(path)}", 5000)
        except Exception as e:
            QMessageBox.critical(self, "Open Project", f"Failed to open project:\n{e}")

    def _proj_save(self):
        if not self.project_path:
            self._proj_save_as()
            return
        try:
            self._save_project(self.project_path)
            self.statusBar().showMessage(f"Saved project: {os.path.basename(self.project_path)}", 5000)
        except Exception as e:
            QMessageBox.critical(self, "Save Project", f"Failed to save project:\n{e}")

    def _proj_save_as(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Project As...", os.getcwd(), f"OriginLite Project (*.{PROJECT_EXT})"
        )
        if not path:
            return
        if not path.lower().endswith(f".{PROJECT_EXT}"):
            path += f".{PROJECT_EXT}"
        try:
            self._save_project(path)
            self.project_path = path
            self.statusBar().showMessage(f"Saved project: {os.path.basename(self.project_path)}", 5000)
        except Exception as e:
            QMessageBox.critical(self, "Save Project", f"Failed to save project:\n{e}")

    def _confirm_discard_changes(self) -> bool:
        return True  # TODO: dirty tracking prompt

    def _close_all_subwindows(self):
        for sub in list(self.mdi.subWindowList()):
            sub.close()
        with QSignalBlocker(self.explorer):
            self.root_ws.takeChildren()
            self.root_graphs.takeChildren()
        self._uid_to_item.clear()
        self._uid_to_sub.clear()
        QTimer.singleShot(0, self._purge_orphans)

    # ---------- Data: Import/Export ----------
    def _ws_import_csv(self):
        ws = self._active_ws()
        if not ws:
            return
        path, _ = QFileDialog.getOpenFileName(self, "Import Data (CSV)", os.getcwd(), "CSV Files (*.csv)")
        if not path:
            return
        try:
            ws.overlay_csv(path)
        except Exception as e:
            QMessageBox.critical(self, "Import error", str(e))

    def _ws_export_csv(self):
        ws = self._active_ws()
        if not ws:
            return
        path, _ = QFileDialog.getSaveFileName(self, "Export Data (CSV)", os.getcwd(), "CSV Files (*.csv)")
        if not path:
            return
        try:
            ws.save_csv(path)
        except Exception as e:
            QMessageBox.critical(self, "Export error", str(e))

    # ---------- Data ops ----------
    def _ws_add_empty_column(self):
        ws = self._active_ws()
        if not ws:
            return
        name, ok = QInputDialog.getText(self, "Add Empty Column", "Name (optional):")
        if not ok:
            return
        try:
            blank = np.full(ws.data.shape[0], np.nan, dtype=float)
            ws.add_column(blank, name.strip())
        except Exception as e:
            QMessageBox.warning(self, "Add Column", str(e))

    def _ws_add_expr(self):
        ws = self._active_ws()
        if not ws:
            return
        dlg = TransformDialog(self, columns=ws.headers)
        if dlg.exec() != dlg.Accepted:
            return
        expr, name = dlg.get_values()
        if not expr:
            QMessageBox.information(self, "Missing", "Provide an expression.")
            return
        locals_map = ws.locals_map()
        try:
            new_col = np.asarray(eval_expression(expr, locals_map), dtype=float).reshape(-1)
        except Exception as e:
            QMessageBox.warning(self, "Expression error", str(e))
            return
        if new_col.shape[0] != ws.data.shape[0]:
            QMessageBox.warning(self, "Length error", "Resulting column has a different length.")
            return
        try:
            ws.add_column(new_col, name if name.strip() else None)
        except Exception as e:
            QMessageBox.warning(self, "Add column error", str(e))

    def _ws_delete_cols(self):
        ws = self._active_ws()
        if not ws:
            return
        cols = ws.get_selected_columns()
        if not cols:
            QMessageBox.information(self, "Delete", "Select at least one column (click header).")
            return
        ws.delete_columns(cols)

    # ---------- Manipulation ----------
    def _ws_column_arith_compute(self, A: np.ndarray, op: str, B: np.ndarray) -> np.ndarray:
        if op == "+": return A + B
        if op == "-": return A - B
        if op == "×": return A * B
        if op == "÷":
            if np.any(B == 0):
                raise ValueError("Division by zero")
            return A / B
        raise ValueError(f"Unknown op {op}")

    def _ws_column_arith_dialog(self, headers: List[str]) -> Optional[tuple[int, str, int, str]]:
        from PySide6.QtWidgets import QDialog, QComboBox, QFormLayout, QDialogButtonBox, QVBoxLayout, QHBoxLayout, QInputDialog, QLabel
        dlg = QDialog(self); dlg.setWindowTitle("Column Arithmetic")
        col_a = QComboBox(); col_a.addItems(headers)
        op = QComboBox(); op.addItems(["+", "-", "×", "÷"])
        col_b = QComboBox(); col_b.addItems(headers)
        form = QFormLayout()
        row = QHBoxLayout(); row.addWidget(col_a); row.addWidget(op); row.addWidget(col_b)
        form.addRow("A (op) B", row)
        hint = QLabel("You’ll be asked to name the result after OK."); hint.setWordWrap(True)
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        lay = QVBoxLayout(dlg); lay.addLayout(form); lay.addWidget(hint); lay.addWidget(btns)
        btns.accepted.connect(dlg.accept); btns.rejected.connect(dlg.reject)
        if dlg.exec() != dlg.Accepted:
            return None
        name, ok = QInputDialog.getText(self, "Result Column Name", "Name:")
        if not ok or not name.strip():
            return None
        return col_a.currentIndex(), op.currentText(), col_b.currentIndex(), name.strip()

    def _ws_column_arith(self):
        ws = self._active_ws()
        if not ws:
            return
        res = self._ws_column_arith_dialog(ws.headers)
        if not res:
            return
        a, op, b, name = res
        try:
            C = self._ws_column_arith_compute(ws.data[:, a], op, ws.data[:, b])
        except Exception as e:
            QMessageBox.warning(self, "Arithmetic error", str(e))
            return
        ws.add_column(C, name)

    def _ws_column_stats(self):
        ws = self._active_ws()
        if not ws:
            return
        cols = ws.get_selected_columns()
        if len(cols) != 1:
            QMessageBox.information(self, "Stats", "Select exactly one column.")
            return
        j = cols[0]
        x = ws.data[:, j]
        valid = ~np.isnan(x)
        if not np.any(valid):
            QMessageBox.information(self, "Column Statistics", "No numeric data in this column.")
            return
        xv = x[valid]
        msg = (
            f"Column: {ws.headers[j]}\n"
            f"Count: {xv.size}\n"
            f"Mean:  {np.mean(xv):.6g}\n"
            f"Std:   {np.std(xv, ddof=1):.6g}\n"
            f"Min:   {np.min(xv):.6g}\n"
            f"Max:   {np.max(xv):.6g}\n"
        )
        QMessageBox.information(self, "Column Statistics", msg)

    # ---------- Roles ----------
    def _ws_set_x(self):
        ws = self._active_ws()
        if not ws:
            return
        cols = ws.get_selected_columns()
        if not cols:
            QMessageBox.information(self, "Roles", "Select one column for X.")
            return
        ws.set_role_x(cols[0])

    def _ws_set_y(self):
        ws = self._active_ws()
        if not ws:
            return
        cols = ws.get_selected_columns()
        if not cols:
            QMessageBox.information(self, "Roles", "Select one or more columns for Y.")
            return
        ws.add_role_y(cols)

    def _ws_set_z(self):
        ws = self._active_ws()
        if not ws:
            return
        cols = ws.get_selected_columns()
        if not cols:
            QMessageBox.information(self, "Roles", "Select one column for Z.")
            return
        ws.set_role_z(cols[0])

    def _ws_clear_roles(self):
        ws = self._active_ws()
        if ws:
            ws.clear_roles()

    # ---------- Plot ----------
    def _ws_plot_line(self):
        ws = self._active_ws()
        if not ws:
            return
        xj = ws.x_col if ws.x_col is not None else (0 if ws.data.shape[1] >= 1 else None)
        yjs = list(ws.y_cols) if ws.y_cols else ([1] if ws.data.shape[1] >= 2 else [])
        if xj is None or not yjs:
            QMessageBox.information(self, "Plot", "Set roles (X and Y) first via Roles menu.")
            return

        sub = QMdiSubWindow()
        win = PlotWindow(self)  # live-linked plot window
        sub.setWidget(win); sub.setAttribute(Qt.WA_DeleteOnClose, True)
        sub.setWindowTitle("Graph*")
        self.mdi.addSubWindow(sub)
        sub.resize(820, 520); sub.show()
        self._explorer_add_subwindow(sub)

        # Link plot to worksheet; it will update on any sheet_changed
        win.set_source(ws, xj, yjs)

    # ---------- Save/Load project ----------
    # (unchanged below here)
    def _save_project(self, path: str):
        meta: Dict[str, object] = {
            "app_version": APP_VERSION,
            "saved_at": datetime.utcnow().isoformat() + "Z",
            "windows": []
        }

        subwindows = self.mdi.subWindowList()
        ws_count = 0

        with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for sub in subwindows:
                rect: QRect = sub.geometry()
                win_meta: Dict[str, object] = {
                    "type": "worksheet" if isinstance(sub.widget(), WorksheetWindow) else "graph",
                    "title": sub.windowTitle(),
                    "geometry": {"x": rect.x(), "y": rect.y(), "w": rect.width(), "h": rect.height()}
                }

                if isinstance(sub.widget(), WorksheetWindow):
                    ws_count += 1
                    ws: WorksheetWindow = sub.widget()
                    win_meta["long_names"] = list(ws.long_names)
                    win_meta["roles"] = {
                        "x": ws.x_col,
                        "y": sorted(list(ws.y_cols)),
                        "z": ws.z_col
                    }
                    csv_name = f"worksheets/W{ws_count:02d}-{_sanitize(sub.windowTitle())}.csv"
                    buf = io.StringIO()
                    import csv as _csv
                    writer = _csv.writer(buf)
                    header_row = [ws.long_names[j] if ws.long_names[j] else ws._base_headers[j]
                                  for j in range(len(ws._base_headers))]
                    writer.writerow(header_row)
                    out = np.where(np.isnan(ws.data), "", ws.data).tolist()
                    writer.writerows(out)
                    zf.writestr(csv_name, buf.getvalue())
                    win_meta["csv"] = csv_name
                else:
                    win_meta["series"] = []  # TODO: capture plotted series

                meta["windows"].append(win_meta)

            zf.writestr("project.json", json.dumps(meta, indent=2))

    def _load_project(self, path: str):
        if not zipfile.is_zipfile(path):
            raise ValueError("Not a valid .olite project")
        self._close_all_subwindows()

        with zipfile.ZipFile(path, "r") as zf:
            try:
                meta = json.loads(zf.read("project.json").decode("utf-8"))
            except Exception as e:
                raise ValueError(f"project.json missing or invalid: {e}")

            for win in meta.get("windows", []):
                if win.get("type") != "worksheet":
                    continue
                title = str(win.get("title", "Book*"))
                geom = win.get("geometry", {})
                sub = QMdiSubWindow()
                ws = WorksheetWindow(self)
                sub.setWidget(ws)
                sub.setAttribute(Qt.WA_DeleteOnClose, True)
                sub.setWindowTitle(title)
                self.mdi.addSubWindow(sub)
                try:
                    sub.setGeometry(QRect(int(geom.get("x", 50)), int(geom.get("y", 50)),
                                          int(geom.get("w", 820)), int(geom.get("h", 520))))
                except Exception:
                    pass
                sub.show()
                self._explorer_add_subwindow(sub)

                csv_name = win.get("csv")
                if csv_name and csv_name in zf.namelist():
                    data_np, header_row = _load_csv_from_zip(zf, csv_name)
                    rows, cols = data_np.shape
                    ws.ensure_size(rows, cols)
                    ws.data[:rows, :cols] = data_np
                    ws.long_names = header_row[:cols] + [""] * max(0, cols - len(header_row))
                    ws.names_model.update_names(ws.long_names)
                    roles = win.get("roles", {})
                    ws.x_col = roles.get("x", None)
                    ws.y_cols = set(roles.get("y", []))
                    ws.z_col = roles.get("z", None)
                    ws._apply_role_labels()

            for win in meta.get("windows", []):
                if win.get("type") != "graph":
                    continue
                title = str(win.get("title", "Graph*"))
                geom = win.get("geometry", {})
                sub = QMdiSubWindow()
                gw = PlotWindow(self)
                sub.setWidget(gw)
                sub.setAttribute(Qt.WA_DeleteOnClose, True)
                sub.setWindowTitle(title)
                self.mdi.addSubWindow(sub)
                try:
                    sub.setGeometry(QRect(int(geom.get("x", 60)), int(geom.get("y", 60)),
                                          int(geom.get("w", 820)), int(geom.get("h", 520))))
                except Exception:
                    pass
                sub.show()
                self._explorer_add_subwindow(sub)


# ==================== helpers ====================

def _sanitize(s: str) -> str:
    bad = '<>:"/\\|?*'
    return "".join("_" if c in bad else c for c in s).strip() or "untitled"


def _load_csv_from_zip(zf: zipfile.ZipFile, member: str):
    import csv as _csv
    with zf.open(member, "r") as f:
        text = io.TextIOWrapper(f, encoding="utf-8", newline="")
        reader = list(_csv.reader(text))
    if not reader:
        return np.zeros((0, 0), dtype=float), []
    header = reader[0]
    rows = reader[1:]

    parsed = []
    widths = []
    for r in rows:
        row_f = []
        for x in r:
            x = x.strip()
            if x == "":
                row_f.append(np.nan)
            else:
                try:
                    row_f.append(float(x))
                except Exception:
                    row_f.append(np.nan)
        parsed.append(row_f)
        widths.append(len(row_f))
    if not parsed:
        return np.zeros((0, 0), dtype=float), header
    w = min(widths)
    A = np.asarray([row[:w] for row in parsed], dtype=float)
    return A, header
