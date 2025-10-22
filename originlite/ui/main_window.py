# originlite/ui/main_window.py
from __future__ import annotations

import os
from typing import Optional, Tuple

import numpy as np
from PySide6.QtCore import Qt
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QMainWindow,
    QWidget,
    QFileDialog,
    QSplitter,
    QVBoxLayout,
    QToolBar,
    QStatusBar,
    QMessageBox,
)

from ..io.datatable import DataTable
from .plot_canvas import PlotCanvas
from .control_panel import ControlPanel
from .data_dock import DataDock
from .transform_dialog import TransformDialog
from ..plotting import models
from ..plotting.fitter import fit_and_predict
from ..data.eval import eval_expression


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("OriginLite — MVP")
        self.resize(1400, 800)

        # --- central area: controls + plot ---
        self.canvas = PlotCanvas(self)
        self.ctrl = ControlPanel()
        self.ctrl.setDisabled(True)

        splitter = QSplitter()
        splitter.addWidget(self.ctrl)
        splitter.addWidget(self.canvas)
        splitter.setStretchFactor(1, 1)

        central = QWidget()
        lay = QVBoxLayout(central)
        lay.addWidget(splitter)
        self.setCentralWidget(central)

        self.status = QStatusBar()
        self.setStatusBar(self.status)

        # Data state
        self.dataset: Optional[DataTable] = None
        self.current_xy: Optional[Tuple[np.ndarray, np.ndarray]] = None

        # --- Data dock (spreadsheet) ---
        self.data_dock = DataDock(self)
        self.addDockWidget(Qt.RightDockWidgetArea, self.data_dock)
        self.data_dock.set_callbacks(
            on_add_expr=self._on_add_expr_clicked,
            on_delete_column=self._on_delete_selected_column,
        )

        # Menus/toolbars
        self._build_toolbar()
        self._build_fit_menu()
        self._build_data_menu()

        # Signals
        self.ctrl.add_btn.clicked.connect(self.add_trace)

    # ---------------- UI builders ----------------
    def _build_toolbar(self) -> None:
        tb = QToolBar("Main")
        self.addToolBar(tb)

        open_act = QAction("Open CSV", self)
        open_act.triggered.connect(self.open_csv)
        tb.addAction(open_act)

        clear_act = QAction("Clear Plot", self)
        clear_act.triggered.connect(self.canvas.clear_axes)
        tb.addAction(clear_act)

        export_act = QAction("Export Figure…", self)
        export_act.triggered.connect(self.export_figure)
        tb.addAction(export_act)

    def _build_fit_menu(self) -> None:
        menubar = self.menuBar()
        fit_menu = menubar.addMenu("Fit")

        actions = [
            ("Linear", self.fit_linear),
            ("Exponential", self.fit_exponential),
            ("Gaussian", self.fit_gaussian),
            ("Lorentzian", self.fit_lorentzian),
            ("Voigt", self.fit_voigt),
        ]
        for name, fn in actions:
            act = QAction(name, self)
            act.triggered.connect(fn)
            fit_menu.addAction(act)

    def _build_data_menu(self) -> None:
        menubar = self.menuBar()
        data_menu = menubar.addMenu("Data")

        add_expr_act = QAction("Add Column from Expression…", self)
        add_expr_act.triggered.connect(self._on_add_expr_clicked)
        data_menu.addAction(add_expr_act)

        del_col_act = QAction("Delete Selected Column", self)
        del_col_act.triggered.connect(lambda: self._on_delete_selected_column(None))
        data_menu.addAction(del_col_act)

    # ---------------- File ops ----------------
    def open_csv(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Open data",
            os.getcwd(),
            "Data files (*.csv *.tsv *.txt);;All files (*.*)",
        )
        if not path:
            return
        try:
            table = DataTable.from_csv(path)
        except Exception as e:
            QMessageBox.critical(self, "Load error", str(e))
            return

        self.dataset = table
        self.ctrl.set_headers(table.headers)
        self.ctrl.setDisabled(False)
        self.data_dock.bind(table.data, table.headers)
        self.status.showMessage(
            f"Loaded: {os.path.basename(path)}  |  shape={table.data.shape}", 5000
        )

    # ---------------- Plotting ----------------
    def add_trace(self) -> None:
        if not self.dataset:
            return
        xi = self.ctrl.x_combo.currentIndex()
        yi = self.ctrl.y_combo.currentIndex()
        x = self.dataset.data[:, xi]
        y = self.dataset.data[:, yi]

        label = (
            self.ctrl.label_edit.text().strip()
            or f"{self.dataset.headers[yi]} vs {self.dataset.headers[xi]}"
        )
        ls = "-" if self.ctrl.line_chk.isChecked() else ""
        mk = "o" if self.ctrl.marker_chk.isChecked() else ""
        self.canvas.ax.plot(x, y, ls + mk, label=label)
        self.canvas.ax.legend(frameon=False)
        self.canvas.draw_idle()

        self.current_xy = (x, y)

    # ---------------- Export ----------------
    def export_figure(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Export figure",
            os.getcwd(),
            "Vector (*.svg *.pdf);;PNG (*.png)",
        )
        if not path:
            return
        try:
            if path.lower().endswith(".png"):
                self.canvas.fig.savefig(path, dpi=300)
            else:
                self.canvas.fig.savefig(path)
            self.status.showMessage(f"Saved: {path}", 5000)
        except Exception as e:
            QMessageBox.critical(self, "Export error", str(e))

    # ---------------- Fitting ----------------
    def _do_fit(self, model_fn, p0=None, label: str = "fit") -> None:
        if self.current_xy is None:
            QMessageBox.information(self, "No data", "Add a trace first.")
            return
        x, y = self.current_xy
        try:
            res = fit_and_predict(x, y, model_fn, p0=p0)
        except Exception as e:
            QMessageBox.warning(self, "Fit failed", str(e))
            return

        self.canvas.ax.plot(res.xfit, res.yfit, "--", label=label)
        self.canvas.ax.legend(frameon=False)
        self.canvas.draw_idle()
        self.status.showMessage(
            f"Params: {np.array2string(res.params, precision=4, separator=', ')}",
            10000,
        )

    def fit_linear(self) -> None:
        self._do_fit(models.linear, label="Linear fit")

    def fit_exponential(self) -> None:
        if self.current_xy is None:
            return
        x, y = self.current_xy
        self._do_fit(models.exponential, p0=models.guess_exponential(x, y), label="Exp fit")

    def fit_gaussian(self) -> None:
        if self.current_xy is None:
            return
        x, y = self.current_xy
        self._do_fit(models.gaussian, p0=models.guess_gaussian(x, y), label="Gaussian fit")

    def fit_lorentzian(self) -> None:
        if self.current_xy is None:
            return
        x, y = self.current_xy
        self._do_fit(models.lorentzian, p0=models.guess_lorentzian(x, y), label="Lorentzian fit")

    def fit_voigt(self) -> None:
        if self.current_xy is None:
            return
        x, y = self.current_xy
        self._do_fit(models.voigt, p0=models.guess_voigt(x, y), label="Voigt fit")

    # ---------------- Data transforms ----------------
    def _on_add_expr_clicked(self) -> None:
        if not self.dataset:
            QMessageBox.information(self, "No data", "Load a table first.")
            return
        dlg = TransformDialog(self, columns=self.dataset.headers)
        if dlg.exec() != dlg.Accepted:
            return
        expr, name = dlg.get_values()
        if not expr or not name:
            QMessageBox.information(self, "Missing", "Provide both expression and new column name.")
            return

        locals_map = {h: self.dataset.data[:, i] for i, h in enumerate(self.dataset.headers)}
        try:
            new_col = eval_expression(expr, locals_map)
        except Exception as e:
            QMessageBox.warning(self, "Expression error", str(e))
            return

        try:
            self.dataset.add_column(name, new_col)
        except Exception as e:
            QMessageBox.warning(self, "Add column error", str(e))
            return

        # refresh views
        self.data_dock.bind(self.dataset.data, self.dataset.headers)
        self.ctrl.set_headers(self.dataset.headers)

    def _on_delete_selected_column(self, col: Optional[int]) -> None:
        if not self.dataset:
            return
        if col is None:
            col = self.data_dock.current_column()
        if col is None:
            QMessageBox.information(self, "Delete Column", "Select a column (click a cell) first.")
            return

        self.dataset.delete_column(col)
        self.data_dock.bind(self.dataset.data, self.dataset.headers)
        self.ctrl.set_headers(self.dataset.headers)
