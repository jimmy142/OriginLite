# originlite/ui/plot_window.py
from __future__ import annotations

from typing import List, Optional, Dict

import numpy as np
import matplotlib.pyplot as plt
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QMessageBox, QMdiSubWindow
)
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure


def _find_subwindow(widget) -> QMdiSubWindow | None:
    w = widget.parentWidget()
    while w is not None and not isinstance(w, QMdiSubWindow):
        w = w.parentWidget()
    return w


def _find_workspace(widget):
    w = widget.parentWidget()
    while w is not None and not isinstance(w, QMainWindow):
        w = w.parentWidget()
    return w


class PlotWindow(QMainWindow):
    """
    Matplotlib plot window (dark style) with optional live link to a WorksheetWindow.
    Cleans up figure and signal connections so the Python process can exit cleanly.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Graph*")
        self.resize(820, 520)

        # Figure/axes with dark styling
        self.fig = Figure(figsize=(6, 4), dpi=100, facecolor="#262626")
        self.ax = self.fig.add_subplot(111)
        self.ax.set_facecolor("#1e1e1e")
        self.ax.tick_params(colors="#dddddd")
        for side in ("bottom", "left"):
            self.ax.spines[side].set_color("#aaaaaa")
        for side in ("top", "right"):
            self.ax.spines[side].set_visible(False)
        self.ax.grid(True, color="#333333", linewidth=0.6, alpha=0.6)

        self.canvas = FigureCanvas(self.fig)
        container = QWidget()
        lay = QVBoxLayout(container)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(self.canvas)
        self.setCentralWidget(container)

        # Live-link state
        self._src_ws = None
        self._x_idx: Optional[int] = None
        self._y_idxs: List[int] = []
        self._y_lines: Dict[int, object] = {}

    # ---------------------- public API ----------------------
    def add_line(self, x, y, label=None):
        self.ax.plot(x, y, label=label)

    def set_source(self, worksheet, x_index: int, y_indices: List[int]):
        if self._src_ws is not None:
            try:
                self._src_ws.sheet_changed.disconnect(self._on_sheet_changed)
            except Exception:
                pass

        self._src_ws = worksheet
        self._x_idx = int(x_index) if x_index is not None else None
        self._y_idxs = [int(j) for j in (y_indices or [])]
        self._y_lines.clear()
        self.ax.cla()
        self._apply_axes_style()

        self._build_or_update_lines(full_reset=True)

        if self._src_ws is not None:
            self._src_ws.sheet_changed.connect(self._on_sheet_changed)

        self.finish()

    def finish(self):
        handles, labels = self.ax.get_legend_handles_labels()
        if labels:
            self.ax.legend(frameon=False)
        self.canvas.draw_idle()

    # ---------------------- internals ----------------------
    def _apply_axes_style(self):
        self.ax.set_facecolor("#1e1e1e")
        self.ax.tick_params(colors="#dddddd")
        for side in ("bottom", "left"):
            self.ax.spines[side].set_color("#aaaaaa")
        for side in ("top", "right"):
            self.ax.spines[side].set_visible(False)
        self.ax.grid(True, color="#333333", linewidth=0.6, alpha=0.6)

    def _on_sheet_changed(self):
        self._build_or_update_lines(full_reset=False)
        self.finish()

    def _get_xy(self, xj: int, yj: int):
        ws = self._src_ws
        if ws is None or ws.data is None:
            return np.array([]), np.array([])
        if xj is None:
            return np.array([]), np.array([])
        if xj >= ws.data.shape[1] or yj >= ws.data.shape[1] or xj < 0 or yj < 0:
            return np.array([]), np.array([])
        x = ws.data[:, xj]
        y = ws.data[:, yj]
        mask = ~np.isnan(x) & ~np.isnan(y)
        if not np.any(mask):
            return np.array([]), np.array([])
        return x[mask], y[mask]

    def _series_label(self, yj: int) -> str:
        ws = self._src_ws
        if ws is None:
            return f"Y{yj+1}"
        name = ""
        try:
            name = ws.long_names[yj] if ws.long_names[yj] else ws.headers[yj]
        except Exception:
            name = f"Col {yj+1}"
        return name

    def _build_or_update_lines(self, full_reset: bool):
        if self._src_ws is None or self._x_idx is None or not self._y_idxs:
            return

        for yj in list(self._y_lines.keys()):
            if (yj not in self._y_idxs) or (yj >= self._src_ws.data.shape[1]):
                line = self._y_lines.pop(yj, None)
                if line in self.ax.lines:
                    try:
                        line.remove()
                    except Exception:
                        pass

        for yj in self._y_idxs:
            x, y = self._get_xy(self._x_idx, yj)
            if yj not in self._y_lines:
                (line_handle,) = self.ax.plot(x, y, label=self._series_label(yj))
                self._y_lines[yj] = line_handle
            else:
                line = self._y_lines[yj]
                line.set_data(x, y)
                line.set_label(self._series_label(yj))

        try:
            self.ax.relim()
            self.ax.autoscale_view()
        except Exception:
            pass

    # ---------------------- cleanup / close ----------------------
    def cleanup_for_quit(self):
        """Called by Workspace during app shutdown (no prompts)."""
        if self._src_ws is not None:
            try:
                self._src_ws.sheet_changed.disconnect(self._on_sheet_changed)
            except Exception:
                pass
            self._src_ws = None
        # Close the Matplotlib figure explicitly to release resources
        try:
            plt.close(self.fig)
        except Exception:
            pass
        try:
            self.canvas.deleteLater()
        except Exception:
            pass

    def closeEvent(self, event):
        # If the app is shutting down, skip prompts and clean quickly.
        ws = _find_workspace(self)
        if getattr(ws, "_shutting_down", False):
            self.cleanup_for_quit()
            super().closeEvent(event)
            return

        # Normal interactive close
        box = QMessageBox(self)
        box.setWindowTitle("Close Graph")
        box.setText("Do you want to save this graph or delete it?")
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

        # Delete chosen: detach and close figure to avoid lingering resources
        self.cleanup_for_quit()
        super().closeEvent(event)
