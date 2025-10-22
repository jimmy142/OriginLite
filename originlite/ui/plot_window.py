# originlite/ui/plot_window.py
from __future__ import annotations

from typing import List, Optional, Dict

import numpy as np
import matplotlib.pyplot as plt
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QMessageBox, QMdiSubWindow, QFileDialog, QMenu, QInputDialog
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QAction
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qtagg import NavigationToolbar2QT
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

        # Figure/axes with light (paper-friendly) styling by default
        self.fig = Figure(figsize=(6, 4), dpi=100, facecolor="#ffffff")
        self.ax = self.fig.add_subplot(111)
        self.ax.set_facecolor("#ffffff")
        self.ax.tick_params(colors="#000000")
        for side in ("bottom", "left", "top", "right"):
            self.ax.spines[side].set_color("#666666")
        self.ax.grid(True, color="#cccccc", linewidth=0.6, alpha=1.0)

        self.canvas = FigureCanvas(self.fig)
        self.toolbar = NavigationToolbar2QT(self.canvas, self)
        container = QWidget()
        lay = QVBoxLayout(container)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(self.toolbar)
        lay.addWidget(self.canvas)
        self.setCentralWidget(container)

        # Right-click context menu on the canvas
        try:
            self.canvas.setContextMenuPolicy(Qt.CustomContextMenu)
            self.canvas.customContextMenuRequested.connect(self._on_context_menu)
        except Exception:
            pass

        # Live-link state
        self._src_ws = None
        self._x_idx: Optional[int] = None
        self._y_idxs: List[int] = []
        self._y_lines: Dict[int, object] = {}
        self._mode: str = "line"  # line, line_markers, bar, double_y, pie, surface3d
        self._opts: Dict[str, object] = {}
        self._ax_right = None  # for double_y
        self._ax_top = None    # for independent top axis (twiny)
        self._props: Dict[str, object] = {}
        self._data_cursor_enabled = False
        self._hover_annot = None

        # Menu
        m = self.menuBar().addMenu("Graph")
        act_props = QAction("Properties...", self)
        act_props.triggered.connect(self._open_properties)
        m.addAction(act_props)
        m.addSeparator()
        self.act_add_trace = QAction("Add Trace...", self)
        self.act_add_trace.triggered.connect(self._add_trace_dialog)
        m.addAction(self.act_add_trace)
        m.addSeparator()
        self.act_data_cursor = QAction("Data Cursor", self)
        self.act_data_cursor.setCheckable(True)
        self.act_data_cursor.triggered.connect(self._toggle_data_cursor)
        m.addAction(self.act_data_cursor)
        m.addSeparator()
        exp_png = QAction("Export PNG...", self); exp_png.triggered.connect(lambda: self._export_figure('png'))
        exp_svg = QAction("Export SVG...", self); exp_svg.triggered.connect(lambda: self._export_figure('svg'))
        exp_pdf = QAction("Export PDF...", self); exp_pdf.triggered.connect(lambda: self._export_figure('pdf'))
        m.addAction(exp_png); m.addAction(exp_svg); m.addAction(exp_pdf)
        m.addSeparator()
        act_axes = QAction("Axis && Ticks...", self); act_axes.triggered.connect(self._open_axes_dialog)
        m.addAction(act_axes)

        # Canvas picking for double-click on line artists
        # Enable picking on future lines by setting picker when we create them

    # ---------------------- public API ----------------------
    def add_line(self, x, y, label=None):
        self.ax.plot(x, y, label=label)

    def set_source(self, worksheet, x_index: int, y_indices: List[int], *, mode: str = "line", **opts):
        if self._src_ws is not None:
            try:
                self._src_ws.sheet_changed.disconnect(self._on_sheet_changed)
            except Exception:
                pass

        self._src_ws = worksheet
        self._x_idx = int(x_index) if x_index is not None else None
        self._y_idxs = [int(j) for j in (y_indices or [])]
        self._y_lines.clear()
        self._mode = str(mode or "line")
        self._opts = dict(opts) if opts else {}
        self._reset_axes()

        if self._mode in ("line", "line_markers"):
            self._build_or_update_lines(full_reset=True)
        else:
            self._replot_full()
        self._apply_saved_props()

        if self._src_ws is not None:
            self._src_ws.sheet_changed.connect(self._on_sheet_changed)

        self.finish()

    def finish(self):
        # Apply legend options if any; otherwise fallback to default simple legend
        if isinstance(self._props.get('legend'), dict):
            self._apply_legend_opts()
        else:
            handles, labels = self.ax.get_legend_handles_labels()
            if labels:
                self.ax.legend(frameon=False)
        self.canvas.draw_idle()

    # ---------------------- internals ----------------------
    def _apply_axes_style(self):
        self.ax.set_facecolor("#ffffff")
        self.ax.tick_params(colors="#000000")
        for side in ("bottom", "left", "top", "right"):
            self.ax.spines[side].set_color("#666666")
            self.ax.spines[side].set_visible(True)
        self.ax.grid(True, color="#cccccc", linewidth=0.6, alpha=1.0)

    def _reset_axes(self):
        # Clear and (re)configure axes based on mode
        self.fig.clf()
        self._ax_right = None
        self._ax_top = None
        if self._mode == "surface3d":
            from mpl_toolkits.mplot3d import Axes3D  # noqa: F401
            self.ax = self.fig.add_subplot(111, projection='3d')
            self.ax.set_facecolor("#ffffff")
        else:
            self.ax = self.fig.add_subplot(111)
            self._apply_axes_style()
            if self._mode == "double_y":
                self._ax_right = self.ax.twinx()
                # style right axis
                rc = "#000000"
                self._ax_right.tick_params(colors=rc)
                self._ax_right.spines["top"].set_visible(True)
        # init annotation for data cursor
        try:
            if self._hover_annot is None:
                self._hover_annot = self.ax.annotate("", xy=(0, 0), xytext=(10, 10), textcoords="offset points",
                                                    color="#000000", bbox=dict(boxstyle="round", fc="#ffffffCC", ec="#666"))
            self._hover_annot.set_visible(False)
        except Exception:
            self._hover_annot = None
        # Make key text artists pickable for double-click
        try:
            self.ax.xaxis.label.set_picker(True)
            self.ax.yaxis.label.set_picker(True)
            if self._ax_right is not None:
                self._ax_right.yaxis.label.set_picker(True)
        except Exception:
            pass

    def _on_sheet_changed(self):
        if self._mode in ("line", "line_markers"):
            self._build_or_update_lines(full_reset=False)
        else:
            self._replot_full()
        self._apply_saved_props()
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

    def _label_for_y(self, yj: int) -> str:
        try:
            lmap = self._props.get('legend_labels') or {}
            if isinstance(lmap, dict):
                v = lmap.get(int(yj)) or lmap.get(str(yj))
                if v:
                    return str(v)
        except Exception:
            pass
        return self._series_label(yj)

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
                linespec = "-" if self._mode == "line" else "-o"
                (line_handle,) = self.ax.plot(x, y, linespec, label=self._label_for_y(yj))
                try:
                    line_handle.set_picker(5)
                except Exception:
                    pass
                self._y_lines[yj] = line_handle
            else:
                line = self._y_lines[yj]
                line.set_data(x, y)
                line.set_label(self._label_for_y(yj))

        try:
            self.ax.relim()
            self.ax.autoscale_view()
        except Exception:
            pass

    # ---------------------- full replot modes ----------------------
    def _replot_full(self):
        self.ax.cla()
        self._y_lines.clear()
        if self._mode != "surface3d":
            self._apply_axes_style()
        if self._src_ws is None or self._x_idx is None or not self._y_idxs:
            return

        if self._mode == "bar":
            self._plot_bar()
        elif self._mode == "double_y":
            self._plot_double_y()
        elif self._mode == "pie":
            self._plot_pie()
        elif self._mode == "scatter":
            self._plot_scatter()
        elif self._mode == "stem":
            self._plot_stem()
        elif self._mode == "errorbar":
            self._plot_errorbar()
        elif self._mode == "surface3d":
            self._plot_surface3d()
        else:
            # fallback: simple line
            for yj in self._y_idxs:
                x, y = self._get_xy(self._x_idx, yj)
                self.ax.plot(x, y, label=self._series_label(yj))
        try:
            if self._mode != "pie":
                self.ax.relim(); self.ax.autoscale_view()
        except Exception:
            pass

    def _plot_bar(self):
        # Simple grouped bars per Y series at X positions
        ws = self._src_ws
        x = ws.data[:, self._x_idx]
        maskx = ~np.isnan(x)
        x = x[maskx]
        n = len(self._y_idxs)
        if n == 0 or x.size == 0:
            return
        width = 0.8 / max(1, n)
        for k, yj in enumerate(self._y_idxs):
            y = ws.data[:, yj]
            y = y[maskx]
            offs = (k - (n-1)/2) * width
            self.ax.bar(x + offs, y, width=width, label=self._series_label(yj), align='center')

    def _plot_double_y(self):
        # First Y series on left, second (if any) on right
        ws = self._src_ws
        if len(self._y_idxs) == 1:
            x, y = self._get_xy(self._x_idx, self._y_idxs[0])
            (lh,) = self.ax.plot(x, y, label=self._series_label(self._y_idxs[0]))
            self._y_lines[self._y_idxs[0]] = lh
            return
        left_y = self._y_idxs[0]
        right_ys = self._y_idxs[1:]
        x, y = self._get_xy(self._x_idx, left_y)
        (lh_left,) = self.ax.plot(x, y, label=self._label_for_y(left_y))
        try:
            lh_left.set_picker(5)
        except Exception:
            pass
        self._y_lines[left_y] = lh_left
        if self._ax_right is None:
            self._ax_right = self.ax.twinx()
        for yj in right_ys:
            xr, yr = self._get_xy(self._x_idx, yj)
            (lh_r,) = self._ax_right.plot(xr, yr, '--', label=self._label_for_y(yj))
            try:
                lh_r.set_picker(5)
            except Exception:
                pass
            self._y_lines[yj] = lh_r
        # Merge legends
        h1, l1 = self.ax.get_legend_handles_labels()
        h2, l2 = (self._ax_right.get_legend_handles_labels() if self._ax_right else ([], []))
        if l1 or l2:
            leg = self.ax.legend(h1 + h2, l1 + l2, frameon=False)

    def _plot_pie(self):
        # Use first Y series values; X values as labels (stringified)
        ws = self._src_ws
        yj = self._y_idxs[0]
        x = ws.data[:, self._x_idx]
        y = ws.data[:, yj]
        mask = ~np.isnan(x) & ~np.isnan(y) & (y > 0)
        x = x[mask]; y = y[mask]
        if y.size == 0:
            return
        labels = [f"{v:g}" for v in x]
        self.ax.pie(y, labels=labels, autopct='%1.1f%%', textprops={"color": "#dddddd"})
        self.ax.axis('equal')

    def _plot_surface3d(self):
        # Triangulated surface if X,Y,Z roles available; if only one Y, we use it.
        ws = self._src_ws
        z_col = getattr(ws, 'z_col', None)
        if z_col is None:
            # fallback to first Y as Z and X as X, and second Y as Y if available
            if len(self._y_idxs) >= 2:
                xi, yi, zi = self._x_idx, self._y_idxs[0], self._y_idxs[1]
                x = ws.data[:, xi]; y = ws.data[:, yi]; z = ws.data[:, zi]
            else:
                return
        else:
            yi = self._y_idxs[0] if self._y_idxs else None
            if yi is None:
                return
            x = ws.data[:, self._x_idx]; y = ws.data[:, yi]; z = ws.data[:, z_col]
        mask = ~np.isnan(x) & ~np.isnan(y) & ~np.isnan(z)
        x = x[mask]; y = y[mask]; z = z[mask]
        if x.size < 3:
            return
        try:
            surf = self.ax.plot_trisurf(x, y, z, cmap='viridis', linewidth=0.2, antialiased=True)
            self.fig.colorbar(surf, ax=self.ax, shrink=0.6, pad=0.1)
        except Exception:
            # fallback to scatter
            self.ax.scatter(x, y, z, s=10, alpha=0.8)

    def _plot_scatter(self):
        ws = self._src_ws
        for yj in self._y_idxs:
            x, y = self._get_xy(self._x_idx, yj)
            (sc,) = self.ax.plot(x, y, linestyle="", marker="o", label=self._label_for_y(yj))
            try:
                sc.set_picker(5)
            except Exception:
                pass
            self._y_lines[yj] = sc

    def _plot_stem(self):
        ws = self._src_ws
        for yj in self._y_idxs:
            x, y = self._get_xy(self._x_idx, yj)
            markerline, stemlines, baseline = self.ax.stem(x, y, label=self._label_for_y(yj))
            markerline.set_marker('o')
            # store markerline as representative handle
            self._y_lines[yj] = markerline

    def _plot_errorbar(self):
        # If multiple Y columns selected, use the second as yerr for first
        if len(self._y_idxs) == 0:
            return
        x, y = self._get_xy(self._x_idx, self._y_idxs[0])
        yerr = None
        if len(self._y_idxs) >= 2:
            _, y2 = self._get_xy(self._x_idx, self._y_idxs[1])
            yerr = np.abs(y2)
        (eh,) = self.ax.errorbar(x, y, yerr=yerr, fmt='o-', label=self._label_for_y(self._y_idxs[0]))
        try:
            eh.set_picker(5)
        except Exception:
            pass
        self._y_lines[self._y_idxs[0]] = eh

    # ---------------------- interactions ----------------------
    def _toggle_data_cursor(self, checked: bool):
        self._data_cursor_enabled = bool(checked)
        cid = getattr(self, '_cid_click', None)
        if self._data_cursor_enabled:
            if cid is None:
                self._cid_click = self.canvas.mpl_connect('button_press_event', self._on_click)
                self._cid_pick = self.canvas.mpl_connect('pick_event', self._on_pick)
        else:
            if cid is not None:
                try:
                    self.canvas.mpl_disconnect(self._cid_click)
                except Exception:
                    pass
                self._cid_click = None
            try:
                self.canvas.mpl_disconnect(getattr(self, '_cid_pick', None))
            except Exception:
                pass
            self._cid_pick = None
            if self._hover_annot is not None:
                self._hover_annot.set_visible(False)
                self.canvas.draw_idle()

    def _on_click(self, event):
        # Double-click handling: if a line is double-clicked, handled in pick; else open properties
        if not event.inaxes:
            return
        if event.dblclick and event.button == 1:
            # If not on an artist, open properties dialog as a shortcut
            self._open_properties()
            return
        if not self._data_cursor_enabled or event.button != 1:
            return
        self._update_annotation_nearest(event)

    def _on_pick(self, event):
        # Double-click on a line opens style dialog
        if getattr(event, 'mouseevent', None) is None:
            return
        me = event.mouseevent
        if me.dblclick and hasattr(event, 'artist'):
            art = event.artist
            try:
                from matplotlib.legend import Legend as _Legend
                from matplotlib.text import Text as _Text
            except Exception:
                _Legend = None; _Text = None
            # Legend double-click -> legend dialog
            if _Legend is not None and isinstance(art, _Legend):
                self._open_legend_dialog()
                return
            # Axis label/text double-click -> axes dialog
            if _Text is not None and isinstance(art, _Text):
                self._open_axes_dialog()
                return
            # Fallback: line style editor
            self._edit_line_style(art)

    def _update_annotation_nearest(self, event):
        if not event.inaxes:
            return
        best = None
        for line in list(self.ax.lines) + (list(self._ax_right.lines) if self._ax_right else []):
            try:
                xd, yd = line.get_xdata(), line.get_ydata()
            except Exception:
                continue
            if len(xd) == 0:
                continue
            # find nearest index by data distance in display coords
            try:
                trans = line.axes.transData
                pts = trans.transform(np.column_stack([xd, yd]))
                dist = np.hypot(pts[:,0]-event.x, pts[:,1]-event.y)
                i = int(np.nanargmin(dist))
                d = float(dist[i])
                if best is None or d < best[0]:
                    best = (d, line, xd[i], yd[i])
            except Exception:
                pass
        if best is None:
            return
        _, line, x, y = best
        if self._hover_annot is not None:
            self._hover_annot.xy = (x, y)
            self._hover_annot.set_text(f"x={x:g}\ny={y:g}")
            self._hover_annot.set_visible(True)
            self.canvas.draw_idle()
        ws = _find_workspace(self)
        if ws is not None and hasattr(ws, 'statusBar'):
            try:
                ws.statusBar().showMessage(f"Data cursor: x={x:g}, y={y:g}", 4000)
            except Exception:
                pass

    def _edit_line_style(self, artist):
        try:
            from .plot_line_dialog import LineStyleDialog
        except Exception:
            return
        color = None
        lw = None
        try:
            color = artist.get_color()
            lw = float(artist.get_linewidth())
        except Exception:
            pass
        dlg = LineStyleDialog(self, initial_color=color, initial_width=lw)
        from PySide6.QtWidgets import QDialog
        if dlg.exec() != QDialog.Accepted:
            return
        style = dlg.get_values()
        try:
            if style.get('color'):
                artist.set_color(style['color'])
            if style.get('linewidth'):
                artist.set_linewidth(float(style['linewidth']))
        except Exception:
            pass
        self.canvas.draw_idle()

    def _add_trace_dialog(self):
        try:
            from .add_trace_dialog import AddTraceDialog
            from .worksheet_window import WorksheetWindow
        except Exception:
            return
        ws_widget = _find_workspace(self)
        if ws_widget is None:
            return
        worksheets = []
        for sub in ws_widget.mdi.subWindowList():
            w = sub.widget()
            if isinstance(w, WorksheetWindow):
                worksheets.append((sub.windowTitle(), w))
        if not worksheets:
            QMessageBox.information(self, "Add Trace", "No worksheets available.")
            return
        dlg = AddTraceDialog(self, worksheets)
        from PySide6.QtWidgets import QDialog
        if dlg.exec() != QDialog.Accepted:
            return
        vals = dlg.values()
        if not vals:
            return
        ws, xi, yjs, label = vals
        if ws is None or xi is None or not yjs:
            return
        for yj in yjs:
            x = ws.data[:, int(xi)]
            y = ws.data[:, int(yj)]
            mask = ~np.isnan(x) & ~np.isnan(y)
            x = x[mask]; y = y[mask]
            lab = label or f"{ws.headers[yj]} vs {ws.headers[xi]}"
            (lh,) = self.ax.plot(x, y, '-', label=lab)
            try:
                lh.set_picker(5)
            except Exception:
                pass
        self.ax.legend(frameon=False)
        self.canvas.draw_idle()

    # ---------------------- context menu ----------------------
    def _on_context_menu(self, pos):
        menu = QMenu(self)
        a_props = menu.addAction("Properties...", self._open_properties)
        a_add = menu.addAction("Add Trace...", self._add_trace_dialog)
        menu.addSeparator()
        a_cursor = menu.addAction("Data Cursor")
        a_cursor.setCheckable(True)
        a_cursor.setChecked(self._data_cursor_enabled)
        a_cursor.triggered.connect(lambda checked: self._toggle_data_cursor(checked))
        a_grid = menu.addAction("Toggle Grid", lambda: self._toggle_grid_action())
        menu.addAction("Legend...", self._open_legend_dialog)
        menu.addAction("Axis && Ticks...", self._open_axes_dialog)
        menu.addSeparator()
        a_edit_line = None
        if self._y_lines:
            a_edit_line = menu.addAction("Edit Line Style...", self._edit_line_style_menu)
        menu.addSeparator()
        exp = menu.addMenu("Export")
        exp.addAction("PNG...", lambda: self._export_figure('png'))
        exp.addAction("SVG...", lambda: self._export_figure('svg'))
        exp.addAction("PDF...", lambda: self._export_figure('pdf'))
        menu.exec(self.canvas.mapToGlobal(pos))

    def _edit_line_style_menu(self):
        # Let user choose which series (any current line on axes)
        lines = list(self.ax.lines) + (list(self._ax_right.lines) if self._ax_right else [])
        if not lines:
            return
        entries = [(ln.get_label() or "series") for ln in lines]
        sel, ok = QInputDialog.getItem(self, "Edit Line Style", "Series:", entries, 0, False)
        if not ok:
            return
        try:
            idx = entries.index(sel)
        except ValueError:
            return
        artist = lines[idx]
        self._edit_line_style(artist)

    def _toggle_grid_action(self):
        # Toggle and store in props
        g = not bool(self._props.get('grid', True))
        self._props['grid'] = g
        try:
            self.ax.grid(g)
        except Exception:
            pass
        self.canvas.draw_idle()

    def _export_figure(self, kind: str):
        kind = str(kind).lower()
        flt = {
            'png': 'PNG (*.png)',
            'svg': 'SVG (*.svg)',
            'pdf': 'PDF (*.pdf)'
        }.get(kind, 'PNG (*.png)')
        path, _ = QFileDialog.getSaveFileName(self, f"Export {kind.upper()}", '', flt)
        if not path:
            return
        try:
            if kind == 'png' and not path.lower().endswith('.png'):
                path += '.png'
            if kind == 'svg' and not path.lower().endswith('.svg'):
                path += '.svg'
            if kind == 'pdf' and not path.lower().endswith('.pdf'):
                path += '.pdf'
        except Exception:
            pass
        try:
            if kind == 'png':
                self.fig.savefig(path, dpi=300)
            else:
                self.fig.savefig(path)
        except Exception as e:
            QMessageBox.warning(self, "Export", str(e))

    # ---------------------- legend dialog ----------------------
    def _open_legend_dialog(self):
        try:
            from .legend_dialog import LegendDialog
        except Exception:
            return
        # Gather series labels and mapping to y-index if possible
        lines = list(self.ax.lines) + (list(self._ax_right.lines) if self._ax_right else [])
        series = [(ln.get_label() or "series", ln.get_label() or "") for ln in lines]
        opts = self._props.get('legend', {}) if isinstance(self._props.get('legend'), dict) else {}
        dlg = LegendDialog(self, series=series, opts=opts)
        from PySide6.QtWidgets import QDialog
        if dlg.exec() != QDialog.Accepted:
            return
        vals = dlg.get_values()
        # Apply labels directly to artists
        new_labels = vals.pop('labels', [])
        for ln, label in zip(lines, new_labels):
            try:
                if label:
                    ln.set_label(label)
            except Exception:
                pass
        # Persist label overrides for y-linked series
        lab_map: Dict[int, str] = {}
        inv = {v: k for (k, v) in self._y_lines.items()}
        for ln, label in zip(lines, new_labels):
            yj = inv.get(ln)
            if yj is not None and label:
                lab_map[int(yj)] = label
        if lab_map:
            self._props['legend_labels'] = lab_map
        # Save legend options and apply
        self._props['legend'] = vals
        self._apply_legend_opts()
        self.canvas.draw_idle()

    def _apply_legend_opts(self):
        p = self._props or {}
        legopt = p.get('legend') or {}
        # Update labels for y-linked lines if overrides exist
        lab_map = p.get('legend_labels') or {}
        if isinstance(lab_map, dict):
            for yj, ln in list(self._y_lines.items()):
                lbl = lab_map.get(int(yj)) or lab_map.get(str(yj))
                if lbl:
                    try:
                        ln.set_label(str(lbl))
                    except Exception:
                        pass
        # Remove existing legend
        try:
            lg = self.ax.get_legend()
            if lg is not None:
                lg.remove()
        except Exception:
            pass
        if not bool(legopt.get('visible', True)):
            return
        # Build legend with options
        try:
            loc = str(legopt.get('loc', 'best'))
            ncol = int(legopt.get('ncol', 1))
            ff = legopt.get('fontfamily')
            fs = legopt.get('fontsize')
            prop = None
            if ff or fs:
                from matplotlib import font_manager as _fm
                prop = {'family': ff} if ff else {}
                if fs:
                    prop['size'] = fs
            lg = self.ax.legend(loc=loc, ncol=ncol, frameon=bool(legopt.get('frameon', False)), prop=prop)
            frm = lg.get_frame()
            ec = legopt.get('edgecolor'); fc = legopt.get('facecolor'); ew = legopt.get('edgewidth')
            if fc:
                try: frm.set_facecolor(fc)
                except Exception: pass
            if ec:
                try: frm.set_edgecolor(ec)
                except Exception: pass
            if ew:
                try: frm.set_linewidth(float(ew))
                except Exception: pass
            try:
                lg.set_draggable(True)
                lg.set_picker(True)
            except Exception:
                pass
        except Exception:
            pass

    # ---------------------- axis & ticks dialog ----------------------
    def _open_axes_dialog(self):
        try:
            from .axis_ticks_dialog import AxisTicksDialog
        except Exception:
            QMessageBox.information(self, "Axis", "Axis dialog not available.")
            return
        cur = self._props.get('axes') or {}
        dlg = AxisTicksDialog(self, initial=cur)
        from PySide6.QtWidgets import QDialog
        if dlg.exec() != QDialog.Accepted:
            return
        vals = dlg.get_values()
        self._props['axes'] = vals
        self._apply_saved_props()
        self.finish()

    # ---------------------- properties ----------------------
    def _open_properties(self):
        try:
            from .plot_props_dialog import PlotPropertiesDialog
        except Exception:
            QMessageBox.warning(self, "Properties", "Properties dialog not available.")
            return
        # Build current series list from mapping
        series = []  # list of (yj, label, color, linewidth)
        for yj, line in self._y_lines.items():
            try:
                color = line.get_color()
                lw = float(line.get_linewidth())
            except Exception:
                color = None; lw = None
            series.append((yj, self._series_label(yj), color, lw))

        dlg = PlotPropertiesDialog(self, mode=self._mode, has_right=(self._ax_right is not None), series=series)

        # Prefill from saved props
        props = self._props.copy()
        dlg.set_initial(props)
        from PySide6.QtWidgets import QDialog
        if dlg.exec() != QDialog.Accepted:
            return
        new_props = dlg.get_values()
        self.apply_properties(new_props)

    def apply_properties(self, props: Dict[str, object]):
        # Store and apply
        self._props.update(props or {})
        self._apply_saved_props()
        self.finish()

    def _apply_saved_props(self):
        p = self._props or {}
        # Labels
        if 'xlabel' in p:
            self.ax.set_xlabel(str(p.get('xlabel') or ''))
        if 'ylabel' in p:
            self.ax.set_ylabel(str(p.get('ylabel') or ''))
        if self._ax_right is not None and 'y2label' in p:
            try:
                self._ax_right.set_ylabel(str(p.get('y2label') or ''))
            except Exception:
                pass
        if self._mode == 'surface3d' and 'zlabel' in p:
            try:
                self.ax.set_zlabel(str(p.get('zlabel') or ''))
            except Exception:
                pass
        # Background colors
        if p.get('fig_face'):
            try:
                self.fig.set_facecolor(p['fig_face'])
            except Exception:
                pass
        if p.get('ax_face'):
            try:
                self.ax.set_facecolor(p['ax_face'])
            except Exception:
                pass
        # Grid toggle
        try:
            self.ax.grid(bool(p.get('grid', True)), color="#cccccc", linewidth=0.6, alpha=1.0)
        except Exception:
            pass
        # Fonts
        ff = p.get('font_family')
        fs = p.get('font_size')
        try:
            if ff or fs:
                lab = self.ax.xaxis.label; lab.set_fontfamily(ff or lab.get_fontfamily());
                if fs: lab.set_fontsize(fs)
                lab = self.ax.yaxis.label; lab.set_fontfamily(ff or lab.get_fontfamily());
                if fs: lab.set_fontsize(fs)
                for lbl in self.ax.get_xticklabels() + self.ax.get_yticklabels():
                    if ff: lbl.set_fontfamily(ff)
                    if fs: lbl.set_fontsize(fs)
                if self._ax_right is not None:
                    lab = self._ax_right.yaxis.label
                    if ff: lab.set_fontfamily(ff)
                    if fs: lab.set_fontsize(fs)
                    for lbl in self._ax_right.get_yticklabels():
                        if ff: lbl.set_fontfamily(ff)
                        if fs: lbl.set_fontsize(fs)
        except Exception:
            pass
        # Top/right spines and ticks
        try:
            show_top = bool(p.get('show_top', False))
            show_right = bool(p.get('show_right', False))
            self.ax.spines['top'].set_visible(show_top)
            self.ax.tick_params(top=show_top)
            # Right spine: if double_y mode, right axis exists
            if self._ax_right is not None:
                self._ax_right.spines['right'].set_visible(show_right)
                self._ax_right.tick_params(right=show_right)
            else:
                self.ax.spines['right'].set_visible(show_right)
                self.ax.tick_params(right=show_right)
        except Exception:
            pass
        # Axis/Ticks advanced options
        axes_cfg = p.get('axes') or {}
        ticks = axes_cfg.get('ticks') or {}
        try:
            self.ax.tick_params(direction=ticks.get('direction', 'out'), length=float(ticks.get('length', 3.5)),
                                width=float(ticks.get('width', 0.8)))
            for lbl in self.ax.get_xticklabels():
                try: lbl.set_rotation(float(ticks.get('xrot', 0)))
                except Exception: pass
            for lbl in self.ax.get_yticklabels():
                try: lbl.set_rotation(float(ticks.get('yrot', 0)))
                except Exception: pass
            if bool(ticks.get('minor', False)):
                self.ax.minorticks_on()
            else:
                self.ax.minorticks_off()
        except Exception:
            pass
        # Top axis
        top_cfg = axes_cfg.get('top') or {}
        try:
            mode = str(top_cfg.get('mode', 'Off'))
            if mode == 'Mirror X':
                self.ax.tick_params(top=True, labeltop=True)
                self.ax.spines['top'].set_visible(True)
            elif mode == 'Independent':
                if self._ax_top is None:
                    self._ax_top = self.ax.twiny()
                self._ax_top.spines['top'].set_visible(True)
                self._ax_top.set_xlabel(str(top_cfg.get('label', '')))
                xlim = top_cfg.get('xlim') or [None, None]
                try:
                    lo = xlim[0]; hi = xlim[1]
                    if lo is not None and hi is not None and hi > lo:
                        self._ax_top.set_xlim(lo, hi)
                except Exception:
                    pass
        except Exception:
            pass
        # Right axis advanced (independent)
        right_cfg = axes_cfg.get('right') or {}
        try:
            mode = str(right_cfg.get('mode', 'Off'))
            if mode == 'Mirror Y':
                self.ax.tick_params(right=True, labelright=True)
                self.ax.spines['right'].set_visible(True)
            elif mode == 'Independent':
                if self._ax_right is None:
                    self._ax_right = self.ax.twinx()
                self._ax_right.spines['right'].set_visible(True)
                self._ax_right.set_ylabel(str(right_cfg.get('label', '')))
                ylim = right_cfg.get('ylim') or [None, None]
                try:
                    lo = ylim[0]; hi = ylim[1]
                    if lo is not None and hi is not None and hi > lo:
                        self._ax_right.set_ylim(lo, hi)
                except Exception:
                    pass
        except Exception:
            pass
        # Limits
        def _apply_lim(ax, key):
            lim = p.get(key)
            if isinstance(lim, (list, tuple)) and len(lim) == 2:
                try:
                    lo = None if lim[0] is None or lim[0] == '' else float(lim[0])
                    hi = None if lim[1] is None or lim[1] == '' else float(lim[1])
                    if lo is not None and hi is not None and hi > lo:
                        ax.set_xlim(lo, hi) if key.startswith('x') else ax.set_ylim(lo, hi)
                    elif lo is None and hi is None:
                        pass  # leave autoscale
                except Exception:
                    pass
        _apply_lim(self.ax, 'xlim')
        _apply_lim(self.ax, 'ylim')
        if self._ax_right is not None:
            lim = p.get('y2lim')
            if isinstance(lim, (list, tuple)) and len(lim) == 2:
                try:
                    lo = None if lim[0] is None or lim[0] == '' else float(lim[0])
                    hi = None if lim[1] is None or lim[1] == '' else float(lim[1])
                    if lo is not None and hi is not None and hi > lo:
                        self._ax_right.set_ylim(lo, hi)
                except Exception:
                    pass
        if self._mode == 'surface3d':
            lim = p.get('zlim')
            if isinstance(lim, (list, tuple)) and len(lim) == 2:
                try:
                    lo = None if lim[0] is None or lim[0] == '' else float(lim[0])
                    hi = None if lim[1] is None or lim[1] == '' else float(lim[1])
                    if lo is not None and hi is not None and hi > lo:
                        self.ax.set_zlim(lo, hi)
                except Exception:
                    pass
        # Line styles: map by y-index
        styles = p.get('line_styles') or {}
        if isinstance(styles, dict):
            for yj, line in list(self._y_lines.items()):
                s = styles.get(int(yj)) or styles.get(str(yj))
                if not s:
                    continue
                try:
                    if 'color' in s and s['color']:
                        line.set_color(str(s['color']))
                    if 'linewidth' in s and s['linewidth']:
                        line.set_linewidth(float(s['linewidth']))
                except Exception:
                    pass
        # Legend after other props
        self._apply_legend_opts()

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
