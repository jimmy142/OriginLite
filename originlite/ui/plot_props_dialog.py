from __future__ import annotations

from typing import Dict, List, Tuple

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout, QGroupBox, QLineEdit,
    QDialogButtonBox, QLabel, QDoubleSpinBox, QPushButton, QComboBox, QCheckBox, QFontComboBox
)
from PySide6.QtGui import QFont


def _qcolor_to_hex(c: QColor) -> str:
    return c.name()  # #RRGGBB


class PlotPropertiesDialog(QDialog):
    def __init__(self, parent=None, *, mode: str, has_right: bool, series: List[Tuple[int, str, object, float]]):
        super().__init__(parent)
        self.setWindowTitle("Graph Properties")
        self._mode = mode
        self._has_right = bool(has_right)
        # series: list of (yj, label, color, linewidth)
        self._series = series

        self._values: Dict[str, object] = {}

        lay = QVBoxLayout(self)

        # Axes group
        axes_group = QGroupBox("Axes")
        axes_form = QFormLayout()
        self.x_label = QLineEdit(); self.y_label = QLineEdit(); self.y2_label = QLineEdit(); self.z_label = QLineEdit()
        self.xmin = QLineEdit(); self.xmax = QLineEdit(); self.ymin = QLineEdit(); self.ymax = QLineEdit()
        self.y2min = QLineEdit(); self.y2max = QLineEdit(); self.zmin = QLineEdit(); self.zmax = QLineEdit()
        axes_form.addRow("X label", self.x_label)
        axes_form.addRow("Y label", self.y_label)
        if self._has_right:
            axes_form.addRow("Y2 label", self.y2_label)
        if self._mode == 'surface3d':
            axes_form.addRow("Z label", self.z_label)
        xlim_row = QHBoxLayout(); xlim_row.addWidget(QLabel("Min")); xlim_row.addWidget(self.xmin); xlim_row.addWidget(QLabel("Max")); xlim_row.addWidget(self.xmax)
        ylim_row = QHBoxLayout(); ylim_row.addWidget(QLabel("Min")); ylim_row.addWidget(self.ymin); ylim_row.addWidget(QLabel("Max")); ylim_row.addWidget(self.ymax)
        axes_form.addRow("X limits", xlim_row)
        axes_form.addRow("Y limits", ylim_row)
        if self._has_right:
            y2_row = QHBoxLayout(); y2_row.addWidget(QLabel("Min")); y2_row.addWidget(self.y2min); y2_row.addWidget(QLabel("Max")); y2_row.addWidget(self.y2max)
            axes_form.addRow("Y2 limits", y2_row)
        if self._mode == 'surface3d':
            z_row = QHBoxLayout(); z_row.addWidget(QLabel("Min")); z_row.addWidget(self.zmin); z_row.addWidget(QLabel("Max")); z_row.addWidget(self.zmax)
            axes_form.addRow("Z limits", z_row)
        axes_group.setLayout(axes_form)
        lay.addWidget(axes_group)

        # Appearance group
        app_group = QGroupBox("Appearance")
        app_form = QFormLayout()
        self.bg_fig_btn = QPushButton("Figure color…")
        self.bg_ax_btn = QPushButton("Axes color…")
        self.grid_chk = QCheckBox("Show grid")
        self.show_top_chk = QCheckBox("Show top ticks/spine")
        self.show_right_chk = QCheckBox("Show right ticks/spine")
        self.font_family = QFontComboBox(); self.font_family.setEditable(False)
        self.font_size = QDoubleSpinBox(); self.font_size.setRange(5.0, 48.0); self.font_size.setSingleStep(0.5); self.font_size.setValue(10.0)
        self.bg_fig_btn.clicked.connect(lambda: self._choose_color_button(self.bg_fig_btn))
        self.bg_ax_btn.clicked.connect(lambda: self._choose_color_button(self.bg_ax_btn))
        app_form.addRow("Figure background", self.bg_fig_btn)
        app_form.addRow("Axes background", self.bg_ax_btn)
        app_form.addRow("Font family", self.font_family)
        app_form.addRow("Font size", self.font_size)
        app_form.addRow(self.grid_chk)
        app_form.addRow(self.show_top_chk)
        app_form.addRow(self.show_right_chk)
        app_group.setLayout(app_form)
        lay.addWidget(app_group)

        # Series group (only meaningful for line-like modes)
        self.series_group = QGroupBox("Series Style")
        sform = QFormLayout()
        self.series_combo = QComboBox()
        for yj, label, color, lw in self._series:
            self.series_combo.addItem(f"{label} (Y{yj+1})", userData=yj)
        self.color_btn = QPushButton("Pick color")
        self._current_color = None  # hex string
        self.lw_spin = QDoubleSpinBox(); self.lw_spin.setRange(0.1, 20.0); self.lw_spin.setSingleStep(0.1); self.lw_spin.setValue(1.5)
        self.color_btn.clicked.connect(self._choose_color)
        self.series_combo.currentIndexChanged.connect(self._series_changed)
        sform.addRow("Series", self.series_combo)
        srow = QHBoxLayout(); srow.addWidget(QLabel("Line width")); srow.addWidget(self.lw_spin); srow.addSpacing(12); srow.addWidget(self.color_btn); srow.addStretch(1)
        sform.addRow("Style", srow)
        self.series_group.setLayout(sform)
        visible_series = (mode in ("line", "line_markers", "double_y")) and len(self._series) > 0
        self.series_group.setVisible(visible_series)
        lay.addWidget(self.series_group)

        # Buttons
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel | QDialogButtonBox.Apply)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        btns.button(QDialogButtonBox.Apply).clicked.connect(self._apply)
        lay.addWidget(btns)

        # Internal per-series styles
        self._line_styles: Dict[int, Dict[str, object]] = {}
        # Prefill current series styles
        for yj, label, color, lw in self._series:
            st = {}
            if color:
                st['color'] = color if isinstance(color, str) else None
            if lw:
                st['linewidth'] = lw
            if st:
                self._line_styles[int(yj)] = st
        if visible_series:
            self._series_changed(0)

    def set_initial(self, props: Dict[str, object]):
        # Labels
        self.x_label.setText(str(props.get('xlabel') or ''))
        self.y_label.setText(str(props.get('ylabel') or ''))
        if self._has_right:
            self.y2_label.setText(str(props.get('y2label') or ''))
        if self._mode == 'surface3d':
            self.z_label.setText(str(props.get('zlabel') or ''))
        # Limits
        def _fill(ed_lo, ed_hi, key):
            v = props.get(key)
            if isinstance(v, (list, tuple)) and len(v) == 2:
                ed_lo.setText('' if v[0] is None else str(v[0]))
                ed_hi.setText('' if v[1] is None else str(v[1]))
        _fill(self.xmin, self.xmax, 'xlim')
        _fill(self.ymin, self.ymax, 'ylim')
        if self._has_right:
            _fill(self.y2min, self.y2max, 'y2lim')
        if self._mode == 'surface3d':
            _fill(self.zmin, self.zmax, 'zlim')
        # Appearance
        self._set_btn_color(self.bg_fig_btn, props.get('fig_face'))
        self._set_btn_color(self.bg_ax_btn, props.get('ax_face'))
        fam = str(props.get('font_family') or '')
        if fam:
            try:
                self.font_family.setCurrentFont(QFont(fam))
            except Exception:
                pass
        try:
            if props.get('font_size'):
                self.font_size.setValue(float(props.get('font_size')))
        except Exception:
            pass
        self.grid_chk.setChecked(bool(props.get('grid', True)))
        self.show_top_chk.setChecked(bool(props.get('show_top', False)))
        self.show_right_chk.setChecked(bool(props.get('show_right', False)))

    def _choose_color(self):
        from PySide6.QtWidgets import QColorDialog
        c = QColorDialog.getColor(parent=self)
        if c.isValid():
            self._current_color = _qcolor_to_hex(c)
            self._save_current_series_style()

    def _choose_color_button(self, btn: QPushButton):
        from PySide6.QtWidgets import QColorDialog
        c = QColorDialog.getColor(parent=self)
        if c.isValid():
            self._set_btn_color(btn, _qcolor_to_hex(c))

    def _set_btn_color(self, btn: QPushButton, color: str | None):
        if color:
            btn.setProperty('chosen_color', color)
            btn.setStyleSheet(f"background-color: {color};")
        else:
            btn.setProperty('chosen_color', None)
            btn.setStyleSheet("")

    def _series_changed(self, idx: int):
        if idx < 0:
            return
        yj = self.series_combo.currentData()
        st = self._line_styles.get(int(yj), {})
        self._current_color = st.get('color')
        lw = st.get('linewidth')
        if lw:
            try:
                self.lw_spin.setValue(float(lw))
            except Exception:
                pass

    def _save_current_series_style(self):
        if self.series_combo.count() == 0:
            return
        yj = int(self.series_combo.currentData())
        st = self._line_styles.get(yj, {})
        st['linewidth'] = float(self.lw_spin.value())
        if self._current_color:
            st['color'] = self._current_color
        self._line_styles[yj] = st

    def _collect(self) -> Dict[str, object]:
        self._save_current_series_style()
        out: Dict[str, object] = {}
        # Labels
        out['xlabel'] = self.x_label.text().strip()
        out['ylabel'] = self.y_label.text().strip()
        if self._has_right:
            out['y2label'] = self.y2_label.text().strip()
        if self._mode == 'surface3d':
            out['zlabel'] = self.z_label.text().strip()
        # Limits
        def _pair(a: QLineEdit, b: QLineEdit):
            lo = a.text().strip(); hi = b.text().strip()
            lo_v = float(lo) if lo not in ('', None) else None
            hi_v = float(hi) if hi not in ('', None) else None
            return [lo_v, hi_v]
        out['xlim'] = _pair(self.xmin, self.xmax)
        out['ylim'] = _pair(self.ymin, self.ymax)
        if self._has_right:
            out['y2lim'] = _pair(self.y2min, self.y2max)
        if self._mode == 'surface3d':
            out['zlim'] = _pair(self.zmin, self.zmax)
        # Series styles
        if self._line_styles:
            out['line_styles'] = self._line_styles
        # Appearance
        out['fig_face'] = self.bg_fig_btn.property('chosen_color')
        out['ax_face'] = self.bg_ax_btn.property('chosen_color')
        out['font_family'] = self.font_family.currentFont().family()
        out['font_size'] = float(self.font_size.value())
        out['grid'] = self.grid_chk.isChecked()
        out['show_top'] = self.show_top_chk.isChecked()
        out['show_right'] = self.show_right_chk.isChecked()
        return out

    def _apply(self):
        vals = self._collect()
        self._values.update(vals)
        # live-apply via parent if available
        if hasattr(self.parent(), 'apply_properties'):
            try:
                self.parent().apply_properties(vals)
            except Exception:
                pass

    def accept(self):
        self._apply()
        super().accept()

    def get_values(self) -> Dict[str, object]:
        return self._values or self._collect()
