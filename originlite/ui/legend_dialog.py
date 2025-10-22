from __future__ import annotations

from typing import Dict, List, Tuple, Any

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout, QGroupBox, QLabel, QLineEdit,
    QDialogButtonBox, QComboBox, QCheckBox, QSpinBox, QDoubleSpinBox, QPushButton, QScrollArea, QWidget,
    QFontComboBox
)


def _set_btn_color(btn: QPushButton, color: str | None):
    if color:
        btn.setProperty('chosen_color', color)
        btn.setStyleSheet(f"background-color: {color};")
    else:
        btn.setProperty('chosen_color', None)
        btn.setStyleSheet("")


class LegendDialog(QDialog):
    def __init__(self, parent=None, *, series: List[Tuple[str, str]], opts: Dict[str, Any] | None = None):
        """
        series: list of (display_name, current_label)
        opts: existing legend options
        """
        super().__init__(parent)
        self.setWindowTitle("Legend Properties")
        self._labels_edits: List[QLineEdit] = []

        lay = QVBoxLayout(self)

        # Options group
        g_opts = QGroupBox("Options")
        f = QFormLayout()
        self.show_chk = QCheckBox("Show legend")
        self.loc_combo = QComboBox()
        locs = [
            'best','upper right','upper left','lower left','lower right','right',
            'center left','center right','lower center','upper center','center'
        ]
        for s in locs:
            self.loc_combo.addItem(s)
        self.ncol = QSpinBox(); self.ncol.setRange(1, 10); self.ncol.setValue(1)
        self.font_box = QFontComboBox(); self.font_box.setEditable(False)
        self.font_size = QDoubleSpinBox(); self.font_size.setRange(5.0, 48.0); self.font_size.setSingleStep(0.5); self.font_size.setValue(10.0)
        self.frame_chk = QCheckBox("Show frame (outline)")
        self.face_btn = QPushButton("Face color…")
        self.edge_btn = QPushButton("Edge color…")
        self.edge_w = QDoubleSpinBox(); self.edge_w.setRange(0.1, 10.0); self.edge_w.setSingleStep(0.1); self.edge_w.setValue(1.0)
        self.face_btn.clicked.connect(lambda: self._pick_color(self.face_btn))
        self.edge_btn.clicked.connect(lambda: self._pick_color(self.edge_btn))

        f.addRow(self.show_chk)
        f.addRow("Location", self.loc_combo)
        f.addRow("Columns", self.ncol)
        f.addRow("Font", self.font_box)
        f.addRow("Font size", self.font_size)
        f.addRow(self.frame_chk)
        f.addRow("Frame face", self.face_btn)
        f.addRow("Frame edge", self.edge_btn)
        f.addRow("Edge width", self.edge_w)
        g_opts.setLayout(f)
        lay.addWidget(g_opts)

        # Labels area (scrollable if many)
        lay.addWidget(QLabel("Series Labels"))
        labels_form = QFormLayout()
        for display, current in series:
            ed = QLineEdit(current)
            self._labels_edits.append(ed)
            labels_form.addRow(QLabel(display), ed)
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        inner = QWidget(); inner.setLayout(labels_form)
        scroll.setWidget(inner)
        lay.addWidget(scroll)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)

        # Prefill opts
        opts = opts or {}
        self.show_chk.setChecked(bool(opts.get('visible', True)))
        loc = str(opts.get('loc', 'best'))
        idx = max(0, self.loc_combo.findText(loc))
        self.loc_combo.setCurrentIndex(idx)
        self.ncol.setValue(int(opts.get('ncol', 1)))
        if opts.get('fontfamily'):
            try:
                self.font_box.setCurrentFont(QFont(str(opts.get('fontfamily'))))
            except Exception:
                pass
        if opts.get('fontsize'):
            try:
                self.font_size.setValue(float(opts.get('fontsize')))
            except Exception:
                pass
        self.frame_chk.setChecked(bool(opts.get('frameon', False)))
        _set_btn_color(self.face_btn, opts.get('facecolor'))
        _set_btn_color(self.edge_btn, opts.get('edgecolor'))
        if opts.get('edgewidth'):
            try:
                self.edge_w.setValue(float(opts.get('edgewidth')))
            except Exception:
                pass

    def _pick_color(self, btn: QPushButton):
        from PySide6.QtWidgets import QColorDialog
        c = QColorDialog.getColor(parent=self)
        if c.isValid():
            _set_btn_color(btn, c.name())

    def get_values(self) -> Dict[str, Any]:
        return {
            'visible': self.show_chk.isChecked(),
            'loc': self.loc_combo.currentText(),
            'ncol': int(self.ncol.value()),
            'fontfamily': self.font_box.currentFont().family(),
            'fontsize': float(self.font_size.value()),
            'frameon': self.frame_chk.isChecked(),
            'facecolor': self.face_btn.property('chosen_color'),
            'edgecolor': self.edge_btn.property('chosen_color'),
            'edgewidth': float(self.edge_w.value()),
            'labels': [ed.text().strip() for ed in self._labels_edits],
        }
