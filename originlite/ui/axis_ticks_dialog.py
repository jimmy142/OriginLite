from __future__ import annotations

from typing import Dict, Any

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QDialogButtonBox, QComboBox, QLineEdit,
    QHBoxLayout, QLabel, QDoubleSpinBox, QCheckBox
)


class AxisTicksDialog(QDialog):
    def __init__(self, parent=None, *, initial: Dict[str, Any] | None = None):
        super().__init__(parent)
        self.setWindowTitle("Axis && Ticks")
        lay = QVBoxLayout(self)
        form = QFormLayout()

        # Top axis mode: off/mirror/independent
        self.top_mode = QComboBox(); self.top_mode.addItems(["Off", "Mirror X", "Independent"])
        self.top_label = QLineEdit(); self.top_xlim_lo = QLineEdit(); self.top_xlim_hi = QLineEdit()
        top_row = QHBoxLayout();
        top_row.addWidget(QLabel("Label")); top_row.addWidget(self.top_label)
        top_row.addWidget(QLabel("Min")); top_row.addWidget(self.top_xlim_lo)
        top_row.addWidget(QLabel("Max")); top_row.addWidget(self.top_xlim_hi)
        form.addRow("Top axis", self.top_mode)
        form.addRow("Top settings", top_row)

        # Right axis mode
        self.right_mode = QComboBox(); self.right_mode.addItems(["Off", "Mirror Y", "Independent"])
        self.right_label = QLineEdit(); self.right_ylim_lo = QLineEdit(); self.right_ylim_hi = QLineEdit()
        right_row = QHBoxLayout();
        right_row.addWidget(QLabel("Label")); right_row.addWidget(self.right_label)
        right_row.addWidget(QLabel("Min")); right_row.addWidget(self.right_ylim_lo)
        right_row.addWidget(QLabel("Max")); right_row.addWidget(self.right_ylim_hi)
        form.addRow("Right axis", self.right_mode)
        form.addRow("Right settings", right_row)

        # Ticks
        self.tick_dir = QComboBox(); self.tick_dir.addItems(["out", "in"])
        self.tick_len = QDoubleSpinBox(); self.tick_len.setRange(1.0, 20.0); self.tick_len.setValue(3.5)
        self.tick_w = QDoubleSpinBox(); self.tick_w.setRange(0.1, 5.0); self.tick_w.setValue(0.8)
        self.x_rot = QDoubleSpinBox(); self.x_rot.setRange(-90, 90); self.x_rot.setValue(0)
        self.y_rot = QDoubleSpinBox(); self.y_rot.setRange(-90, 90); self.y_rot.setValue(0)
        self.minor_chk = QCheckBox("Show minor ticks")
        ticks_row = QHBoxLayout();
        ticks_row.addWidget(QLabel("Direction")); ticks_row.addWidget(self.tick_dir)
        ticks_row.addWidget(QLabel("Length")); ticks_row.addWidget(self.tick_len)
        ticks_row.addWidget(QLabel("Width")); ticks_row.addWidget(self.tick_w)
        ticks_row.addWidget(QLabel("X rot")); ticks_row.addWidget(self.x_rot)
        ticks_row.addWidget(QLabel("Y rot")); ticks_row.addWidget(self.y_rot)
        form.addRow("Ticks", ticks_row)
        form.addRow(self.minor_chk)

        lay.addLayout(form)
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept); btns.rejected.connect(self.reject)
        lay.addWidget(btns)

        if initial:
            self._load(initial)

    def _load(self, d: Dict[str, Any]):
        top = d.get('top', {})
        mode = str(top.get('mode', 'Off'))
        self.top_mode.setCurrentText(mode)
        self.top_label.setText(str(top.get('label', '')))
        xl = top.get('xlim') or [None, None]
        self.top_xlim_lo.setText('' if xl[0] is None else str(xl[0]))
        self.top_xlim_hi.setText('' if xl[1] is None else str(xl[1]))

        right = d.get('right', {})
        mode = str(right.get('mode', 'Off'))
        self.right_mode.setCurrentText(mode)
        self.right_label.setText(str(right.get('label', '')))
        yl = right.get('ylim') or [None, None]
        self.right_ylim_lo.setText('' if yl[0] is None else str(yl[0]))
        self.right_ylim_hi.setText('' if yl[1] is None else str(yl[1]))

        ticks = d.get('ticks', {})
        self.tick_dir.setCurrentText(str(ticks.get('direction', 'out')))
        try:
            if 'length' in ticks: self.tick_len.setValue(float(ticks['length']))
            if 'width' in ticks: self.tick_w.setValue(float(ticks['width']))
            if 'xrot' in ticks: self.x_rot.setValue(float(ticks['xrot']))
            if 'yrot' in ticks: self.y_rot.setValue(float(ticks['yrot']))
        except Exception:
            pass
        self.minor_chk.setChecked(bool(ticks.get('minor', False)))

    def get_values(self) -> Dict[str, Any]:
        def _pair(a: QLineEdit, b: QLineEdit):
            lo = a.text().strip(); hi = b.text().strip()
            lo_v = float(lo) if lo not in ('', None) else None
            hi_v = float(hi) if hi not in ('', None) else None
            return [lo_v, hi_v]
        return {
            'top': {
                'mode': self.top_mode.currentText(),
                'label': self.top_label.text().strip(),
                'xlim': _pair(self.top_xlim_lo, self.top_xlim_hi),
            },
            'right': {
                'mode': self.right_mode.currentText(),
                'label': self.right_label.text().strip(),
                'ylim': _pair(self.right_ylim_lo, self.right_ylim_hi),
            },
            'ticks': {
                'direction': self.tick_dir.currentText(),
                'length': float(self.tick_len.value()),
                'width': float(self.tick_w.value()),
                'xrot': float(self.x_rot.value()),
                'yrot': float(self.y_rot.value()),
                'minor': self.minor_chk.isChecked(),
            }
        }

