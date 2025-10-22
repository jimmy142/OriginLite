from __future__ import annotations

from typing import Optional, Dict

from PySide6.QtWidgets import QDialog, QVBoxLayout, QDialogButtonBox, QLabel, QDoubleSpinBox, QPushButton


class LineStyleDialog(QDialog):
    def __init__(self, parent=None, *, initial_color: Optional[str] = None, initial_width: Optional[float] = None):
        super().__init__(parent)
        self.setWindowTitle("Line Style")
        lay = QVBoxLayout(self)
        self.lw = QDoubleSpinBox(); self.lw.setRange(0.1, 20.0); self.lw.setSingleStep(0.1)
        if initial_width:
            try:
                self.lw.setValue(float(initial_width))
            except Exception:
                pass
        self.color_btn = QPushButton("Color…")
        if initial_color:
            self.color_btn.setStyleSheet(f"background-color: {initial_color};")
            self.color_btn.setProperty('chosen_color', initial_color)
        self.color_btn.clicked.connect(self._pick_color)
        lay.addWidget(QLabel("Line width"))
        lay.addWidget(self.lw)
        lay.addWidget(self.color_btn)
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)

    def _pick_color(self):
        from PySide6.QtWidgets import QColorDialog
        c = QColorDialog.getColor(parent=self)
        if c.isValid():
            self.color_btn.setProperty('chosen_color', c.name())
            self.color_btn.setStyleSheet(f"background-color: {c.name()};")

    def get_values(self) -> Dict[str, object]:
        out = {
            'linewidth': float(self.lw.value()),
            'color': self.color_btn.property('chosen_color')
        }
        return out

