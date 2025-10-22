from PySide6.QtWidgets import (
    QWidget, QFormLayout, QComboBox, QLineEdit, QCheckBox, QPushButton, QHBoxLayout
)

class ControlPanel(QWidget):
    def __init__(self):
        super().__init__()
        self.x_combo = QComboBox()
        self.y_combo = QComboBox()
        self.label_edit = QLineEdit()
        self.label_edit.setPlaceholderText("Trace label (optional)")
        self.marker_chk = QCheckBox("Markers")
        self.line_chk = QCheckBox("Lines")
        self.line_chk.setChecked(True)
        self.add_btn = QPushButton("Add Trace")

        form = QFormLayout()
        form.addRow("X column", self.x_combo)
        form.addRow("Y column", self.y_combo)
        form.addRow("Label", self.label_edit)
        style_row = QHBoxLayout()
        style_row.addWidget(self.line_chk)
        style_row.addWidget(self.marker_chk)
        form.addRow("Style", style_row)
        form.addRow(self.add_btn)
        self.setLayout(form)

    def set_headers(self, headers):
        self.x_combo.clear(); self.y_combo.clear()
        self.x_combo.addItems(headers)
        self.y_combo.addItems(headers)
