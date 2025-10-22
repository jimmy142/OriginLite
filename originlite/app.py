# originlite/app.py
from __future__ import annotations
import sys
from PySide6.QtWidgets import QApplication
from .ui.workspace import Workspace

def run():
    app = QApplication.instance() or QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(True)

    win = Workspace()
    # Ensure we clean up all windows/figures/signals on exit
    app.aboutToQuit.connect(win.shutdown)

    win.show()
    sys.exit(app.exec())
