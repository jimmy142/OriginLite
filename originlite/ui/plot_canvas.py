from PySide6.QtCore import Qt
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

class PlotCanvas(FigureCanvas):
    def __init__(self, parent=None):
        self.fig = Figure(figsize=(6, 4), layout='constrained')
        super().__init__(self.fig)
        self.ax = self.fig.add_subplot(111)
        self.ax.grid(True, which='both', alpha=0.25)
        self.setFocusPolicy(Qt.ClickFocus)
        self.setFocus()

    def clear_axes(self):
        self.ax.clear()
        self.ax.grid(True, which='both', alpha=0.25)
        self.draw_idle()
