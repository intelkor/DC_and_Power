import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QLabel, QPushButton, QFileDialog, QWidget, QAction, QToolBar, QStyle
from PyQt5.QtGui import QPixmap, QColor
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import pandas as pd
import matplotlib.pyplot as plt
from scipy.interpolate import interp1d, pchip, UnivariateSpline

class ExcelPlotterApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # Make the main window approximately 1.5 times bigger
        self.setGeometry(100, 100, int(1.5 * 800), int(1.5 * 600))

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()

        self.file_label = QLabel("Selected File:")
        self.layout.addWidget(self.file_label)

        # Add QLabel to display selected file with white background
        self.selected_file_display = QLabel(self)
        self.selected_file_display.setFrameShape(QLabel.Box)
        self.selected_file_display.setFrameShadow(QLabel.Sunken)
        self.selected_file_display.setStyleSheet("background-color: white")
        self.layout.addWidget(self.selected_file_display)

        # Browse button with icon
        self.browse_button = QPushButton("Browse")
        self.browse_button.setIcon(self.style().standardIcon(QStyle.SP_DirIcon))
        self.browse_button.clicked.connect(self.browse_file)
        self.layout.addWidget(self.browse_button)

        # Plot button with icon
        self.plot_button = QPushButton("Plot")
        self.plot_button.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))
        self.plot_button.clicked.connect(self.plot_data)
        self.layout.addWidget(self.plot_button)

        # Save button with icon
        self.save_button = QPushButton("Save")
        self.save_button.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        self.save_button.clicked.connect(self.save_plot)
        self.layout.addWidget(self.save_button)

        self.figure, self.axes = plt.subplots(2, 1, figsize=(5, 5))
        self.canvas = FigureCanvas(self.figure)
        self.layout.addWidget(self.canvas)

        self.central_widget.setLayout(self.layout)

        self.file_path = None

        # Add Save action to the toolbar
        toolbar = QToolBar("Toolbar")
        self.addToolBar(toolbar)

        save_action = QAction("Save", self)
        save_action.triggered.connect(self.save_plot)
        save_action.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        toolbar.addAction(save_action)

    def browse_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_dialog = QFileDialog()
        file_dialog.setOptions(options)
        file_path, _ = file_dialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx);;All Files (*)")

        if file_path:
            self.file_path = file_path
            self.file_label.setText(f"Selected File: {file_path}")

            # Display selected file with white background in QLabel
            self.selected_file_display.setText(file_path)
            self.selected_file_display.adjustSize()

    def check_dependencies(self):
        try:
            import openpyxl
            import scipy
            return True
        except ImportError:
            return False

    def plot_data(self):
        if self.file_path:
            if self.check_dependencies():
                try:
                    data = pd.read_excel(self.file_path)
                    selion = data.iloc[:, 0]
                    mP = data.iloc[:, 1]

                    DCm = data.iloc[:, [0, 1]]
                    xq = range(64)

                    DC = interp1d(DCm.iloc[:, 0], DCm.iloc[:, 1], kind='linear', fill_value='extrapolate')(xq)
                    DC1 = pchip(DCm.iloc[:, 0], DCm.iloc[:, 1])(xq)

                    spline_degree = 2
                    DC2 = UnivariateSpline(DCm.iloc[:, 0], DCm.iloc[:, 1], k=spline_degree)(xq)

                    P = mP / DC1    # Calculation of Peak Power
                    
                    self.axes[0].plot(xq, DC, 'b+', label='Interpolated DC')
                    self.axes[0].plot(xq, DC1, 'r*', label='PCHIP')
                    self.axes[0].plot(xq, DC2, 'gx', label=f'Spline (degree {spline_degree})')
                    self.axes[0].plot(DCm.iloc[:, 0], DCm.iloc[:, 1], '.', label='Measured DC')
                    self.axes[0].set_ylabel('Duty Cycle')
                    self.axes[0].set_title('Selion vs. Duty Cycle')
                    self.axes[0].grid(True, which='both', linestyle='--', linewidth=0.5)
                    self.axes[0].legend(loc='lower right')  # Reposition the legend

                    self.axes[1].plot(selion, P, label='Interpolated Ppk')
                    self.axes[1].plot(selion, DCm.iloc[:, 6], label='Measured Ppk')
                    self.axes[1].set_ylabel('Peak Power')
                    self.axes[1].set_title('Selion vs. Peak Power')
                    self.axes[1].grid(True, which='both', linestyle='--', linewidth=0.5)
                    self.axes[1].legend(loc='lower right')  # Reposition the legend

                    self.axes[0].set_xlabel('Selion')
                    self.axes[1].set_xlabel('Selion')

                    # Add minor grids
                    self.axes[0].minorticks_on()
                    self.axes[0].grid(which='minor', linestyle=':', linewidth='0.5', color='black')

                    self.axes[1].minorticks_on()
                    self.axes[1].grid(which='minor', linestyle=':', linewidth='0.5', color='black')

                    self.figure.tight_layout()
                    self.canvas.draw()
                except Exception as e:
                    print(f"Error plotting data: {e}")
            else:
                print("Please install the 'openpyxl' and 'scipy' libraries.")
        else:
            print("Please select a file first.")

    def save_plot(self):
        if self.file_path:
            try:
                save_path, _ = QFileDialog.getSaveFileName(self, "Save Plot", "", "PNG Files (*.png);;All Files (*)")
                if save_path:
                    self.figure.savefig(save_path)
                    print(f"Plot saved to: {save_path}")
            except Exception as e:
                print(f"Error saving plot: {e}")
        else:
            print("Please plot data first.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelPlotterApp()
    window.show()
    sys.exit(app.exec_())
