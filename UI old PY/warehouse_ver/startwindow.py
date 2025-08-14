from PyQt5 import QtWidgets
from PyQt5.QtGui import QIcon
from startwindow_ui import Ui_StartWindow  # Adjust if your generated class name is different
from ColourMasterBox_collection import ColorSearchApp  # Adjust class name if needed
from warehouse_form import WarehouseForm  # Adjust class name if needed
import sys
import resources_rc

class StartWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_StartWindow()
        self.ui.setupUi(self)
        self.ui.QCappButton.clicked.connect(self.open_qc_app)
        self.ui.WHappButton.clicked.connect(self.open_wh_app)

    def open_qc_app(self):
        self.qc_window = ColorSearchApp()
        self.qc_window.show()

    def open_wh_app(self):
        self.wh_window = WarehouseForm()
        self.wh_window.show()

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    app.setWindowIcon(QIcon(":/images/tlclogo4.png"))
    window = StartWindow()
    window.setWindowIcon(QIcon(":/images/tlclogo4.png"))
    window.show()
    sys.exit(app.exec_())