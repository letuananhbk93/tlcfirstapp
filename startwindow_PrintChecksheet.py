from PyQt5 import QtWidgets
from PyQt5.QtGui import QIcon
from PyQt5 import QtCore
from startwindow_ui import Ui_StartWindow  # Adjust if your generated class name is different
from ColourMasterBox_collection import ColorSearchApp  # Adjust class name if needed
from warehouse_form import WarehouseForm  # Adjust class name if needed
from PyQt5.QtCore import QTranslator, QLocale, QLibraryInfo
import sys
import resources_rc

class StartWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_StartWindow()
        self.ui.setupUi(self)
        self.ui.QCappButton.clicked.connect(self.open_qc_app)
        self.ui.WHappButton.clicked.connect(self.open_wh_app)

        self.translator = QtCore.QTranslator()
        self.ui.LanguageBox.addItem(QIcon(":/images/vi.png"), "Tiếng Việt", "vi")
        self.ui.LanguageBox.addItem(QIcon(":/images/en.png"), "English", "en")
        self.ui.LanguageBox.currentIndexChanged.connect(self.change_language)
        self.ui.LanguageBox.setMinimumWidth(140)
    
    def change_language(self):
        lang_code = self.ui.LanguageBox.currentData()
        if lang_code:
            QtWidgets.QApplication.instance().removeTranslator(self.translator)
            if self.translator.load(f"app_{lang_code}.qm"):
                QtWidgets.QApplication.instance().installTranslator(self.translator)
            # Retranslate UI
            self.ui.retranslateUi(self)

    def open_qc_app(self):
        self.qc_window = ColorSearchApp()
        self.qc_window.show()

    def open_wh_app(self):
        self.wh_window = WarehouseForm()
        self.wh_window.show()

if __name__ == "__main__":
    global lang_code
    app = QtWidgets.QApplication(sys.argv)
    app.setWindowIcon(QIcon(":/images/tlclogo4.png"))
    translator = QTranslator()
    lang_code = "vi"  # or "en", or use a config/setting
    translator.load(f"app_{lang_code}.qm")
    app.installTranslator(translator)

    window = StartWindow()
    window.setWindowIcon(QIcon(":/images/tlclogo4.png"))
    window.show()
    sys.exit(app.exec_())