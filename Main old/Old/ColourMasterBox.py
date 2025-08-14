import sys
import os
import pandas as pd
from PyQt5 import QtWidgets
from form import Ui_Form  # form.py sinh ra từ form.ui

class ColorSearchApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.ui = Ui_Form()
        self.ui.setupUi(self)

        server_path = r"C:\Users\Admins\The Lacquer Company\Company Files - Tài liệu\THE LACQUER COMPANY - VIETNAM OFFICE\QC FOLDER\MASTER COLOR LIST QC\TLC - MASTER COLOR LIST FOR PRODUCTION.xlsx"
        local_path = 'TLC - MASTER COLOR LIST FOR PRODUCTION.xlsx'

        try:
            self.df_main = pd.read_excel(server_path, sheet_name='Lacquer FIN')
            self.df_custom = pd.read_excel(server_path, sheet_name='Custom color')
            self.df_effect = pd.read_excel(server_path, sheet_name='Effect Color Swatch Statistics')
        except Exception:
            self.df_main = pd.read_excel(local_path, sheet_name='Lacquer FIN')
            self.df_custom = pd.read_excel(local_path, sheet_name='Custom color')
            self.df_effect = pd.read_excel(local_path, sheet_name='Effect Color Swatch Statistics')

        # Strip spaces from column names for robustness
        self.df_main.columns = self.df_main.columns.str.strip()
        self.df_custom.columns = self.df_custom.columns.str.strip()
        self.df_effect.columns = self.df_effect.columns.str.strip()

        self.ui.pushButton.clicked.connect(self.search_color)
        self.ui.HieuUngButton.clicked.connect(self.search_effect_color)  # Connect the new button

    def search_color(self):
        keyword = self.ui.lineEdit.text().strip().lower()

        # Search in 'Lacquer FIN'
        matched = self.df_main[self.df_main['Name'].str.lower().str.contains(keyword, na=False)]
        color_type = "STANDARD COLOR"

        # If not found, search in 'Custom color'
        if matched.empty:
            matched = self.df_custom[self.df_custom['Color name'].str.lower().str.contains(keyword, na=False)]
            color_type = "CUSTOM COLOR"
            if matched.empty:
                self.ui.textEdit.setText("🔍 Không tìm thấy kết quả.")
                return

        # Add the first sentence with bold and underline
        html = f'<p><b><u>{color_type}</u></b></p>'

        for _, row in matched.iterrows():
            for col in matched.columns:
                value = str(row[col]) if pd.notna(row[col]) else ""
                if col.strip().lower() in ["ref image"]:
                    if os.path.exists(value):
                        html += f'<p><b>{col}:</b><br><img src="{value}" width="200"></p>'
                    else:
                        html += f"<p><b>{col}:</b> (Không tìm thấy ảnh)</p>"
                else:
                    html += f"<p><b>{col}:</b> {value}</p>"
            html += "<hr>"

        self.ui.textEdit.setHtml(html)

    def search_effect_color(self):
        keyword = self.ui.lineEdit.text().strip().lower()
        # Adjust the column name below if needed
        if 'Color Name' in self.df_effect.columns:
            matched = self.df_effect[self.df_effect['Color Name'].str.lower().str.contains(keyword, na=False)]
        else:
            # Print columns for debugging if needed
            print(self.df_effect.columns.tolist())
            self.ui.textEdit.setText("Không tìm thấy cột 'Color Name' trong sheet hiệu ứng.")
            return

        if matched.empty:
            self.ui.textEdit.setText("🔍 Không tìm thấy kết quả trong hiệu ứng.")
            return

        html = '<p><b><u>EFFECT COLOR</u></b></p>'
        for _, row in matched.iterrows():
            for col in matched.columns:
                value = str(row[col]) if pd.notna(row[col]) else ""
                if col.strip().lower() in ["ref image"]:
                    if os.path.exists(value):
                        html += f'<p><b>{col}:</b><br><img src="{value}" width="200"></p>'
                    else:
                        html += f"<p><b>{col}:</b> (Không tìm thấy ảnh)</p>"
                else:
                    html += f"<p><b>{col}:</b> {value}</p>"
            html += "<hr>"

        self.ui.textEdit.setHtml(html)

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = ColorSearchApp()
    window.setWindowTitle("Tra cứu thẻ màu")
    window.resize(800, 600)
    window.show()
    sys.exit(app.exec_())
