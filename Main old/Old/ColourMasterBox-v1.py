import sys
import os
import urllib.parse
import urllib.request
import pandas as pd
from PyQt5 import QtWidgets
from form import Ui_Form  # form.py sinh ra từ form.ui

def find_company_folder():
    # List all possible company folder paths
    possible_paths = [
        r"C:\Users\Admins\The Lacquer Company\Company Files - Tài liệu",
        r"C:\Users\Admins\The Lacquer Company\Company Files - Documents"
    ]
    for path in possible_paths:
        if os.path.exists(path):
            return path
    raise RuntimeError("Không tìm thấy thư mục công ty trên máy tính này.")

def all_words_in_text(words, text):
    return all(word in text for word in words)

class ColorSearchApp(QtWidgets.QWidget):    
    def __init__(self):
        super().__init__()
        self.ui = Ui_Form()
        self.ui.setupUi(self)

        try:
            company_folder=find_company_folder()
            self.server_path=os.path.join(
                company_folder,
                "THE LACQUER COMPANY - VIETNAM OFFICE",
                "QC FOLDER",
                "MASTER COLOR LIST QC"
            )
            serverexcel_path = os.path.join(self.server_path,"TLC - MASTER COLOR LIST FOR PRODUCTION.xlsx")
        except RuntimeError as e:
            QtWidgets.QMessageBox.critical(self,"Lỗi",str(e))
            self.close()

        local_path = 'TLC - MASTER COLOR LIST FOR PRODUCTION.xlsx'

        try:
            self.df_main = pd.read_excel(serverexcel_path, sheet_name='Lacquer FIN')
            self.df_custom = pd.read_excel(serverexcel_path, sheet_name='Custom color')
            self.df_effect = pd.read_excel(serverexcel_path, sheet_name='Effect Color Swatch Statistics')
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
        words = keyword.split()
        matched = self.df_main[self.df_main['Name'].str.lower().apply(lambda x: all_words_in_text(words, x) if pd.notna(x) else False)]
        #color_type = "STANDARD COLOR"

        # If not found, search in 'Custom color'
        if matched.empty:
            words = keyword.split()
            matched = self.df_custom[self.df_custom['Color name'].str.lower().apply(lambda x: all_words_in_text(words, x) if pd.notna(x) else False)]
            #color_type = "CUSTOM COLOR"
            if matched.empty:
                self.ui.textEdit.setText("🔍 Không tìm thấy kết quả.")
                return

        # Add the first sentence with bold and underline
        #html = f'<p><b><u>{color_type}</u></b></p>'

        html=""
        for _, row in matched.iterrows():
            name = str(row.get("Name", "")).strip()
            for col in matched.columns:
                value = str(row[col]) if pd.notna(row[col]) else ""
                if col.strip().lower() in ["ref image"]:
                    image_path=os.path.join(self.server_path,"Images",f"{name}.jpg")
                    if os.path.exists(image_path):
                        file_uri=urllib.parse.urljoin('file:',urllib.request.pathname2url(image_path))
                        html += f'<p><b>{col}:</b><br><img src="{file_uri}" width="200"></p>'
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
            words = keyword.split()
            matched = self.df_effect[self.df_effect['Color Name'].str.lower().apply(lambda x: all_words_in_text(words, x) if pd.notna(x) else False)]
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
            name=str(row.get("Color Name","")).strip()
            for col in matched.columns:
                value = str(row[col]) if pd.notna(row[col]) else ""
                if col.strip().lower() in ["ref image"]:
                    image_path=os.path.join(self.server_path,"Images",f"{name}.jpg")
                    if os.path.exists(image_path):
                        file_uri=urllib.parse.urljoin('file:',urllib.request.pathname2url(image_path))
                        html += f'<p><b>{col}:</b><br><img src="{file_uri}" width="200"></p>'
                    else:
                        html += f"<p><b>{col}:</b> (Không tìm thấy ảnh)</p>"
                else:
                    html += f"<p><b>{col}:</b> {value}</p>"
            html += "<hr>"

        self.ui.textEdit.setHtml(html)

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = ColorSearchApp()
    window.setWindowTitle("The Lacquer Company application")
    window.resize(800, 600)
    window.show()
    sys.exit(app.exec_())
