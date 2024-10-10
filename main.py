import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt
import os
from PyQt5.QtWidgets import QMenuBar, QAction

class ExcelFilterApp(QWidget):
    
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        # Dastur oyna dizayni
        self.setWindowTitle('Excel-Filter(Monitoring uchun) by Firdavs')
        self.setGeometry(100, 100, 400, 200)
        
        layout = QVBoxLayout()

        # Faylni tanlash tugmasi
        self.label = QLabel('Exel faylni mishkada sudrab ustimga tashla', self)
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)

        self.button = QPushButton('Exel fayl tanlash', self)
        self.button.clicked.connect(self.select_file)
        layout.addWidget(self.button)

        self.setLayout(layout)
        self.setAcceptDrops(True)  # Drag and drop ni yoqish


        # Menyu yaratish
        menubar = QMenuBar(self)
        about_action = QAction('About', self)
        about_action.triggered.connect(self.show_about)
        menubar.addAction(about_action)

        layout.setMenuBar(menubar)

    def show_about(self):
        # About oynasi
        QMessageBox.information(self, 'About', 'Excel Duplicate Link Remover\nAuthor: Firdavs Abdurasulov')

    def dragEnterEvent(self, event):
        # Drag qilingan faylni ko'rsatish
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        # Fayl tashlanganda ishlaydigan funksiya
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.endswith('.xlsx'):
                self.process_file(file_path)

    def select_file(self):
        # Fayl tanlash dialog oynasi
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx)")
        if file_path:
            self.process_file(file_path)

    def process_file(self, file_path):
        try:
            # Excel faylni o'qish
            df = pd.read_excel(file_path)

            # Faqat bitta ustunda takrorlanmalarni olib tashlash
            df_filtered = df.drop_duplicates(subset=[df.columns[1]])

            # Natijani yangi faylga yozish
            output_path = file_path.replace('.xlsx', '_filtered.xlsx')
            df_filtered.to_excel(output_path, index=False)

            # Muvaffaqiyatli natijani ko'rsatish
            QMessageBox.information(self, 'Success', f'Filtered file saved as: {output_path}')
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'An error occurred: {e}')

        



if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ExcelFilterApp()
    window.show()
    sys.exit(app.exec_())
