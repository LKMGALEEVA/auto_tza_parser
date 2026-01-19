import sys

from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QMessageBox, QLineEdit
import os
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

from full_parsing_2 import mTP
from tzaSchemes import my_obj_finder, cleanup_output_images
from tzaTables import myPdfReader


class FileProcessorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def main(self, files_path, xls_path):

        path_to_pdf = f"{xls_path}/SCHEMES"
        path_to_xlsx = f"{xls_path}/TABLES"


        if not os.path.exists(path_to_pdf):
            os.makedirs(path_to_pdf)
        if not os.path.exists(path_to_xlsx):
            os.makedirs(path_to_xlsx)

        #----------------схемы
        dir_list = os.listdir(files_path)
        for f in dir_list:
            filepath = os.path.abspath(f"{files_path}/{f}")
            kks = f.strip('.pdf')

            myOF = my_obj_finder(filepath)
            myOF.find_imp_pages()
            res = myOF.look_for_objects()
            cords = myOF.get_coords(res)
            d = myOF.get_KKS(cords)
            myOF.write_excel(d, f"{xls_path}/SCHEMES/{kks}.xlsx")
        #-------------------------таблицы
        #
        # if hasattr(self, 'input_folder') and hasattr(self, 'output_folder'):
        #     # Получаем список PDF файлов в выбранной папке
        #     pdf_files = [f for f in os.listdir(self.input_folder) if f.lower().endswith('.pdf')]
        #
        #     if pdf_files:
        for f in dir_list:
            filepath = os.path.abspath(f"{files_path}/{f}")
            k = myPdfReader()
            imp_page_groups = k.find_imp_pages(filepath)
            i = 0
            kks = f.strip('.pdf')
            for group in imp_page_groups:
                imp_pages = [int(x) for x in group]
                l = myPdfReader()
                tables = l.extract_imp_tables(filepath, pages=imp_pages)
                ind = 0
                list_df = []
                combined_data = []
                if k.names.count(k.names[i]) > 1:
                    excelPath = f'{xls_path}/TABLES/{kks}_{k.names[i]}_{i}.xlsx'
                else:
                    excelPath = f'{xls_path}/TABLES/{kks}_{k.names[i]}.xlsx'

                workbook = Workbook()
                workbook.active.title = "CommonList"
                worksheet = workbook['CommonList']

                if tables:
                    combined_data.extend(tables[0])

                for table in tables[1:]:
                    if k.names[i] in ["Перечень сигналов", "Перечень алгоритмов"]:
                        combined_data.extend(table[1:])
                    else:
                        combined_data.extend(table[4:])
                for row in combined_data:
                    worksheet.append(row)

                worksheet.insert_rows(1)
                worksheet['A1'] = f"{k.names[i]}"

                # Проходим по всем ячейкам
                for row in worksheet.iter_rows():
                    for cell in row:
                        if isinstance(cell.value, str):  # Проверяем, является ли значение ячейки строкой
                            # Удаляем (cid:13) из текста
                            cell.value = cell.value.replace('(cid:13)', '')

                workbook.save(excelPath)
                print(f'{kks}_{k.names[i]}.xlsx is ready')
                i += 1

        #self.show_message("Успешно!", "PDF файлы обработаны и сохранены.")

            #------------------------------парсинг
        dir_list = os.listdir(f"{xls_path}/SCHEMES")
        for f in dir_list:
            kks = f.replace(".xlsx", "")  # безопаснее, чем strip()
            mtp = mTP()
            scheme_path = f"{xls_path}/SCHEMES/{f}"
            table_path = f"{xls_path}/TABLES"
            mtp.parse_all_related(kks, scheme_path, table_path)

    def initUI(self):
        # Настройка основного окна
        self.setWindowTitle("Обработка технологических схем (ТЗА)")
        self.setGeometry(400, 400, 600, 300)
        self.setStyleSheet("background-color: rgb(253, 245, 230);")  # Светло-голубой фон

        # Создание элементов интерфейса
        self.input_label = QLabel("Путь к папке с исходными данными:", self)
        self.input_path = QLineEdit(self)
        self.input_button = QPushButton("Выбрать папку", self)
        self.input_button.clicked.connect(self.select_input_folder)

        self.output_label = QLabel("Путь к папке для выгрузки результатов:", self)
        self.output_path = QLineEdit(self)
        self.output_button = QPushButton("Выбрать папку", self)
        self.output_button.clicked.connect(self.select_output_folder)

        self.process_button = QPushButton("Выполнить обработку", self)
        self.process_button.clicked.connect(self.process_files)

        # Размещение элементов в вертикальном layout
        layout = QVBoxLayout()
        layout.addWidget(self.input_label)
        layout.addWidget(self.input_path)
        layout.addWidget(self.input_button)
        layout.addWidget(self.output_label)
        layout.addWidget(self.output_path)
        layout.addWidget(self.output_button)
        layout.addWidget(self.process_button)

        self.setLayout(layout)


    def select_input_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку с ТЗА")
        if folder:
            self.input_path.setText(folder)


    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку для выгрузки результатов")
        if folder:
            self.output_path.setText(folder)


    def process_files(self):
        input_folder = self.input_path.text()
        output_folder = self.output_path.text()

        if not input_folder or not output_folder:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, укажите оба пути.")
            return

        self.main(self.input_path.text(), self.output_path.text())

        # Для демонстрации просто покажем сообщение об успешной обработке
        QMessageBox.information(self, "Успех", "Файлы успешно обработаны!")

    def closeEvent(self, event):
        cleanup_output_images()
        super().closeEvent(event)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    # Глобальная настройка шрифта для всего приложения
    app_font = QFont("Calibri", 14, QFont.Light)  # Шрифт Arial, размер 14, полужирный
    app.setFont(app_font)
    ex = FileProcessorApp()
    ex.show()
    sys.exit(app.exec_())