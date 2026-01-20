import sys

from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QMessageBox, \
    QLineEdit, QTextEdit
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

    def log_message(self, message):
        """Добавить сообщение в поле логов"""
        self.log_text.append(message)
        # Прокручиваем вниз, чтобы показать последнее сообщение
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())
        # Обновляем интерфейс, чтобы изменения отобразились сразу
        QApplication.processEvents()

    def main(self, files_path, xls_path):

        path_to_pdf = f"{xls_path}/SCHEMES"
        path_to_xlsx = f"{xls_path}/TABLES"

        self.log_message("Начало обработки файлов...")
        self.log_message(f"Входная папка: {files_path}")
        self.log_message(f"Выходная папка: {xls_path}")

        if not os.path.exists(path_to_pdf):
            os.makedirs(path_to_pdf)
            self.log_message(f"Создана папка: {path_to_pdf}")
        if not os.path.exists(path_to_xlsx):
            os.makedirs(path_to_xlsx)
            self.log_message(f"Создана папка: {path_to_xlsx}")

        # ----------------схемы
        self.log_message("\n=== Обработка схем ===")
        dir_list = os.listdir(files_path)
        self.log_message(f"Найдено файлов: {len(dir_list)}")

        for idx, f in enumerate(dir_list, 1):
            filepath = os.path.abspath(f"{files_path}/{f}")
            kks = f.strip('.pdf')

            self.log_message(f"[{idx}/{len(dir_list)}] Обработка схемы: {f}")
            myOF = my_obj_finder(filepath)
            myOF.find_imp_pages()
            res = myOF.look_for_objects()
            cords = myOF.get_coords(res)
            d = myOF.get_KKS(cords)
            output_file = f"{xls_path}/SCHEMES/{kks}.xlsx"
            myOF.write_excel(d, output_file)
            self.log_message(f"  ✓ Создан файл: {kks}.xlsx")
        # -------------------------таблицы
        self.log_message("\n=== Обработка таблиц ===")
        #
        # if hasattr(self, 'input_folder') and hasattr(self, 'output_folder'):
        #     # Получаем список PDF файлов в выбранной папке
        #     pdf_files = [f for f in os.listdir(self.input_folder) if f.lower().endswith('.pdf')]
        #
        #     if pdf_files:
        for idx, f in enumerate(dir_list, 1):
            filepath = os.path.abspath(f"{files_path}/{f}")
            self.log_message(f"[{idx}/{len(dir_list)}] Извлечение таблиц из: {f}")
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
                self.log_message(f"  ✓ Создан файл: {kks}_{k.names[i]}.xlsx")
                i += 1

        # self.show_message("Успешно!", "PDF файлы обработаны и сохранены.")

        # ------------------------------парсинг
        self.log_message("\n=== Финальный парсинг данных ===")
        dir_list = os.listdir(f"{xls_path}/SCHEMES")
        self.log_message(f"Обработка {len(dir_list)} схем...")

        for idx, f in enumerate(dir_list, 1):
            kks = f.replace(".xlsx", "")  # безопаснее, чем strip()
            self.log_message(f"[{idx}/{len(dir_list)}] Парсинг: {kks}")
            mtp = mTP()
            scheme_path = f"{xls_path}/SCHEMES/{f}"
            table_path = f"{xls_path}/TABLES"
            mtp.parse_all_related(kks, scheme_path, table_path)
            self.log_message(f"  ✓ Завершено: {kks}")

        self.log_message("\n=== Обработка завершена успешно! ===")

    def initUI(self):
        # Настройка основного окна
        self.setWindowTitle("Обработка технологических схем (ТЗА)")
        self.setGeometry(200, 100, 1200, 800)

        # Минималистичный фон
        self.setStyleSheet("""
            QWidget {
                background-color: #f5f5f5;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
        """)

        # Создание элементов интерфейса
        self.input_label = QLabel("Папка с исходными данными", self)
        self.input_label.setStyleSheet("""
            QLabel {
                color: #333333;
                font-size: 24px;
                font-weight: 500;
                padding: 8px 0px;
            }
        """)

        self.input_path = QLineEdit(self)
        self.input_path.setStyleSheet("""
            QLineEdit {
                padding: 14px 18px;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                background-color: #ffffff;
                color: #333333;
                font-size: 24px;
            }
            QLineEdit:focus {
                border: 1px solid #4a90e2;
            }
        """)
        self.input_path.setMinimumHeight(50)

        self.input_button = QPushButton("Выбрать", self)
        self.input_button.setStyleSheet("""
            QPushButton {
                background-color: #4a90e2;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 10px 28px;
                font-size: 24px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #357abd;
            }
            QPushButton:pressed {
                background-color: #2a5f99;
            }
        """)
        self.input_button.setMinimumHeight(50)
        self.input_button.clicked.connect(self.select_input_folder)

        self.output_label = QLabel("Папка для выгрузки результатов", self)
        self.output_label.setStyleSheet("""
            QLabel {
                color: #333333;
                font-size: 24px;
                font-weight: 500;
                padding: 8px 0px;
                margin-top: 20px;
            }
        """)

        self.output_path = QLineEdit(self)
        self.output_path.setStyleSheet("""
            QLineEdit {
                padding: 14px 18px;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                background-color: #ffffff;
                color: #333333;
                font-size: 24px;
            }
            QLineEdit:focus {
                border: 1px solid #4a90e2;
            }
        """)
        self.output_path.setMinimumHeight(50)

        self.output_button = QPushButton("Выбрать", self)
        self.output_button.setStyleSheet("""
            QPushButton {
                background-color: #4a90e2;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 10px 28px;
                font-size: 24px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #357abd;
            }
            QPushButton:pressed {
                background-color: #2a5f99;
            }
        """)
        self.output_button.setMinimumHeight(50)
        self.output_button.clicked.connect(self.select_output_folder)

        self.process_button = QPushButton("Выполнить обработку", self)
        self.process_button.setStyleSheet("""
            QPushButton {
                background-color: #4a90e2;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 28px;
                font-size: 24px;
                font-weight: 600;
                margin-top: 30px;
            }
            QPushButton:hover {
                background-color: #357abd;
            }
            QPushButton:pressed {
                background-color: #2a5f99;
            }
        """)
        self.process_button.setMinimumHeight(80)
        self.process_button.clicked.connect(self.process_files)

        # Поле для логов
        self.log_label = QLabel("Логи обработки", self)
        self.log_label.setStyleSheet("""
            QLabel {
                color: #333333;
                font-size: 24px;
                font-weight: 500;
                padding: 8px 0px;
                margin-top: 25px;
            }
        """)

        self.log_text = QTextEdit(self)
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #ffffff;
                color: #333333;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 18px;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 14px;
                line-height: 1.6;
            }
        """)
        self.log_text.setMinimumHeight(280)

        # Размещение элементов в вертикальном layout
        layout = QVBoxLayout()
        layout.setSpacing(12)
        layout.setContentsMargins(40, 40, 40, 40)

        layout.addWidget(self.input_label)
        layout.addWidget(self.input_path)
        layout.addWidget(self.input_button)
        layout.addWidget(self.output_label)
        layout.addWidget(self.output_path)
        layout.addWidget(self.output_button)
        layout.addWidget(self.process_button)
        layout.addWidget(self.log_label)
        layout.addWidget(self.log_text, 1)

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

        # Очищаем логи перед новой обработкой
        self.log_text.clear()

        try:
            self.main(self.input_path.text(), self.output_path.text())
            # Для демонстрации просто покажем сообщение об успешной обработке
            QMessageBox.information(self, "Успех", "Файлы успешно обработаны!")
        except Exception as e:
            self.log_message(f"\n❌ ОШИБКА: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка при обработке:\n{str(e)}")

    def closeEvent(self, event):
        cleanup_output_images()
        super().closeEvent(event)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    # Глобальная настройка шрифта для всего приложения
    app_font = QFont("Segoe UI", 14)
    app.setFont(app_font)
    ex = FileProcessorApp()
    ex.show()
    sys.exit(app.exec_())