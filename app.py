"""
Главное приложение для автоматизации верификации технологических схем (ТЗА).

Модуль содержит графический интерфейс пользователя для обработки PDF документов
с технологическими схемами и таблицами.
"""
import sys
import os

from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, 
    QFileDialog, QLabel, QMessageBox, QLineEdit, QTextEdit
)
from openpyxl import Workbook

from full_parsing_2 import mTP
from tzaSchemes import my_obj_finder, cleanup_output_images
from tzaTables import myPdfReader


class FileProcessorApp(QWidget):
    """
    Главное окно приложения для обработки технологических схем.
    
    Предоставляет графический интерфейс для выбора папок с исходными данными
    и запуска процесса обработки PDF файлов.
    """
    
    def __init__(self):
        super().__init__()
        self.initUI()

    def log_message(self, message):
        """
        Добавляет сообщение в поле логов.
        
        Args:
            message: Текст сообщения для отображения
        """
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
        QApplication.processEvents()

    def process_files(self, files_path, xls_path):
        """
        Основной метод обработки файлов.
        
        Args:
            files_path: Путь к папке с исходными PDF файлами
            xls_path: Путь к папке для сохранения результатов
        """
        schemes_path = os.path.join(xls_path, "SCHEMES")
        tables_path = os.path.join(xls_path, "TABLES")

        self.log_message("Начало обработки файлов...")
        self.log_message(f"Входная папка: {files_path}")
        self.log_message(f"Выходная папка: {xls_path}")

        # Создание необходимых директорий
        self._create_output_directories(schemes_path, tables_path)
        
        # Обработка схем
        self._process_schemes(files_path, schemes_path)
        
        # Обработка таблиц
        self._process_tables(files_path, tables_path)
        
        # Финальный парсинг и сравнение данных
        self._parse_and_compare(schemes_path, tables_path, xls_path)
        
        self.log_message("\n=== Обработка завершена успешно! ===")

    def _create_output_directories(self, schemes_path, tables_path):
        """Создает необходимые директории для выходных данных."""
        for path in [schemes_path, tables_path]:
            if not os.path.exists(path):
                os.makedirs(path)
                self.log_message(f"Создана папка: {path}")

    def _process_schemes(self, files_path, schemes_path):
        """
        Обрабатывает технологические схемы.
        
        Args:
            files_path: Путь к папке с PDF файлами
            schemes_path: Путь к папке для сохранения схем
        """
        self.log_message("\n=== Обработка схем ===")
        dir_list = os.listdir(files_path)
        self.log_message(f"Найдено файлов: {len(dir_list)}")

        for idx, filename in enumerate(dir_list, 1):
            if not filename.lower().endswith('.pdf'):
                continue
                
            filepath = os.path.abspath(os.path.join(files_path, filename))
            kks = filename.replace('.pdf', '')

            self.log_message(f"[{idx}/{len(dir_list)}] Обработка схемы: {filename}")
            
            obj_finder = my_obj_finder(filepath)
            obj_finder.find_imp_pages()
            results = obj_finder.look_for_objects()
            coords = obj_finder.get_coords(results)
            kks_dict = obj_finder.get_KKS(coords)
            
            output_file = os.path.join(schemes_path, f"{kks}.xlsx")
            obj_finder.write_excel(kks_dict, output_file)
            self.log_message(f"  ✓ Создан файл: {kks}.xlsx")
    def _process_tables(self, files_path, tables_path):
        """
        Извлекает таблицы из PDF файлов.
        
        Args:
            files_path: Путь к папке с PDF файлами
            tables_path: Путь к папке для сохранения таблиц
        """
        self.log_message("\n=== Обработка таблиц ===")
        dir_list = os.listdir(files_path)

        for idx, filename in enumerate(dir_list, 1):
            if not filename.lower().endswith('.pdf'):
                continue
                
            filepath = os.path.abspath(os.path.join(files_path, filename))
            self.log_message(f"[{idx}/{len(dir_list)}] Извлечение таблиц из: {filename}")
            
            pdf_reader = myPdfReader()
            page_groups = pdf_reader.find_imp_pages(filepath)
            kks = filename.replace('.pdf', '')
            
            for group_idx, group in enumerate(page_groups):
                pages = [int(x) for x in group]
                table_reader = myPdfReader()
                tables = table_reader.extract_imp_tables(filepath, pages=pages)
                
                # Формирование имени файла
                table_name = pdf_reader.names[group_idx]
                if pdf_reader.names.count(table_name) > 1:
                    excel_path = os.path.join(tables_path, f"{kks}_{table_name}_{group_idx}.xlsx")
                else:
                    excel_path = os.path.join(tables_path, f"{kks}_{table_name}.xlsx")

                # Создание Excel файла
                self._create_table_excel(tables, excel_path, table_name, 
                                        pdf_reader.names[group_idx])
                self.log_message(f"  ✓ Создан файл: {kks}_{table_name}.xlsx")

    def _create_table_excel(self, tables, excel_path, table_name, original_name):
        """
        Создает Excel файл с извлеченными таблицами.
        
        Args:
            tables: Список таблиц для сохранения
            excel_path: Путь к выходному файлу
            table_name: Название таблицы
            original_name: Оригинальное название из PDF
        """
        workbook = Workbook()
        workbook.active.title = "CommonList"
        worksheet = workbook['CommonList']

        combined_data = []
        if tables:
            combined_data.extend(tables[0])

        # Объединение данных из всех таблиц
        for table in tables[1:]:
            if original_name in ["Перечень сигналов", "Перечень алгоритмов"]:
                combined_data.extend(table[1:])
            else:
                combined_data.extend(table[4:])

        # Запись данных
        for row in combined_data:
            worksheet.append(row)

        worksheet.insert_rows(1)
        worksheet['A1'] = table_name

        # Очистка артефактов PDF
        for row in worksheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    cell.value = cell.value.replace('(cid:13)', '')

        workbook.save(excel_path)

    def _parse_and_compare(self, schemes_path, tables_path, output_path):
        """
        Выполняет финальный парсинг и сравнение данных схем и таблиц.
        
        Args:
            schemes_path: Путь к папке со схемами
            tables_path: Путь к папке с таблицами
            output_path: Путь для сохранения результатов
        """
        self.log_message("\n=== Финальный парсинг данных ===")
        scheme_files = os.listdir(schemes_path)
        self.log_message(f"Обработка {len(scheme_files)} схем...")

        for idx, filename in enumerate(scheme_files, 1):
            if not filename.endswith('.xlsx'):
                continue
                
            kks = filename.replace(".xlsx", "")
            self.log_message(f"[{idx}/{len(scheme_files)}] Парсинг: {kks}")
            
            parser = mTP()
            scheme_path = os.path.join(schemes_path, filename)
            parser.parse_all_related(kks, scheme_path, tables_path)
            self.log_message(f"  ✓ Завершено: {kks}")

    def initUI(self):
        """Инициализирует пользовательский интерфейс."""
        self.setWindowTitle("Обработка технологических схем (ТЗА)")
        self.setGeometry(200, 100, 1200, 800)
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
        self.process_button.clicked.connect(self.start_processing)

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
        """Открывает диалог выбора папки с исходными данными."""
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку с ТЗА")
        if folder:
            self.input_path.setText(folder)

    def select_output_folder(self):
        """Открывает диалог выбора папки для результатов."""
        folder = QFileDialog.getExistingDirectory(
            self, "Выберите папку для выгрузки результатов"
        )
        if folder:
            self.output_path.setText(folder)

    def start_processing(self):
        """Запускает процесс обработки файлов."""
        input_folder = self.input_path.text()
        output_folder = self.output_path.text()

        if not input_folder or not output_folder:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, укажите оба пути.")
            return

        self.log_text.clear()

        try:
            self.process_files(input_folder, output_folder)
            QMessageBox.information(self, "Успех", "Файлы успешно обработаны!")
        except Exception as e:
            self.log_message(f"\n❌ ОШИБКА: {str(e)}")
            QMessageBox.critical(
                self, "Ошибка", f"Произошла ошибка при обработке:\n{str(e)}"
            )

    def closeEvent(self, event):
        """Обработчик события закрытия окна."""
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