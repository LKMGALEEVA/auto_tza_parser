"""
Модуль для извлечения таблиц из PDF документов.

Содержит класс myPdfReader для поиска и извлечения табличных данных
из технических заданий.
"""
import os
import re
import pypdfium2 as pdfium
import pdfplumber
from openpyxl.workbook import Workbook


# Паттерны для поиска таблиц в документах
TABLE_PATTERNS = [
    r'Технологическое задание.[а-яА-Я ]+контроля',
    r'Технологическое задание.[а-яА-Я ]+арматура',
    r"Технологическое задание. Регулирующая арматура",
    r"Технологическое задание. Запорная арматура",
    r"Технологическое задание. Дистанционные указатели положения",
    r"Технологическое задание. Регуляторы",
    r"Технологическое задание. Механизмы",
    r"Технологическое задание. Точки контроля",
    r"Технологическое задание. Точки теплотехнического контроля",
    r"Технологическое задание. Точки радиационного контроля",
    r"Технологическое задание. ДУП",
    r"Технологическое задание. Алгоритмы",
    r"Технологическое задание. Сигналы",
    r"Перечень алгоритмов",
    r"Перечень оборудования",
    r"Перечень сигналов"
]

COMBINED_PATTERN = '|'.join(map(re.escape, TABLE_PATTERNS))
PAGE_NUMBER_PATTERN = r'[A-Z]{3}\.[0-9]{4}\.[0-9]{1,2}[A-Z]{0,3}\.[0-9A-Z]{1,3}\.[A-Z]{2}\.[A-Z]{2}0{3}[0-9]/[0-9]+(?:\.[0-9]+)?'


class myPdfReader:
    """
    Класс для чтения и извлечения таблиц из PDF документов.
    
    Attributes:
        names: Список названий найденных таблиц
        start: Начальная страница для поиска
    """
    

    def __init__(self):
        self.names = []
        self.start_page = 6

    def find_blocks(self, file_path):
        """
        Находит блоки страниц по номерам в нижней части страницы.
        
        Args:
            file_path: Путь к PDF файлу
            
        Returns:
            dict: Словарь {номер_страницы: номер_блока}
        """
        page_blocks = {}
        
        with open(file_path, 'rb') as file:
            data = file.read()
            pdf = pdfium.PdfDocument(data)
            
            for i in range(self.start_page, len(pdf)):
                page = pdf.get_page(i)
                textpage = page.get_textpage()
                text = textpage.get_text_bounded()
                
                match = re.findall(PAGE_NUMBER_PATTERN, text)
                if match:
                    page_number = match[0].split('/')[-1]
                    page_blocks[str(i + 1)] = float(page_number)
                    
                textpage.close()
                page.close()
                
            pdf.close()
            
        return page_blocks

    def find_imp_pages(self, file_path):
        """
        Находит страницы с важными таблицами в PDF.
        
        Args:
            file_path: Путь к PDF файлу
            
        Returns:
            list: Список множеств с номерами страниц для каждой таблицы
        """
        with open(file_path, 'rb') as file:
            data = file.read()
            pdf = pdfium.PdfDocument(data)
            important_pages = []
            blocks = self.find_blocks(file_path)
            current_block = 0
            table_index = -1
            
            for i in range(self.start_page, len(pdf) - 1):
                page = pdf.get_page(i)
                textpage = page.get_textpage()
                text = textpage.get_text_bounded()
                
                # Поиск начала новой таблицы
                matches = re.findall(COMBINED_PATTERN, text)
                if matches:
                    for table_name in matches:
                        table_index += 1
                        important_pages.append({str(i + 1)})
                        self.names.append(table_name)
                        current_block = int(blocks.get(str(i + 1), 0))
                
                # Добавление продолжения таблицы
                if current_block == int(blocks.get(str(i + 1), 0)):
                    if table_index >= 0:
                        important_pages[table_index].add(str(i + 1))
                
                textpage.close()
                page.close()
                
            pdf.close()
            
        print(self.names)
        return important_pages

    def extract_imp_tables(self, file_path, pages):
        """
        Извлекает таблицы из указанных страниц PDF.
        
        Args:
            file_path: Путь к PDF файлу
            pages: Список номеров страниц для извлечения
            
        Returns:
            list: Список извлеченных таблиц
        """
        pages = sorted(pages)
        extracted_tables = []
        
        pdf = pdfplumber.open(file_path)
        for page_num in pages:
            page = pdf.pages[page_num - 1]
            table = page.extract_table()
            if table:
                extracted_tables.append(table)
        
        pdf.close()
        return extracted_tables
