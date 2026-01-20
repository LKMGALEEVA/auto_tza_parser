import os
import sys
import shutil
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QLineEdit, QVBoxLayout, QFileDialog, QMessageBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QFont

import os
import re
from ultralytics import YOLO
from openpyxl.workbook import Workbook
import pypdfium2 as pdfium

OUTPUT_IMAGES_DIR = os.path.join(os.path.dirname(__file__), "output_images")


def ensure_output_images_dir():
    os.makedirs(OUTPUT_IMAGES_DIR, exist_ok=True)


def cleanup_output_images():
    if not os.path.isdir(OUTPUT_IMAGES_DIR):
        return
    for entry in os.listdir(OUTPUT_IMAGES_DIR):
        path = os.path.join(OUTPUT_IMAGES_DIR, entry)
        if os.path.isdir(path):
            shutil.rmtree(path, ignore_errors=True)
        else:
            try:
                os.remove(path)
            except FileNotFoundError:
                continue

class my_obj_finder:

    def __init__(self, filePath):
        self.filePath = filePath
        self.pages = []

    def find_imp_pages(self):
        pages = []
        with open(self.filePath, 'rb') as file:
            data = file.read()
            pdf = pdfium.PdfDocument(data)
            if len(pdf) < 30:
                for i in range(3, len(pdf)):
                    page = pdf.get_page(i)
                    textpage = page.get_textpage()
                    text = textpage.get_text_bounded()
                    pattern = r'P&I диаграмма|Принципиальная схема'   #нельзя искать P&I диаграмма, потому что это есть в ведомости документов
                    match = re.findall(pattern, text)
                    if len(match) > 0:
                        pages.append(i)
                    self.pages = pages
                    textpage.close()
                    page.close()
                pdf.close()
            else:
                for i in range(3, 30):
                    page = pdf.get_page(i)
                    textpage = page.get_textpage()
                    text = textpage.get_text_bounded()
                    pattern = r'P&I диаграмма|Принципиальная схема'   #нельзя искать P&I диаграмма, потому что это есть в ведомости документов
                    match = re.findall(pattern, text)
                    if len(match) > 0:
                        pages.append(i)
                    self.pages = pages
                    textpage.close()
                    page.close()
                pdf.close()
            return pages

    def look_for_objects(self):
        model = YOLO('best_v5.pt')

        results = []
        pdf = pdfium.PdfDocument(self.filePath)
        ensure_output_images_dir()
        for page_number in self.pages:
            page = pdf.get_page(page_number)
            image_r = page.render(scale=2).to_pil()
            image_path = os.path.join(OUTPUT_IMAGES_DIR, f"output_{page_number:03d}.png")
            image_r.save(image_path)
            results.append(model.predict(image_path, imgsz=2560))  #добавила тут квадратные скобки убрать
            page.close()
        pdf.close()
        print(results)
        return results

    def get_coords(self, results):
        # coords = []
        # j = 0
        # for i in range(len(results)):
        #     for result in results[i]:
        #         coords.append([])
        #         boxes = result.boxes  # Получаем координаты bounding boxes
        #         for box in boxes:
        #             x1, y1, x2, y2 = box.xyxyn[0]  #xyxyn - нормлизованные координаты
        #             #----предположим, он отсчитывает от левого верхнего угла
        #             x1n = x1
        #             x2n = x2
        #             y1n = (1 - y1)
        #             y2n = (1 - y2)
        #             #----------------
        #             class_id = int(box.cls[0].item())
        #             if class_id in [2, 3, 5]:
        #                 coords[j].append((x1n, y1n, x2n, y2n))
        #         j += 1
        # return coords
        coords = []
        j = 0
        for i in range(len(results)):    #страницы
            for result in results[i]:         #найденные объекты на страниц
                coords.append({})
                boxes = result.boxes  # Получаем координаты bounding boxes
                for box in boxes:
                    x1, y1, x2, y2 = box.xyxyn[0]  #xyxyn - нормлизованные координаты
                    #----предположим, он отсчитывает от левого верхнего угла
                    x1n = x1
                    x2n = x2
                    y1n = (1 - y1)
                    y2n = (1 - y2)
                    #----------------
                    class_id = int(box.cls[0].item())
                    #if class_id in [2, 3, 5]:
                    class_name = result.names.get(class_id) if hasattr(result, "names") else None
                    if class_id in [2, 3, 5] or class_name == "motor_operated_valve":
                        coords[j][(x1n, y1n, x2n, y2n)] = "Нечто"
                j += 1
        print(coords)
        return coords     #список словарей [{(): "", (): "", (): ""}, {}, ...]


    def get_KKS(self, coords):
        # dict_list = []
        # pdf = pdfium.PdfDocument(self.filePath)
        # for imp_page in self.pages:
        #     page = pdf.get_page(imp_page)
        #     textpage = page.get_textpage()
        #     bb = page.get_bbox()
        #     print(bb)
        #     p = 0
            # for coords_group in coords[p]:
            #     x1, y1, x2, y2 = coords_group
            #     #---------------кринж матеша
            #     x1n = x1*bb[2]
            #     x2n = x2*bb[2]
            #     y1n = y1*bb[3]
            #     y2n = y2*bb[3]
            #     # ---------------кринж матеша
            #     extracted_text = textpage.get_text_bounded(x1n, y1n, x2n, y2n)
            #     # print(extracted_text)
            #     # print("#------------------")
            #     new_text = "".join(extracted_text.split("\r"))
            #     new_new_text = "".join(new_text.split("\n"))
            #     dict_list.append(new_new_text)
            #     p += 1
        #     textpage.close()
        #     page.close()
        # pdf.close()
        # fin_dict = []
        # for thing in dict_list:
        #     print(thing)
        #     matches = re.findall(r"(\w{2})(\d{3})(\d{2})(\w{3})(\d{2})", thing)
        #     results = [f"{match[2]}{match[3]}{match[4]}{match[0]}{match[1]}" for match in matches]
        #     for result in results:
        #         fin_dict.append(result)
        # return fin_dict
        dict_list = dict()
        pdf = pdfium.PdfDocument(self.filePath)
        p = 0
        for imp_page in self.pages:
            page = pdf.get_page(imp_page)
            textpage = page.get_textpage()
            bb = page.get_bbox()
            print(bb)
            for key_coord in coords[p].keys():
                x1, y1, x2, y2 = key_coord
                #---------------математика
                x1n = x1*bb[2]
                x2n = x2*bb[2]
                y1n = y1*bb[3]
                y2n = y2*bb[3]
                # ---------------математика
                extracted_text = textpage.get_text_bounded(x1n, y1n, x2n, y2n)
                # print(extracted_text)
                # print("#------------------")
                new_text = "".join(extracted_text.split("\r"))
                new_new_text = "".join(new_text.split("\n"))
                dict_list[new_new_text] = coords[p][key_coord]
            p += 1
            textpage.close()
            page.close()
        pdf.close()
        fin_dict = dict()
        for thing in dict_list.keys():
            print(thing)
            matches1 = re.findall(r"(AA|AP|CT|CL|CF|CP|CY|CS|CQ|CU|CM|CR)(\d{3})([A-D]?)(\d{2})(\w{3})(\d{2})", thing)
            matches2 = re.findall(r"(\d{2})(\w{3})(\d{2})(AA|AP|CT|CL|CF|CP|CY|CS|CQ|CU|CM|CR)(\d{3})([A-D]?)", thing)
            finds = [f"{match[3]}{match[4]}{match[5]}{match[0]}{match[1]}{match[2]}" for match in matches1]
            for match in matches2:
                finds.append(f"{match[0]}{match[1]}{match[2]}{match[3]}{match[4]}{match[5]}")
            for result in finds:
                if "AA" in result:
                    fin_dict[result] = "Арматура"
                elif "AP" in result:
                    fin_dict[result] = "Насос"
                else:
                    fin_dict[result] = "Датчик"
        print(fin_dict)
        return fin_dict



    def write_excel(self, dict_list, output_filename: str, sheet_name='Sheet1'):
        # if not os.path.exists('checks'):
        #     os.makedirs('checks')

        # filepath = os.path.join('checks', output_filename)
        filepath = output_filename

        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name

        ws.append(["KKS", "Object"])  # Записываем заголовки

        for key, value in dict_list.items():
            ws.append([key, value])

        wb.save(filepath)
        return filepath
        #     # Записываем данные из списка словарей в Excel
        # if dict_list:
        #     # header = list(dict_list[0].keys())
        #     # ws.append(header)  # Записываем заголовки
        #
        #     for row in dict_list:
        #         ws.append([row])
        #
        # wb.save(filepath)
        # return filepath

    def main(self, files_path, xls_path):
        dir_list = os.listdir(files_path)
        for f in dir_list:
            filepath = os.path.abspath(f"{files_path}/{f}")
            kks = f.strip('.pdf')

            myOF = my_obj_finder(filepath)
            myOF.find_imp_pages()
            res = myOF.look_for_objects()
            cords = myOF.get_coords(res)
            d = myOF.get_KKS(cords)
            myOF.write_excel(d, f"{xls_path}/{kks}.xlsx")


class FileProcessorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def main(self, files_path, xls_path):
        dir_list = os.listdir(files_path)
        for f in dir_list:
            filepath = os.path.abspath(f"{files_path}/{f}")
            kks = f.strip('.pdf')

            myOF = my_obj_finder(filepath)
            myOF.find_imp_pages()
            res = myOF.look_for_objects()
            cords = myOF.get_coords(res)
            d = myOF.get_KKS(cords)
            myOF.write_excel(d, f"{xls_path}/{kks}.xlsx")
#
#     def initUI(self):
#         # Настройка основного окна
#         self.setWindowTitle("Обработка технологических схем (ТЗА)")
#         self.setGeometry(400, 400, 600, 300)
#         self.setStyleSheet("background-color: rgb(253, 245, 230);")  # Светло-голубой фон
#
#
#
#         # Создание элементов интерфейса
#         self.input_label = QLabel("Путь к папке с исходными данными:", self)
#         self.input_path = QLineEdit(self)
#         self.input_button = QPushButton("Выбрать папку", self)
#         self.input_button.clicked.connect(self.select_input_folder)
#
#         self.output_label = QLabel("Путь к папке для выгрузки результатов:", self)
#         self.output_path = QLineEdit(self)
#         self.output_button = QPushButton("Выбрать папку", self)
#         self.output_button.clicked.connect(self.select_output_folder)
#
#         self.process_button = QPushButton("Выполнить обработку", self)
#         self.process_button.clicked.connect(self.process_files)
#
#         # Размещение элементов в вертикальном layout
#         layout = QVBoxLayout()
#         layout.addWidget(self.input_label)
#         layout.addWidget(self.input_path)
#         layout.addWidget(self.input_button)
#         layout.addWidget(self.output_label)
#         layout.addWidget(self.output_path)
#         layout.addWidget(self.output_button)
#         layout.addWidget(self.process_button)
#
#         self.setLayout(layout)
#
#     def select_input_folder(self):
#         folder = QFileDialog.getExistingDirectory(self, "Выберите папку с ТЗА")
#         if folder:
#             self.input_path.setText(folder)
#
#     def select_output_folder(self):
#         folder = QFileDialog.getExistingDirectory(self, "Выберите папку для выгрузки результатов")
#         if folder:
#             self.output_path.setText(folder)
#
#     def process_files(self):
#         input_folder = self.input_path.text()
#         output_folder = self.output_path.text()
#
#         if not input_folder or not output_folder:
#             QMessageBox.warning(self, "Ошибка", "Пожалуйста, укажите оба пути.")
#             return
#
#
#         self.main(self.input_path.text(), self.output_path.text())
#
#         # Для демонстрации просто покажем сообщение об успешной обработке
#         QMessageBox.information(self, "Успех", "Файлы успешно обработаны!")
#
# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#
#     # Глобальная настройка шрифта для всего приложения
#     app_font = QFont("Calibri", 14, QFont.Light)  # Шрифт Arial, размер 14, полужирный
#     app.setFont(app_font)
#
#     ex = FileProcessorApp()
#     ex.show()
#     sys.exit(app.exec_())