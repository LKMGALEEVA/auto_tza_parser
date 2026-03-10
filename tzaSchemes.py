import os
import re
import shutil
from ultralytics import YOLO
from openpyxl.workbook import Workbook
import pypdfium2 as pdfium

OUTPUT_IMAGES_DIR = os.path.join(os.path.dirname(__file__), "output_images")


def ensure_output_images_dir():
    """Создает директорию для временных изображений, если она не существует."""
    os.makedirs(OUTPUT_IMAGES_DIR, exist_ok=True)


def cleanup_output_images():
    """Очищает директорию с временными изображениями."""
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
    """
    Класс для обнаружения и извлечения объектов из технологических схем.
    
    Attributes:
        filePath: Путь к PDF файлу
        pages: Список индексов страниц с важными схемами
    """
    
    def __init__(self, filePath):
        self.filePath = filePath
        self.pages = []

    def find_imp_pages(self):
        """
        Находит страницы с P&I диаграммами или принципиальными схемами в PDF.
        
        Returns:
            list: Список индексов страниц с важными схемами
        """
        pages = []
        pattern = r'P&I диаграмма|Принципиальная схема|P&I diagram'
        
        with open(self.filePath, 'rb') as file:
            data = file.read()
            pdf = pdfium.PdfDocument(data)
            
            # Определяем диапазон страниц для поиска
            max_page = len(pdf)
            
            for i in range(3, max_page):
                page = pdf.get_page(i)
                textpage = page.get_textpage()
                text = textpage.get_text_bounded()
                
                if re.findall(pattern, text):
                    pages.append(i)
                    
                textpage.close()
                page.close()
                
            pdf.close()
            
        self.pages = pages
        return pages

    def look_for_objects(self):
        """
        Обнаруживает объекты на схемах с использованием модели YOLO.
        
        Returns:
            list: Результаты детекции объектов от модели YOLO
        """
        model = YOLO('best_v5.pt')
        results = []
        pdf = pdfium.PdfDocument(self.filePath)
        ensure_output_images_dir()
        
        for page_number in self.pages:
            page = pdf.get_page(page_number)
            image = page.render(scale=2).to_pil()
            image_path = os.path.join(OUTPUT_IMAGES_DIR, f"output_{page_number:03d}.png")
            image.save(image_path)
            results.append(model.predict(image_path, imgsz=2560))
            page.close()
            
        pdf.close()
        print(results)
        return results

    def get_coords(self, results):
        """
        Извлекает координаты обнаруженных объектов из результатов модели YOLO.
        
        Args:
            results: Результаты детекции объектов от модели YOLO
            
        Returns:
            list: Список словарей с координатами объектов для каждой страницы
        """
        coords = []
        page_index = 0
        
        for page_results in results:
            for result in page_results:
                coords.append({})
                boxes = result.boxes
                
                for box in boxes:
                    x1, y1, x2, y2 = box.xyxyn[0]
                    # Преобразование координат (отсчет от левого верхнего угла)
                    x1n = x1
                    x2n = x2
                    y1n = (1 - y1)
                    y2n = (1 - y2)
                    
                    class_id = int(box.cls[0].item())
                    class_name = result.names.get(class_id) if hasattr(result, "names") else None
                    
                    if class_id in [0, 1, 2] or class_name == "motor_operated_valve":
                        coords[page_index][(x1n, y1n, x2n, y2n)] = "Нечто"
                        
                page_index += 1
                
        print(coords)
        return coords


    def get_KKS(self, coords):
        """
        Извлекает коды KKS из обнаруженных областей на схемах.
        
        Args:
            coords: Список словарей с координатами объектов
            
        Returns:
            dict: Словарь с кодами KKS и их типами
        """
        dict_list = dict()
        pdf = pdfium.PdfDocument(self.filePath)
        page_index = 0
        
        for imp_page in self.pages:
            page = pdf.get_page(imp_page)
            textpage = page.get_textpage()
            bbox = page.get_bbox()
            print(bbox)
            
            for coord_key in coords[page_index].keys():
                x1, y1, x2, y2 = coord_key
                # Преобразование нормализованных координат в абсолютные
                x1_abs = x1 * bbox[2]
                x2_abs = x2 * bbox[2]
                y1_abs = y1 * bbox[3]
                y2_abs = y2 * bbox[3]
                
                extracted_text = textpage.get_text_bounded(x1_abs, y1_abs, x2_abs, y2_abs)
                # Удаление переносов строк
                cleaned_text = "".join(extracted_text.split("\r"))
                cleaned_text = "".join(cleaned_text.split("\n"))
                dict_list[cleaned_text] = coords[page_index][coord_key]
                
            page_index += 1
            textpage.close()
            page.close()
            
        pdf.close()
        # Извлечение кодов KKS с помощью регулярных выражений
        fin_dict = dict()
        # pattern1 = r"(AA|CT|CL|CF|CP|CY|CS|CQ|CU|CM|CR)(\d{3})([A-D]?)(\d{2})(\w{3})(\d{2})"
        # pattern2 = r"(\d{2})(\w{3})(\d{2})(AA|CT|CL|CF|CP|CY|CS|CQ|CU|CM|CR)(\d{3})([A-D]?)"

        pattern1 = r"(AA|CT|CL|CF|CP|CY|CS|CQ|CU|CM|CR)(\d{3})([A-D]?)\s?(\d{2})(\w{3})(\d{2})"
        pattern2 = r"(\d{2})(\w{3})(\d{2})\s?(AA|CT|CL|CF|CP|CY|CS|CQ|CU|CM|CR)(\d{3})([A-D]?)"
        
        for text in dict_list.keys():
            print(text)
            matches1 = re.findall(pattern1, text)
            matches2 = re.findall(pattern2, text)
            
            # Формирование кодов KKS
            kks_codes = [f"{match[3]}{match[4]}{match[5]}{match[0]}{match[1]}{match[2]}" 
                        for match in matches1]
            kks_codes.extend([f"{match[0]}{match[1]}{match[2]}{match[3]}{match[4]}{match[5]}" 
                             for match in matches2])
            
            # Определение типа оборудования
            for kks_code in kks_codes:
                if "AA" in kks_code:
                    fin_dict[kks_code] = "Арматура"
                elif "AP" in kks_code:
                    fin_dict[kks_code] = "Насос"
                else:
                    fin_dict[kks_code] = "Датчик"
                    
        print(fin_dict)
        return fin_dict



    def write_excel(self, dict_list, output_filename: str, sheet_name='Sheet1'):
        """
        Сохраняет извлеченные коды KKS в Excel файл.
        
        Args:
            dict_list: Словарь с кодами KKS и их типами
            output_filename: Путь к выходному файлу
            sheet_name: Имя листа в Excel
            
        Returns:
            str: Путь к созданному файлу
        """
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        ws.append(["KKS", "Object"])

        for key, value in dict_list.items():
            ws.append([key, value])

        wb.save(output_filename)
        return output_filename

    def process_directory(self, files_path, xls_path):
        """
        Обрабатывает все PDF файлы в указанной директории.
        
        Args:
            files_path: Путь к папке с PDF файлами
            xls_path: Путь к папке для сохранения результатов
        """
        dir_list = os.listdir(files_path)
        
        for filename in dir_list:
            if not filename.lower().endswith('.pdf'):
                continue
                
            filepath = os.path.abspath(os.path.join(files_path, filename))
            kks = filename.replace('.pdf', '')

            obj_finder = my_obj_finder(filepath)
            obj_finder.find_imp_pages()
            results = obj_finder.look_for_objects()
            coords = obj_finder.get_coords(results)
            kks_dict = obj_finder.get_KKS(coords)
            obj_finder.write_excel(kks_dict, os.path.join(xls_path, f"{kks}.xlsx"))

