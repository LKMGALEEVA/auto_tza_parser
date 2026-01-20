import pandas as pd
import re
import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

class mTP:
    def parse_two_tables(self, file1, file2):
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)
        #df2 = df2.dropna()

        # Получаем значения из первого столбца
        col1 = df1.iloc[:, 0]
        classes = df1.iloc[:, 1] if df1.shape[1] > 1 else pd.Series([""] * len(df1))
        col2 = df2.iloc[3:, 0]

        objects_map = {}
        for raw_value, raw_class in zip(col1, classes):
            if pd.isna(raw_value):
                continue
            value = str(raw_value).strip()
            if not value or value.lower() == "nan":
                continue
            category = self._category_from_raw(raw_class) or "Контроль"
            objects_map[value] = category

        results = []
        used_keys = set()
        table_category = self._category_from_filename(file2)

        for raw_value in col2:
            if pd.isna(raw_value):
                continue
            value = str(raw_value).strip()
            if not value or value.lower() == "nan":
                continue

            matched_keys = []
            for key, category in objects_map.items():
                if re.match(r'^' + re.escape(key) + r'\b', value):
                    matched_keys.append((key, category))

            if matched_keys:
                for key, category in matched_keys:
                    results.append(self._build_result_row(key, "Совпадает", "green", category))
                    used_keys.add(key)
            else:
                results.append(self._build_result_row(value, "Есть в таблице, нет на схеме", "red", table_category))

        for key, category in objects_map.items():
            if key not in used_keys:
                results.append(self._build_result_row(key, "Есть на схеме, нет в таблице", "red", category))

        return results

    def _write_results_to_sheet(self, ws, results):
        headers = ["Значение", "Статус"]
        for col_num, data in enumerate(headers, 1):
            ws.cell(row=1, column=col_num, value=data)

        for row_num, row in enumerate(results, 2):
            ws.cell(row=row_num, column=1, value=row["value"])
            ws.cell(row=row_num, column=2, value=row["status"])

        color_map = {
            "green": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
            "red": PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        }

        for row_num, row in enumerate(results, 2):
            fill_style = color_map.get(row["color"], color_map["red"])
            ws.cell(row=row_num, column=1).fill = fill_style
            ws.cell(row=row_num, column=2).fill = fill_style

        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column].width = adjusted_width

    def _deduplicate_results(self, results):
        unique_results = []
        seen = set()
        for row in results:
            row_key = (row["value"], row["status"], row.get("category"))
            if row_key not in seen:
                seen.add(row_key)
                unique_results.append(row)
        return unique_results

    def _category_from_raw(self, raw_class):
        if pd.isna(raw_class):
            return None
        text = str(raw_class).strip().lower()
        if not text or text == "nan":
            return None
        if "армат" in text:
            return "Арматура"
        return "Контроль"

    def _category_from_filename(self, file_path):
        name = os.path.basename(file_path).lower()
        if ("арматур" in name) or ("регулиру" in name):
            return "Арматура"
        return "Контроль"

    def _build_result_row(self, value, status, color, category):
        return {
            "value": value,
            "status": status,
            "color": color,
            "category": category or "Контроль"
        }

    def parse_all_related(self, kks, file1, search_folder):
        # Список для хранения совпадений
        matching_files = {"Арматура": [], "Контроль": []}
        # Перебираем все файлы в указанной папке
        for file in os.listdir(search_folder):
            # Полный путь к файлу
            full_path = os.path.join(search_folder, file)
            # Проверяем, является ли это файлом и содержит ли имя нужное слово
            if not os.path.isfile(full_path) or kks not in file:
                continue
            if ("Запорная арматура" in file) or ("Регулирующая арматура" in file):
                matching_files["Арматура"].append(full_path)
            elif ("Точки контроля" in file) or ("Точки теплотехнического контроля" in file):
                matching_files["Контроль"].append(full_path)

        #создаем эксельку
        wb = Workbook()
        ws_valves = wb.active
        ws_valves.title = "Арматура"
        ws_control = wb.create_sheet(title="Контроль")

        for sheet_name, files in matching_files.items():
            ws = ws_valves if sheet_name == "Арматура" else ws_control
            collected_results = []
            for matching_file in files:
                collected_results.extend(self.parse_two_tables(file1, matching_file))
            filtered_results = [
                row for row in collected_results
                if (sheet_name == "Арматура" and row.get("category") == "Арматура")
                or (sheet_name == "Контроль" and row.get("category") != "Арматура")
            ]
            unique_results = self._deduplicate_results(filtered_results)
            
            # Проверка дубликатов: если значение повторяется, оставляем только строку со статусом "Совпадает"
            value_groups = {}
            for row in unique_results:
                value = row["value"]
                if value not in value_groups:
                    value_groups[value] = []
                value_groups[value].append(row)
            
            final_results = []
            for value, rows in value_groups.items():
                if len(rows) > 1:
                    # Если есть несколько строк с одним значением, ищем строку со статусом "Совпадает"
                    matching_row = next((row for row in rows if row.get("status") == "Совпадает"), None)
                    if matching_row:
                        final_results.append(matching_row)
                    else:
                        # Если нет строки "Совпадает", оставляем все строки
                        final_results.extend(rows)
                else:
                    # Если только одна строка с этим значением, добавляем её
                    final_results.extend(rows)
            
            self._write_results_to_sheet(ws, final_results)

        # Удаление строк со значением "Краткое наименование" в первом столбце на каждой вкладке
        for ws in [ws_valves, ws_control]:
            rows_to_delete = []
            for row_idx in range(1, ws.max_row + 1):
                cell_value = ws.cell(row=row_idx, column=1).value
                if cell_value and str(cell_value).strip() == "Краткое наименование":
                    rows_to_delete.append(row_idx)
            # Удаляем строки в обратном порядке, чтобы не сбивать индексацию
            for row_idx in reversed(rows_to_delete):
                ws.delete_rows(row_idx, 1)

        parent_dir = os.path.dirname(search_folder)
        wb.save(f'{parent_dir}/{kks}.xlsx')
        print(f"✅ Результаты сохранены в '{kks}.xlsx'")



if __name__ == "__main__":
    files_path = "C:/Users/User/Desktop/best_new_new/SCHEMES"
    folder = "C:/Users/User/Desktop/best_new_new/TABLES"
    dir_list = os.listdir(files_path)
    for f in dir_list:
        mtp = mTP()
        kks = f.strip('.xlsx')
        mtp.parse_all_related(kks, f"{files_path}/{f}", folder)