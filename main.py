import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
import argparse
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from datetime import datetime
import zipfile
import shutil

def prettify(elem):
    """Повертає відформатований XML рядок для запису у файл без XML декларації."""
    rough_string = ET.tostring(elem, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    xml_str = reparsed.toprettyxml(indent="  ")
    # Видаляємо перший рядок (XML declaration), якщо він присутній
    if xml_str.startswith('<?xml'):
        lines = xml_str.splitlines()
        if lines:
            return "\n".join(lines[1:])
    return xml_str

def convert_logic(xlsx_path, output_zip_path, employer_edrpou, employer_name, log_callback=print):
    """Основна логіка конвертації XLSX в XML з подальшим пакуванням у ZIP."""
    try:
        xls = pd.ExcelFile(xlsx_path)
        if 'employees' not in xls.sheet_names:
            log_callback(f"Помилка: Лист 'employees' не знайдено у файлі {xlsx_path}")
            return False

        df = pd.read_excel(xls, sheet_name='employees')
        
        # Видаляємо порожні рядки, де немає Прізвища або Імені
        df = df.dropna(subset=['Прізвище', "Ім'я"])

        if df.empty:
            log_callback("Попередження: Не знайдено даних у листі 'employees'.")
            return False
        
        # Створення кореневого елемента XML для всіх працівників
        root = ET.Element("DOCUMENT")
        
        # Блок Страхувальника (один для всіх)
        employer = ET.SubElement(root, "EMPLOYER")
        ET.SubElement(employer, "EDRPOU").text = str(employer_edrpou)
        ET.SubElement(employer, "NAME").text = str(employer_name)
        
        # Контейнер для всіх працівників
        individuals_root = ET.SubElement(root, "INDIVIDUALS")
        
        # Групування за фізичною особою
        grouped = df.groupby(['РНОКПП', 'Прізвище', "Ім'я", 'По батькові'], sort=False)
        
        for (raw_rnokpp, surname, name, patronymic), group_df in grouped:
            # Створення елемента для кожного працівника
            individual = ET.SubElement(individuals_root, "INDIVIDUAL")
            
            # Логіка: 10 цифр -> RNOKPP, інакше -> SERIA_NUMBER (паспорт)
            rnokpp_val = ""
            seria_number_val = ""
            
            if pd.notna(raw_rnokpp):
                try:
                    cleaned_val = str(int(float(raw_rnokpp)))
                except:
                    cleaned_val = str(raw_rnokpp).strip()
                
                if cleaned_val.isdigit() and len(cleaned_val) == 10:
                    rnokpp_val = cleaned_val
                else:
                    seria_number_val = cleaned_val
            
            ET.SubElement(individual, "RNOKPP").text = rnokpp_val
            ET.SubElement(individual, "SERIA_NUMBER").text = seria_number_val
            ET.SubElement(individual, "SURNAME").text = str(surname)
            ET.SubElement(individual, "NAME").text = str(name)
            ET.SubElement(individual, "PATRONYMIC").text = str(patronymic) if pd.notna(patronymic) else ""
            
            records_root = ET.SubElement(individual, "RECORDS")
            
            for _, row in group_df.iterrows():
                record = ET.SubElement(records_root, "RECORD")
                
                # Номер запису
                emp_code = "1"
                if pd.notna(row['номер місця трудових відносин']):
                    try:
                        emp_code = str(int(float(row['номер місця трудових відносин'])))
                    except:
                        emp_code = str(row['номер місця трудових відносин'])
                ET.SubElement(record, "EMPLOYER_CODE").text = emp_code
                
                # Дані про страхувальника (SIGN та LR коди)
                ET.SubElement(record, "EDRPO_SIGN").text = str(employer_edrpou)
                ET.SubElement(record, "NAME_SIGN").text = str(employer_name)
                ET.SubElement(record, "EDRPO_LR").text = str(employer_edrpou)
                ET.SubElement(record, "NAME_LR").text = str(employer_name)

                ET.SubElement(record, "ACTION_TYPE").text = str(row['Тип події']).split(' — ')[0]
                ET.SubElement(record, "ATTRIBUTE_TYPE").text = str(row['Атрибут події']).split(' — ')[0]
                
                # Форматування дати події
                val = row['Дата події']
                if pd.notna(val):
                    if isinstance(val, pd.Timestamp):
                        formatted_date = val.strftime('%Y-%m-%d')
                    elif hasattr(val, 'strftime'):
                        formatted_date = val.strftime('%Y-%m-%d')
                    else:
                        try:
                            d = pd.to_datetime(val, dayfirst=True)
                            formatted_date = d.strftime('%Y-%m-%d')
                        except:
                            formatted_date = str(val).split(' ')[0]
                    ET.SubElement(record, "ACTION_DT").text = formatted_date
                else:
                    ET.SubElement(record, "ACTION_DT").text = ""

                ET.SubElement(record, "ACTION_TEXT").text = str(row['Текст запису']) if pd.notna(row['Текст запису']) else ""
                
                # Причина звільнення (обов'язкова для коду 3)
                attr_code = str(row['Атрибут події']).split(' — ')[0]
                leave_reason = ""
                if attr_code == "3":
                    leave_reason = str(row['Текст запису']) if pd.notna(row['Текст запису']) else ""
                
                ET.SubElement(record, "LEAVE_REASON").text = leave_reason
                ET.SubElement(record, "DOC_TYPE").text = "Наказ"
                
                # DOC_DT має бути після DOC_TYPE
                val = row['Дата документа підстави']
                if pd.notna(val):
                    if isinstance(val, pd.Timestamp):
                        formatted_date = val.strftime('%Y-%m-%d')
                    elif hasattr(val, 'strftime'):
                        formatted_date = val.strftime('%Y-%m-%d')
                    else:
                        try:
                            d = pd.to_datetime(val, dayfirst=True)
                            formatted_date = d.strftime('%Y-%m-%d')
                        except:
                            formatted_date = str(val).split(' ')[0]
                    ET.SubElement(record, "DOC_DT").text = formatted_date
                else:
                    ET.SubElement(record, "DOC_DT").text = ""
                
                doc_number = str(row['Номер документу підстави']) if pd.notna(row['Номер документу підстави']) else ""
                if doc_number.isdigit():
                    doc_number = f"{doc_number} - к/п"
                ET.SubElement(record, "DOC_NUMBER").text = doc_number
            
            log_callback(f"Додано працівника: {surname} {name}")
        
        # Формування назви файлу на основі оригінального файлу
        base_name = os.path.splitext(os.path.basename(xlsx_path))[0]
        xml_filename = f"{base_name}.xml"
        
        # Збереження XML у ZIP
        xml_str = prettify(root)
        with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            zipf.writestr(xml_filename, xml_str)
        
        log_callback(f"Успішно: {len(grouped)} працівників збережено в {xml_filename}")
        log_callback(f"Архів створено: {output_zip_path}")
        return True

    except Exception as e:
        log_callback(f"Сталася помилка: {e}")
        return False

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Конвертер ПФУ (XLSX -> XML)")
        self.root.geometry("600x650")
        
        # Налаштування стилю
        main_frame = tk.Frame(root, padx=20, pady=20)
        main_frame.pack(expand=True, fill="both")
        
        # Заголовок
        header = tk.Label(main_frame, text="Конвертер ПФУ (XML)", font=("Helvetica", 20, "bold"), fg="#2c3e50")
        header.pack(pady=(0, 20))
        
        # Блок даних підприємства
        emp_frame = tk.LabelFrame(main_frame, text=" Дані про підприємство ", padx=10, pady=10)
        emp_frame.pack(fill="x", pady=(0, 15))
        
        tk.Label(emp_frame, text="ЄДРПОУ:").grid(row=0, column=0, sticky="w", pady=5)
        self.edrpou_entry = tk.Entry(emp_frame, width=30)
        self.edrpou_entry.grid(row=0, column=1, padx=10, pady=5)
        
        tk.Label(emp_frame, text="Назва:").grid(row=1, column=0, sticky="w", pady=5)
        self.name_entry = tk.Entry(emp_frame, width=50)
        self.name_entry.grid(row=1, column=1, padx=10, pady=5)
        
        # Блок вибору файлу та шаблону
        file_frame = tk.LabelFrame(main_frame, text=" Робота з файлами ", padx=10, pady=10)
        file_frame.pack(fill="x", pady=(0, 20))
        
        self.file_path_var = tk.StringVar(value="Файл не обрано")
        
        btn_frame = tk.Frame(file_frame)
        btn_frame.pack(fill="x", pady=5)
        
        tk.Button(btn_frame, text="Обрати XLSX", command=self.select_file, width=15).pack(side="left", padx=(0, 10))
        tk.Button(btn_frame, text="Зберегти шаблон", command=self.save_template, width=15, bg="#ecf0f1").pack(side="left")
        
        tk.Label(file_frame, textvariable=self.file_path_var, fg="#7f8c8d", wraplength=500, justify="left").pack(fill="x", pady=5)
        
        # Кнопка запуску
        self.run_btn = tk.Button(
            main_frame, 
            text="КОНВЕРТУВАТИ В XML", 
            command=self.start_conversion,
            bg="#3498db", 
            fg="white", 
            font=("Helvetica", 12, "bold"),
            height=2
        )
        self.run_btn.pack(fill="x", pady=(0, 15))
        
        # Лог
        tk.Label(main_frame, text="Лог виконання:").pack(anchor="w")
        self.log_area = scrolledtext.ScrolledText(main_frame, height=10, font=("Consolas", 9))
        self.log_area.pack(expand=True, fill="both")
        self.log_area.config(state="disabled")

    def log(self, text):
        self.log_area.config(state="normal")
        self.log_area.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {text}\n")
        self.log_area.see(tk.END)
        self.log_area.config(state="disabled")
        self.root.update_idletasks()

    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.file_path_var.set(path)
            self.log(f"Обрано файл: {os.path.basename(path)}")

    def save_template(self):
        template_name = "Шаблон ПФУ.xlsx"
        # Шукаємо шаблон поруч зі скриптом або в робочій директорії
        possible_paths = [
            template_name,
            os.path.join(os.path.dirname(os.path.abspath(__file__)), template_name)
        ]
        
        template_source = None
        for p in possible_paths:
            if os.path.exists(p):
                template_source = p
                break
        
        if not template_source:
            messagebox.showerror("Помилка", f"Файл '{template_name}' не знайдено!")
            return
            
        dest_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=template_name,
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if dest_path:
            try:
                shutil.copy2(template_source, dest_path)
                self.log(f"Шаблон збережено у: {dest_path}")
                messagebox.showinfo("Успіх", "Шаблон успішно збережено!")
            except Exception as e:
                messagebox.showerror("Помилка", f"Не вдалося зберегти шаблон: {e}")

    def start_conversion(self):
        edrpou = self.edrpou_entry.get().strip() if self.edrpou_entry.get().strip() else "02125639"
        name = self.name_entry.get().strip() if self.name_entry.get().strip() else "ДО УДПУ ІМЕНІ ПАВЛА ТИЧИНИ"
        xlsx_path = self.file_path_var.get()
        
        if not edrpou or not name:
            messagebox.showerror("Помилка", "Введіть ЄДРПОУ та назву підприємства!")
            return
        
        if xlsx_path == "Файл не обрано":
            messagebox.showerror("Помилка", "Будь ласка, оберіть файл XLSX!")
            return
            
        output_zip_path = os.path.splitext(xlsx_path)[0] + ".zip"
        
        self.log("Початок конвертації та архівації...")
        success = convert_logic(xlsx_path, output_zip_path, edrpou, name, self.log)
        
        if success:
            messagebox.showinfo("Готово", f"XML архів успішно створено:\n{output_zip_path}")
        else:
            messagebox.showerror("Помилка", "Під час конвертації сталася помилка. Перевірте лог.")

if __name__ == "__main__":
    # Підтримка командного рядка
    if len(sys.argv) > 1 and "--edrpou" in sys.argv:
        parser = argparse.ArgumentParser(description="Конвертер XLSX в XML (CLI)")
        parser.add_argument("--input", default="Шаблон ПФУ.xlsx")
        parser.add_argument("--output", default="output.zip")
        parser.add_argument("--edrpou", required=True)
        parser.add_argument("--name", required=True)
        args = parser.parse_args()
        convert_logic(args.input, args.output, args.edrpou, args.name)
    else:
        # Запуск GUI на Tkinter
        root = tk.Tk()
        app = App(root)
        root.mainloop()
