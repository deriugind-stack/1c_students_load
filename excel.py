import os
import sys
import subprocess
import importlib
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import logging
from collections import defaultdict
import warnings

# Подавляем предупреждения от openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Настройка логирования
logging.basicConfig(
    filename="process.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    encoding="utf-8"
)

def ensure_package(package_name):
    try:
        importlib.import_module(package_name)
        logging.info(f"Библиотека {package_name} найдена.")
    except ImportError:
        logging.warning(f"Библиотека {package_name} не найдена. Устанавливаю...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
        logging.info(f"Библиотека {package_name} установлена.")

def load_excel_files():
    print("Выберите режим:")
    print("1 - Загрузить один Excel файл")
    print("2 - Загрузить несколько Excel файлов")

    choice = input("Ваш выбор (1/2): ").strip()
    files = []
    root = tk.Tk()
    root.withdraw()

    if choice == "1":
        file_path = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path and os.path.exists(file_path):
            files.append(file_path)
    elif choice == "2":
        file_paths = filedialog.askopenfilenames(
            title="Выберите Excel файлы",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        for path in file_paths:
            if os.path.exists(path):
                files.append(path)
    else:
        print("Неверный выбор.")

    logging.info(f"Выбрано файлов: {len(files)}")
    return files

def split_fio(fio):
    parts = str(fio).strip().split()
    while len(parts) < 3:
        parts.append("")
    return parts[0], parts[1], parts[2]

def extract_date(text):
    import re
    match = re.search(r"\d{2}\.\d{2}\.\d{4}", str(text))
    return match.group(0) if match else ""

def safe_value(val):
    """Заменяем NaN на пустую строку и приводим к строке."""
    return "" if pd.isna(val) else str(val).strip()

def detect_gender_by_relation(relation):
    """Определение пола по степени родства."""
    relation = str(relation).strip().lower()
    female_relations = {"мать", "бабушка", "сестра", "тётя"}
    male_relations = {"отец", "дедушка", "брат", "дядя"}

    if relation in female_relations:
        return "Ж"
    if relation in male_relations:
        return "М"
    return ""  # если определить не удалось

def create_empty_target(file_path):
    columns = [
        "Фамилия ученика","Имя ученика","Отчество ученика","Пол ученика",
        "Дата рождения ученика","Номер личного дела","Дата приема в школу",
        "Класс","Дата зачисления/перевода","Номер приказа","Дата приказа",
        "Фамилия Родителя 1","Имя Родителя 1","Отчество Родителя 1","Пол Родителя 1",
        "Дата рождения Родителя 1","Степень родства","Телефон домашний","Телефон мобильный",
        "E-mail Родителя 1","Фамилия Родителя 2","Имя Родителя 2","Отчество Родителя 2",
        "Пол Родителя 2","Дата рождения Родителя 2","Степень родства Родителя 2",
        "Телефон домашний Родителя 2","Телефон мобильный Родителя 2","E-mail Родителя 2"
    ]
    pd.DataFrame(columns=columns).to_excel(file_path, index=False, engine="openpyxl")
    logging.info(f"Создан новый целевой файл: {file_path}")

def process_files(files, target_file):
    merged_rows = {}
    class_counts = defaultdict(int)
    class_students = defaultdict(list)

    header_keywords = [
        "ФИО", "Пол", "Личное дело №", "ФИО представителя",
        "Тип представителя", "Телефон представителя"
    ]

    def is_header(row):
        return any(str(cell).strip() in header_keywords for cell in row)

    for file in files:
        logging.info(f"Начата обработка файла: {file}")
        try:
            if file.lower().endswith(".xls"):
                df = pd.read_excel(file, engine="xlrd")
            else:
                df = pd.read_excel(file, engine="openpyxl")

            df = df.dropna(how="all")
            valid_rows = df[df.count(axis=1) >= 5]
            valid_rows = valid_rows[~valid_rows.apply(is_header, axis=1)]

            class_name = os.path.splitext(os.path.basename(file))[0]

            for _, row in valid_rows.iterrows():
                fam_u, name_u, otch_u = split_fio(row.iloc[2])
                fio_key = f"{fam_u} {name_u} {otch_u}"

                fam_r, name_r, otch_r = split_fio(row.iloc[8])
                tip_rodstva = safe_value(row.iloc[6])
                gender_r1 = detect_gender_by_relation(tip_rodstva)

                pol_uch = safe_value(row.iloc[3])
                date_uch = extract_date(row.iloc[4])
                lich_delo = safe_value(row.iloc[0])
                date_priem = extract_date(row.iloc[1])

                tel_mob_parent = safe_value(row.iloc[12])
                email_parent = safe_value(row.iloc[13])
                date_parent_birth = extract_date(row.iloc[9])

                if fio_key not in merged_rows:
                    merged_rows[fio_key] = [
                        fam_u, name_u, otch_u,
                        pol_uch,
                        date_uch,
                        lich_delo,
                        date_priem,
                        class_name,
                        "", "", "",
                        fam_r, name_r, otch_r,
                        gender_r1,
                        date_parent_birth,
                        tip_rodstva,
                        "",
                        tel_mob_parent,
                        email_parent,
                        "", "", "", "", "", "", "", "", "", ""
                    ]
                    class_counts[class_name] += 1
                    class_students[class_name].append(fio_key)
                    logging.info(f"Создана запись для ученика {fio_key} ({class_name}) с Родителем 1")
                else:
                    if merged_rows[fio_key][20] == "":
                        tip_rodstva2 = safe_value(row.iloc[6])
                        gender_r2 = detect_gender_by_relation(tip_rodstva2)
                        email_parent2 = safe_value(row.iloc[13])
                        tel_mob_parent2 = safe_value(row.iloc[12])
                        date_parent_birth2 = extract_date(row.iloc[9])

                        merged_rows[fio_key][20] = fam_r
                        merged_rows[fio_key][21] = name_r
                        merged_rows[fio_key][22] = otch_r
                        merged_rows[fio_key][23] = gender_r2
                        merged_rows[fio_key][24] = date_parent_birth2
                        merged_rows[fio_key][25] = tip_rodstva2
                        merged_rows[fio_key][26] = ""
                        merged_rows[fio_key][27] = tel_mob_parent2
                        merged_rows[fio_key][28] = email_parent2

                        logging.info(f"Добавлен Родитель 2 для ученика {fio_key} ({class_name})")
                    else:
                        logging.warning(f"Ученик {fio_key} имеет более двух родителей. Лишние данные проигнорированы.")

            logging.info(f"Добавлено {len(valid_rows)} строк из файла {file}")

        except Exception as e:
            logging.error(f"Ошибка при обработке {file}: {e}")

    if target_file.lower().endswith(".xls"):
        target_file = target_file[:-4] + ".xlsx"

    if not os.path.exists(target_file):
        create_empty_target(target_file)

    target_df = pd.read_excel(target_file, engine="openpyxl")
    last_index = len(target_df)

    for row in merged_rows.values():
        row = ["" if pd.isna(x) else x for x in row]
        while len(row) < len(target_df.columns):
            row.append("")
        if len(row) > len(target_df.columns):
            row = row[:len(target_df.columns)]

        target_df.loc[last_index] = row
        last_index += 1

    target_df.to_excel(target_file, index=False, engine="openpyxl")

    total_students = len(merged_rows)
    logging.info(f"Итоговое количество уникальных учеников: {total_students}")
    for cls, count in class_counts.items():
        logging.info(f"Класс {cls}: {count} учеников")
        logging.info(f"Ученики класса {cls}: {', '.join(class_students[cls])}")

    # Выводим статистику в консоль
    print(f"\nВсе данные успешно добавлены в файл {target_file}.")
    print(f"Итоговое количество учеников: {total_students}")
    print("Статистика по классам:")
    for cls, count in class_counts.items():
        print(f"  {cls}: {count} учеников")
        print(f"    Ученики: {', '.join(class_students[cls])}")
    print("Подробный лог смотри в process.log")


# ВАЖНО: этот блок должен быть на нулевом уровне!
if __name__ == "__main__":
    ensure_package("xlrd")

    excel_files = load_excel_files()
    if excel_files:
        print("\nВыберите целевой Excel файл для записи:")
        root = tk.Tk()
        root.withdraw()
        target_file = filedialog.askopenfilename(
            title="Выберите целевой Excel файл",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )

        if target_file:
            process_files(excel_files, target_file)
        else:
            print("Целевой файл не выбран.")
    else:
        print("Нет файлов для обработки.")