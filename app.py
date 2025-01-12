#!/usr/bin/env python 3

import pandas as pd
from datetime import datetime, timedelta
import logging
from typing import Union, Optional
from openpyxl import load_workbook
import argparse

pd.set_option('future.no_silent_downcasting', True)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

DEFAULT_EXCEL_FILE = 'Черненко Александр Александрович.xlsx'
BASE_DATE = datetime(2024, 10, 3)

def get_sheet_names(excel_file: str) -> list:
    """Получает список имен всех листов в Excel файле."""
    workbook = load_workbook(excel_file, read_only=True)
    return workbook.sheetnames

def choose_sheet(excel_file: str, sheet_index: int) -> str:
    """Позволяет пользователю выбрать лист для работы."""
    sheet_names = get_sheet_names(excel_file)
    if 0 <= sheet_index < len(sheet_names):
        return sheet_names[sheet_index]
    else:
        raise ValueError("Неверный номер листа.")

def parse_date(date: Union[datetime, str, float]) -> Optional[datetime]:
    """Преобразует различные форматы дат в объект datetime."""
    if isinstance(date, (datetime, str)):
        return pd.to_datetime(date)
    number = int(float(str(date)))
    return BASE_DATE + timedelta(days=(number - 1) * 7)

def save_report(content: str, file_name: str) -> None:
    """Сохраняет отчет в HTML файл."""
    with open(file_name, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"Отчет успешно сохранен в файл: {file_name}")

def generate_attendance_report(student_data: pd.DataFrame, dates: list, date: datetime, group_name: str) -> None:
    """Генерирует отчет о посещаемости за указанную дату и предлагает сохранить его."""
    date_str = date.strftime('%Y-%m-%d')
    if date_str in [d.strftime('%Y-%m-%d') for d in dates]:
        index = dates.index(date)
        attendance_column = student_data.columns[4 + index]
        
        attendance_data = student_data[['Фамилия', 'Имя', 'Отчество', attendance_column]].copy()
        attendance_data[attendance_column] = attendance_data[attendance_column].fillna(0).astype(int)
        
        report_content = f"<h1>Отчет по группе за дату</h1>\n<h2>Группа: {group_name}</h2>\n<h3>Всего студентов: {len(student_data)}</h3>\n"
        report_content += attendance_data.to_html(index=False)
        
        save_report(report_content, f"отчет_группы_{group_name}_за_{date_str}.html")
    else:
        print(f"Нет занятий на дату {date_str}.")

def view_all_data(student_data: pd.DataFrame, dates: list, group_name: str) -> None:
    """Просмотр всех данных файла и возможность сохранения в отчет."""
    report_content = f"<h1>Отчет по группе</h1>\n<h2>Группа: {group_name}</h2>\n<h3>Всего студентов: {len(student_data)}</h3>\n"
    report_content += student_data.to_html(index=False)
    
    save_report(report_content, f"полный_отчет_группы_{group_name}.html")

def generate_student_report(student_data: pd.DataFrame, dates: list, group_name: str, student_index: int) -> None:
    """Генерирует отчет о посещаемости для выбранного студента."""
    sorted_student_data = student_data.sort_values('№ п/п')
    if 0 <= student_index < len(sorted_student_data):
        student = sorted_student_data.iloc[student_index]
    else:
        raise ValueError("Неверный номер студента.")
    
    attendance_data = pd.DataFrame({
        'Дата': dates,
        'Присутствие': student.iloc[4:].values
    })
    attendance_data['Присутствие'] = attendance_data['Присутствие'].fillna(0).astype(int)
    
    fio = f"{student['Фамилия']} {student['Имя']} {student['Отчество']}"
    attendance_count = attendance_data['Присутствие'].sum()
    total_classes = len(dates)
    attendance_percentage = (attendance_count / total_classes) * 100
    
    report_content = f"<h1>Отчет по студенту</h1>\n<h2>Группа: {group_name}</h2>\n<h3>Всего студентов: {len(student_data)}</h3>\n<h3>ФИО: {fio}</h3>\n<h3>Посещаемость: Посещено занятий {attendance_count} / всего занятий {total_classes} -- {attendance_percentage:.2f}%</h3>\n"
    report_content += attendance_data.to_html(index=False)
    
    save_report(report_content, f"отчет_о_посещаемости_студента(ки)_{student['Фамилия']}_{group_name}.html")

def main(excel_file: str, sheet_index: int, action: str, student_index: Optional[int] = None, date_index: Optional[int] = None) -> None:
    """Основная функция программы для обработки данных из Excel и взаимодействия с пользователем."""   
    sheet_name = choose_sheet(excel_file, sheet_index)
    
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
    except Exception as e:
        print(f"Ошибка при чтении файла Excel: {e}")
        return

    group_name = df.iloc[0, 0].split(':')[-1].strip()
    dates = df.iloc[2, 4:].dropna().apply(parse_date).tolist()
    headers = ['№ п/п', 'Фамилия', 'Имя', 'Отчество'] + [f'Занятие_{i+1}' for i in range(len(dates))]
    student_data = df.iloc[3:].reset_index(drop=True)
    student_data.columns = headers
    student_data.iloc[:, 4:] = student_data.iloc[:, 4:].fillna(0).astype(int)

    if action == '1':
        if student_index is not None:
            generate_student_report(student_data, dates, group_name, student_index)
        else:
            print("Необходимо указать номер студента для отчета.")
    elif action == '2':
        if date_index is not None:
            if 0 <= date_index < len(dates):
                generate_attendance_report(student_data, dates, dates[date_index], group_name)
            else:
                print("Неверный номер даты.")
        else:
            print("Необходимо указать номер даты для отчета.")
    elif action == '3':
        view_all_data(student_data, dates, group_name)
    elif action == '4':
        print("Выход из программы.")
    else:
        print("Неверный выбор. Пожалуйста, попробуйте снова.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Программа для работы с данными о посещаемости студентов.")
    parser.add_argument("--file", default=DEFAULT_EXCEL_FILE, help="Путь к файлу Excel")
    parser.add_argument("--sheet_index", type=int, required=True, help="Номер листа для работы")
    parser.add_argument("--action", required=True, help="Действие: 1 - отчет по студентам, 2 - отчет по дате, 3 - просмотр всех данных, 4 - выход")
    parser.add_argument("--student_index", type=int, help="Номер студента для отчета (только для действия 1)")
    parser.add_argument("--date_index", type=int, help="Номер даты для отчета (только для действия 2)")
    args = parser.parse_args()

    main(args.file, args.sheet_index, args.action, args.student_index, args.date_index)