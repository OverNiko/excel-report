import pandas as pd
from datetime import datetime, timedelta
import logging
from typing import Union, Optional
from tabulate import tabulate
import argparse

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

DEFAULT_EXCEL_FILE = 'Черненко Александр Александрович.xlsx'
DEFAULT_SHEET_NAME = 'Pyt-9'
BASE_DATE = datetime(2024, 10, 3)

def parse_date(date: Union[datetime, str, float]) -> Optional[datetime]:
    """Преобразует различные форматы дат в объект datetime."""
    if isinstance(date, (datetime, str)):
        return pd.to_datetime(date)
    number = int(float(str(date)))
    return BASE_DATE + timedelta(days=(number - 1) * 7)

def search_student(student_data: pd.DataFrame, last_name: str) -> pd.DataFrame:
    """Поиск студента по фамилии и вывод информации о найденных студентах."""
    result = student_data[student_data['Фамилия'].str.contains(last_name, case=False)]
    if result.empty:
        print(f"Студент с фамилией '{last_name}' не найден.")
    else:
        print("Найденные студенты:")
        print(tabulate(result.fillna(0), headers='keys', tablefmt='pretty', showindex=False))
    return result

def save_report(data: pd.DataFrame, file_name: str, group_name: str) -> None:
    """Сохраняет отчет в HTML файл с информацией о группе."""
    html_content = f"<h1>Отчет для группы: {group_name}</h1>\n"
    html_content += data.to_html(index=False)
    
    with open(file_name, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"Отчет успешно сохранен в файл: {file_name}")

def generate_attendance_report(student_data: pd.DataFrame, dates: list, date: datetime, group_name: str) -> None:
    """Генерирует отчет о посещаемости за указанную дату и предлагает сохранить его."""
    date_str = date.strftime('%Y-%m-%d')
    if date_str in [d.strftime('%Y-%m-%d') for d in dates]:
        index = dates.index(date)
        attendance_column = student_data.columns[4 + index]
        
        attendance_data = student_data[['№ п/п', 'Фамилия', 'Имя', 'Отчество', attendance_column]].copy()
        attendance_data[attendance_column] = attendance_data[attendance_column].fillna(0).astype(int)
        
        print(f"Отчет о посещаемости группы {group_name} за {date_str}:")
        print(tabulate(attendance_data, headers='keys', tablefmt='pretty', showindex=False))

        if input("Хотите сохранить отчет за текущую дату? (да/нет): ").strip().lower() == 'да':
            save_report(attendance_data, f"отчет_группы_{group_name}_за_{date_str}.html", group_name)
    else:
        print(f"Нет занятий на дату {date_str}.")

def view_all_data(student_data: pd.DataFrame, group_name: str) -> None:
    """Просмотр всех данных файла и возможность сохранения в отчет."""
    print(f"Все данные файла для группы {group_name}:")
    print(tabulate(student_data, headers='keys', tablefmt='pretty', showindex=False))
    
    if input("Хотите сохранить все данные в отчет? (да/нет): ").strip().lower() == 'да':
        save_report(student_data, f"полный_отчет_группы_{group_name}.html", group_name)

def generate_student_report(student_data: pd.DataFrame, student: pd.Series, dates: list, group_name: str) -> None:
    """Генерирует отчет о посещаемости для конкретного студента."""
    attendance_data = student[4:].reset_index()
    attendance_data.columns = ['Дата', 'Присутствие']
    attendance_data['Дата'] = dates
    attendance_data['Присутствие'] = attendance_data['Присутствие'].fillna(0).astype(int)
    
    print(f"Отчет о посещаемости студента(ки) {student['Фамилия']} {student['Имя']} {student['Отчество']}:")
    print(tabulate(attendance_data, headers='keys', tablefmt='pretty', showindex=False))
    
    if input("Хотите сохранить отчет для этого студента? (да/нет): ").strip().lower() == 'да':
        html_content = f"""
        <h1>Отчет о посещаемости для группы: {group_name}</h1>
        <h2>Студент: {student['Фамилия']} {student['Имя']} {student['Отчество']}</h2>
        {attendance_data.to_html(index=False)}
        """
        file_name = f"отчет_о_ посещаемости_студента(ки)_{student['Фамилия']}_{group_name}.html"
        with open(file_name, 'w', encoding='utf-8') as f:
            f.write(html_content)
        print(f"Отчет успешно сохранен в файл: {file_name}")

def main(excel_file: str, sheet_name: str) -> None:
    """Основная функция программы для обработки данных из Excel и взаимодействия с пользователем."""
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

    while True:
        print(f"\nРабота с данными группы: {group_name}")
        print("Выберите действие:")
        print("1. Поиск студента по фамилии")
        print("2. Генерация отчета о посещаемости за день/дату")
        print("3. Просмотр всех данных файла")
        print("4. Выход")
        
        choice = input("Введите номер действия: ")
        
        if choice == '1':
            last_name = input("Введите фамилию студента: ")
            result = search_student(student_data, last_name)
            if not result.empty:
                if len(result) == 1:
                    generate_student_report(student_data, result.iloc[0], dates, group_name)
                else:
                    student_index = int(input("Введите номер студента для подробного отчета: ")) - 1
                    if 0 <= student_index < len(result):
                        generate_student_report(student_data, result.iloc[student_index], dates, group_name)
                    else:
                        print("Неверный номер студента.")
        elif choice == '2':
            print("Доступные даты для отчета:")
            available_dates = [date.strftime('%Y-%m-%d') for date in dates]
            for i, d in enumerate(available_dates):
                print(f"{i + 1}. {d}")
            date_choice = input("Выберите номер даты для отчета: ")
            date_index = int(date_choice) - 1
            if 0 <= date_index < len(available_dates):
                generate_attendance_report(student_data, dates, dates[date_index], group_name)
            else:
                print("Неверный номер даты.")
        elif choice == '3':
            view_all_data(student_data, group_name)
        elif choice == '4':
            print("Выход из программы.")
            break
        else:
            print("Неверный выбор. Пожалуйста, попробуйте снова.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Программа для работы с данными о посещаемости студентов.")
    parser.add_argument("--file", default=DEFAULT_EXCEL_FILE, help="Путь к файлу Excel")
    parser.add_argument("--sheet", default=DEFAULT_SHEET_NAME, help="Имя листа в файле Excel")
    args = parser.parse_args()

    main(args.file, args.sheet)