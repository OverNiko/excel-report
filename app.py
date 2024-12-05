import pandas as pd
from datetime import datetime, timedelta
import logging
from typing import Union, Optional
from tabulate import tabulate
from openpyxl import load_workbook
import argparse

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

DEFAULT_EXCEL_FILE = 'Черненко Александр Александрович.xlsx'
BASE_DATE = datetime(2024, 10, 3)

def get_sheet_names(excel_file: str) -> list:
    """Получает список имен всех листов в Excel файле."""
    workbook = load_workbook(excel_file, read_only=True)
    return workbook.sheetnames

def choose_sheet(excel_file: str) -> str:
    """Позволяет пользователю выбрать лист для работы."""
    sheet_names = get_sheet_names(excel_file)
    print("Доступные листы:")
    for i, name in enumerate(sheet_names, 1):
        print(f"{i}. {name}")
    
    while True:
        try:
            choice = int(input("Выберите номер листа: ")) - 1
            if 0 <= choice < len(sheet_names):
                return sheet_names[choice]
            else:
                print("Неверный номер. Попробуйте еще раз.")
        except ValueError:
            print("Пожалуйста, введите число.")

def parse_date(date: Union[datetime, str, float]) -> Optional[datetime]:
    """Преобразует различные форматы дат в объект datetime."""
    if isinstance(date, (datetime, str)):
        return pd.to_datetime(date)
    number = int(float(str(date)))
    return BASE_DATE + timedelta(days=(number - 1) * 7)

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
        
        # Сортируем данные по столбцу '№ п/п'
        attendance_data = attendance_data.sort_values('№ п/п')
        
        # Сбрасываем индекс и пересоздаем столбец '№ п/п'
        attendance_data = attendance_data.reset_index(drop=True)
        attendance_data['№ п/п'] = attendance_data.index + 1
        
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

def generate_student_report(student_data: pd.DataFrame, dates: list, group_name: str) -> None:
    """Генерирует отчет о посещаемости для выбранного студента."""
    print("Список студентов:")
    # Сортируем student_data по '№ п/п' перед выводом списка
    sorted_student_data = student_data.sort_values('№ п/п')
    for i, (_, student) in enumerate(sorted_student_data.iterrows(), 1):
        print(f"{i}. {student['Фамилия']} {student['Имя']}")
    
    while True:
        try:
            student_choice = int(input("Выберите номер студента для отчета: ")) - 1
            if 0 <= student_choice < len(sorted_student_data):
                student = sorted_student_data.iloc[student_choice]
                break
            else:
                print("Неверный номер студента. Попробуйте еще раз.")
        except ValueError:
            print("Пожалуйста, введите число.")
    
    attendance_data = pd.DataFrame({
        'Дата': dates,
        'Присутствие': student.iloc[4:].values
    })
    attendance_data['Присутствие'] = attendance_data['Присутствие'].fillna(0).astype(int)
    
    print(f"\nОтчет о посещаемости студента(ки) {student['Фамилия']} {student['Имя']} {student['Отчество']}:")
    print(tabulate(attendance_data, headers='keys', tablefmt='pretty', showindex=False))
    
    if input("Хотите сохранить отчет для этого студента? (да/нет): ").strip().lower() == 'да':
        html_content = f"""
        <h1>Отчет о посещаемости для группы: {group_name}</h1>
        <h2>Студент: {student['Фамилия']} {student['Имя']} {student['Отчество']}</h2>
        {attendance_data.to_html(index=False)}
        """
        file_name = f"отчет_о_посещаемости_студента(ки)_{student['Фамилия']}_{group_name}.html"
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
        print("1. Генерация отчета о посещаемости для студента")
        print("2. Генерация отчета о посещаемости за день/дату")
        print("3. Просмотр всех данных файла")
        print("4. Выход")
        
        choice = input("Введите номер действия: ")
        
        if choice == '1':
            generate_student_report(student_data, dates, group_name)
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
    args = parser.parse_args()

    sheet_name = choose_sheet(args.file)
    main(args.file, sheet_name)