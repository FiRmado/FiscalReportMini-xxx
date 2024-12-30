import win32com.client
import serial.tools.list_ports
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk, scrolledtext
from ttkbootstrap import Window
from datetime import datetime
import time
import os
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials

###############################################################################
# Константы ( абсолютный путь )
project_dir = os.path.dirname(os.path.abspath(__file__))
json_file_path = os.path.join(project_dir, "credentials.json")
current_date = datetime.now().strftime("%Y-%m-%d")

###############################################################################
# Функция для логирования сообщений в текстовое окно
def log_message(message):
    output_text.insert(tk.END, message + "\n")
    output_text.see(tk.END)

###############################################################################
# Функция для смены темы
def change_theme(theme_name):
    try:
        root.style.theme_use(theme_name)
        log_message(f"Тема змінена на: {theme_name}")
    except Exception as e:
        log_message(f"Помилка зміни теми: {e}")

###############################################################################
# Функция для получения списка COM-портов
def get_com_ports():
    ports = serial.tools.list_ports.comports()
    sorted_ports = sorted(ports, key=lambda p: p.device)  # Сортировка по имени устройства
    return [port.device for port in sorted_ports]
    
###############################################################################
# Функция подключения по сом-порту и подключение к .dll библиотеке
def get_ecr_connection():
    port = port_combo.get()  # Получаем выбранный COM-порт
    if not port:
        log_message("Помилка: Оберіть СОМ-порт.")
        return None, None

    # Извлекаем только цифру порта
    try:
        port_number = ''.join(filter(str.isdigit, port))  # Оставляем только цифры
        if not port_number:
            log_message("Помилка: Невірний формат СОМ-порту.")
            return None, None
    except Exception as e:
        log_message(f"Помилка обробки СОМ-порту: {e}")
        return None, None

    log_message(f"Вибраний СОМ-порт: {port_number}")

    # Подключение к OLE-серверу
    try:
        ecr = win32com.client.Dispatch("ecrmini.t400")
        log_message("Підключення до .dll бібліотеки встановлено!")
        return ecr, port_number
    except Exception as e:
        log_message(f"Помилка підключення: {e}")
        return None, None
###############################################################################
# Функция выполнения команды
def execute_command(command, ecr):
    try:
        result = ecr.T400me(command)
        if result:
            log_message(f"Команда '{command}' виконано успішно!")
            return True
        else:
            log_message(f"Команда '{command}' завершилася з помилкою.")
            return False
    except Exception as e:
        log_message(f"Помилка виконання команди '{command}': {e}")
        return False
##############################################################################
# ФУНКЦИЯ СПИСКА СОТРУДНИКОВ
def get_masters():
    return ["Майстер №1", "Майстер №2", "Майстер №3", "Майстер №4"]

###############################################################################
# ОСНОВНЫЕ ФУНКЦИИ ПРОГРАММЫ:
###############################################################################
# СІНХРОНІЗАЦІЯ ЧАСУ
def sync_time_now():
    ecr, port_number = get_ecr_connection()
    if not ecr or not port_number:
        return

    # Открытие порта
    command = f"open_port;{port_number};115200"
    if not ecr.t400me(command):
        log_message(f"Помилка виконання команди: {command}")
        return

    # Проверка состояния смены
    command = "get_status;"
    if ecr.t400me(command):
        try:
            response = ecr.get_last_result
            if response:
                parameters = response.split(';')
                if len(parameters) >= 26:
                    shift_is_open = parameters[2]
                    if shift_is_open == "1":
                        log_message("УВАГА!!! ЗМІНА ВІДКРИТА! НЕОБХІДНО ЗРОБИТИ Z-ЗВІТ")
                        messagebox.showwarning("УВАГА!!!", "УВАГА!!! ЗМІНА ВІДКРИТА! НЕОБХІДНО ЗРОБИТИ Z-ЗВІТ")
                        # Закрываем порт перед завершением функции
                        command = "close_port;"
                        ecr.t400me(command)
                        return
                    elif shift_is_open == "0":
                        log_message("Зміна закрита. Можна виконати синхронізацію часу.")
                    else:
                        log_message(f"Невідомий стан зміни: {shift_is_open}")
                        return
                else:
                    log_message(f"Помилка: Неповна відповідь від РРО: {response}")
                    return
            else:
                log_message("Помилка: Відповідь від РРО відсутня.")
                return
        except Exception as e:
            log_message(f"Помилка обробки відповіді: {e}")
            return
    else:
        log_message("Помилка виконання команди: get_status;")
        return

    # Синхронизация времени с системным временем Windows
    try:
        now = datetime.now()
        hours = now.strftime("%H")  # Часы
        minutes = now.strftime("%M")  # Минуты
        seconds = now.strftime("%S")  # Секунды

        set_time_command = f"set_time;{hours};{minutes};{seconds};"
        if ecr.t400me(set_time_command):
            response = ecr.get_last_result
            log_message(f"Час на касі синхронізовано: {hours}:{minutes}:{seconds}")
            messagebox.showinfo("Синхронізація часу", f"Час успішно синхронізовано: {hours}:{minutes}:{seconds}")
        else:
            log_message(f"Помилка виконання команди: {set_time_command}")
    except Exception as e:
        log_message(f"Помилка синхронізації часу: {e}")

    # Закрытие порта
    command = "close_port;"
    if not ecr.t400me(command):
        log_message(f"Помилка виконання команди: {command}")

######################################################################################
# Функция для X-звіту
def x_report():
    ecr, port_number = get_ecr_connection()
    if not ecr or not port_number:
        return
    
    # Открытие порта
    command = f"open_port;{port_number};115200"
    if not ecr.t400me(command):
        log_message(f"Помилка виконання команди: {command}")
        return

    # Проверка статуса фискальной памяти
    command = "get_fm_status;"
    if ecr.t400me(command):
        response = ecr.get_last_result.strip()  # Убираем лишние пробелы и символы переноса строки
        #log_message(f"Відповідь від get_fm_status: {response}")
        try:
            # Разбиваем ответ на параметры
            params = response.split(";")
            if len(params) >= 10:  # Проверяем, что есть достаточно параметров
                # Получаем параметры 8 и 9 (индексация начинается с 0)
                total_records = int(params[8])  # Параметр 9 - максимальное число отчётов
                used_records = int(params[7])   # Параметр 8 - использованные отчёты
                log_message(f"Фіскальна пам'ять: Використано {used_records} з {total_records}")
                if used_records >= total_records:
                    log_message("Фіскальна пам'ять переповнена. Зняття Х-звіту неможливо.")
                    messagebox.showinfo("Увага", "Фіскальна пам'ять переповнена. Зняття Х-звіту неможливо.")
                    return
            else:
                log_message(f"Помилка: Неповна відповідь від команди get_fm_status. Відповідь: {response}")
                messagebox.showwarning("Помилка", f"Отримано неповну відповідь від команди get_fm_status. Відповідь: {response}")
                return
        except ValueError as ve:
            log_message(f"Помилка конвертації параметрів get_fm_status: {ve}")
            messagebox.showerror("Помилка", f"Неможливо обробити відповідь get_fm_status: {ve}")
            return
        except Exception as e:
            log_message(f"Невідома помилка обробки get_fm_status: {e}")
            messagebox.showerror("Помилка", f"Неможливо обробити відповідь get_fm_status: {e}")
            return
    else:
        log_message("Помилка виконання команди: get_fm_status.")
        messagebox.showerror("Помилка", "Команда get_fm_status не виконана.")
        return

    # Выполнение команд
    commands = [
        f"open_port;{port_number};115200",
        "cashier_registration;1;0",
        "execute_x_report;12321",
        "send_cmd; vp;4F 43 15 63 90 00 00;",
        "cut_paper;",
    ]

    for command in commands:
        if not execute_command(command, ecr):
            log_message(f"Помилка виконання команди: {command}")
            return

    log_message("X-звіт успішно виконано!")

    # Закрытие порта
    command = "close_port;"
    if not ecr.t400me(command):
        log_message(f"Помилка виконання команди: {command}")

######################################################################################
# Функція для Скасування
def cancel_report():    # СНЯТИЕ ФИСКАЛЬНЫХ ОТЧЁТОВ
    ecr, port_number = get_ecr_connection()
    if not ecr or not port_number:
        return
    
    # Открытие порта
    command = f"open_port;{port_number};115200"
    if not ecr.t400me(command):
        log_message(f"Помилка виконання команди: {command}")
        return

    # Проверка статуса фискальной памяти
    command = "get_fm_status;"
    if ecr.t400me(command):
        response = ecr.get_last_result.strip()
        try:
            # Разбиваем ответ на параметры
            params = response.split(";")
            if len(params) >= 10:  # Проверяем, что есть достаточно параметров
                total_records = int(params[8])  # Параметр 9
                used_records = int(params[7])   # Параметр 8
                if used_records >= total_records:
                    log_message("Фіскальна пам'ять переповнена. Зняття Х-звіту буде пропущено.")
                    messagebox.showinfo("Увага", "Фіскальна пам'ять переповнена. Зняття Х-звіту буде пропущено.")
                    skip_x_report = True
                else:
                    skip_x_report = False
            else:
                log_message("Помилка: Неповна відповідь від команди get_fm_status.")
                messagebox.showwarning("Помилка", "Отримано неповну відповідь від команди get_fm_status.")
                return
        except Exception as e:
            log_message(f"Помилка обробки відповіді get_fm_status: {e}")
            messagebox.showerror("Помилка", f"Неможливо обробити відповідь get_fm_status: {e}")
            return
    else:
        log_message("Помилка виконання команди: get_fm_status.")
        messagebox.showerror("Помилка", "Команда get_fm_status не виконана.")
        return

    # Формирование команд с учётом проверки фискальной памяти
    commands = [
        f"open_port;{port_number};115200",
        "cashier_registration;1;0",
    ]
    if not skip_x_report:
        commands.append("execute_x_report;12321")
    commands.extend([
        "execute_report;703;36963;01/01/2015;31/12/2045",
        "send_cmd; vp;4F 43 15 63 90 00 00;",
        "cut_paper;",
    ])

    # Первый цикл выполнения команд
    for command in commands:
        if not execute_command(command, ecr):
            log_message(f"Помилка виконання команди: {command}")
            return

    #log_message("Перший комплект звітів готовий!")

    # Пауза 10 секунд
    time.sleep(10)

    # Второй цикл выполнения команд
    for command in commands:
        if not execute_command(command, ecr):
            log_message(f"Помилка виконання команди: {command}")
            return

    log_message("Скасування успішно виконане!")
    messagebox.showinfo("Скасування успішно виконане!", "Два комплекта звітів роздруковано!")

    # Закрытие порта
    command = "close_port;"
    if not ecr.t400me(command):
        log_message(f"Помилка виконання команди: {command}")
######################################################################################
# КІЛЬКІСТЬ ПАКЕТІВ В РРО
def packet_count():
    ecr, port_number = get_ecr_connection()
    if not ecr or not port_number:
        return

    # Выполнение команды
    command = f"open_port;{port_number};115200"
    if not ecr.t400me(command):  # Проверка успешности выполнения команды
        log_message(f"Помилка виконання команди: {command}")
        return

    # Выполнение команды get_status
    command = "get_status;"
    if ecr.t400me(command):  # Проверка успешности выполнения команды
        try:
            # Получение результата без круглых скобок
            response = ecr.get_last_result
            #print(response)
            if response:
                parameters = response.split(';')  # Разбиваем ответ на параметры
                if len(parameters) >= 26:  # Проверяем, что ответ содержит достаточно параметров
                    shift_is_open = parameters[2]
                    if shift_is_open == "1":
                        log_message("УВАГА!!! ЗМІНА ВІДКРИТА! НЕОБХІДНО ЗРОБИТИ Z-ЗВІТ")
                        messagebox.showwarning("УВАГА!!!", "УВАГА!!! ЗМІНА ВІДКРИТА! НЕОБХІДНО ЗРОБИТИ Z-ЗВІТ")
                    elif shift_is_open == "0":
                        log_message("Зміна закрита")
                    else:
                        log_message(f"Невідомий стан зміни: {shift_is_open}")
                    transmitted_packets = parameters[25]  # Количество переданных отчётов
                    total_packets = parameters[26]  #  Загальна кількість звітів
                    log_message(f"Кількість переданих пакетів: {transmitted_packets}")
                    log_message(f"Загальна кількість пакетів: {total_packets}")
                    if transmitted_packets == total_packets:
                        log_message("Всі дані передані, можна робити скасування реєстрації.")
                        messagebox.showinfo("Всі дані передані!", "Всі дані передані, можна робити скасування реєстрації.")
                    elif transmitted_packets != total_packets:
                        log_message("УВАГА, НЕОБХІДНО ЗРОБИТИ ПЕРЕДАЧУ ДАНИХ ДО ПОДАТКОВОЇ!")
                        messagebox.showwarning("УВАГА!!!", "УВАГА, НЕОБХІДНО ЗРОБИТИ ПЕРЕДАЧУ ДАНИХ ДО ПОДАТКОВОЇ!")
                else:
                    log_message(f"Помилка: Неповна відповідь від РРО: {response}")
            else:
                log_message("Помилка: Відповідь від РРО відсутня.")
        except Exception as e:
            log_message(f"Помилка обробки відповіді: {e}")
    else:
        log_message("Помилка виконання команди: get_status;")

    # Закрытие порта
    command = "close_port;"
    if not ecr.t400me(command):
        log_message(f"Помилка виконання команди: {command}")

######################################################################################
# ПРИМУСОВА ПЕРЕДАЧА ДАНИХ ДО ПОДАТКОВОЇ
def send_data():    # ПРИНУДИТЕЛЬНАЯ ОТПРАВКА ДАННЫХ В НАЛОГОВУЮ
    ecr, port_number = get_ecr_connection()
    if not ecr or not port_number:
        return

    # Выполнение команд
    commands = [
        f"open_port;{port_number};115200",
        "cashier_registration;1;0",
        "dps;2;",
    ]

    for command in commands:
        if not execute_command(command, ecr):
            log_message(f"Помилка виконання команди: {command}")
            return

    log_message("Ініціалізація відправлення даних до податкової розпочалася. Зачекайте 30 секунд і перевірте ще раз кількість пакетів в РРО!")
    messagebox.showwarning("Зачекайте 30 секунд...", "Ініціалізація відправлення даних до податкової розпочалася. Зачекайте 30 секунд і перевірте ще раз кількість пакетів в РРО!")

    # Закрытие порта
    command = "close_port;"
    if not ecr.t400me(command):
        log_message(f"Помилка виконання команди: {command}")

######################################################################################
# Функция опроса пользователя, точно ли надо записать данные в гугл таблицу
def fill_google_sheet():
    # Спрашиваем подтверждение у пользователя
    confirm = messagebox.askyesno("Підтвердження", "Ви точно хочете заповнити Google Таблицю?")
    if confirm:  # Если пользователь нажал "Да"
        #log_message("Заповнення Google Таблиці розпочато...")
        get_rro_info()   # Собираем информацию про кассу
    else:
        log_message("Дія скасована користувачем.")

# Функция записи данных в Google Таблицу
def write_to_google_sheet(serial_number, model, fiscal_number, firm_name, receipt_header, master):
    current_date = datetime.now().strftime("%Y-%m-%d")
    try:
        # Настройка подключения
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(json_file_path, scope)
        client = gspread.authorize(creds)

        # Открываем таблицу по идентификатору
        spreadsheet_id = "Идентификатор Вашей Гугл таблицы"  # идентификатор находиться в адресной строке, между «spreadsheets/d/» и «/edit».
        spreadsheet = client.open_by_key(spreadsheet_id)

        # Получаем текущий месяц и год на украинском
        now = datetime.now()
        ukrainian_months = {
            1: "Січень", 2: "Лютий", 3: "Березень", 4: "Квітень", 5: "Травень",
            6: "Червень", 7: "Липень", 8: "Серпень", 9: "Вересень",
            10: "Жовтень", 11: "Листопад", 12: "Грудень"
        }
        sheet_name = f"{ukrainian_months[now.month]} {now.year % 100:02}"

        # Проверяем, существует ли лист
        try:
            sheet = spreadsheet.worksheet(sheet_name)
        except gspread.WorksheetNotFound:
            # Создаем новый лист, если его нет
            sheet = spreadsheet.add_worksheet(title=sheet_name, rows="100", cols="10")

        # Ищем первую свободную строку
        next_row = len(sheet.col_values(2)) + 1  # Колонка B
        if next_row < 2:
            next_row = 2

        # Подготовка данных для записи
        data = [
            "",  # A колонка (пустая)
            serial_number,  # B
            fiscal_number,  # C
            model,  # D
            firm_name,  # E
            receipt_header,  # F
            master,  # G
            "3 Скасування реєстрації РРО",  # H
            current_date  # I
        ]

        # Записываем данные в строку
        sheet.insert_row(data, next_row)

        log_message(f"Дані успішно записані у лист '{sheet_name}'.")
    except Exception as e:
        log_message(f"Помилка при роботі з Google Таблицею: {e}")

# Функция для извлечения информации из кассы
def get_rro_info():
    ecr, port_number = get_ecr_connection()
    if not ecr or not port_number:
        return

    # Выполнение команд
    commands = [
        f"open_port;{port_number};115200",
        "cashier_registration;1;0"
    ]

    for command in commands:
        if not execute_command(command, ecr):
            log_message(f"Помилка виконання команди: {command}")
            return

    # Получение заводского номера
    command = "read_fm_table;0;1;"
    if ecr.t400me(command):
        response = ecr.get_last_result.strip()
        match = re.search(r"\b(?:ПБ|ПР)\d{10,}", response)
        serial_number = match.group(0) if match else "Невідомо"
        log_message(f"Заводський номер: {serial_number}")
    else:
        log_message("Помилка отримання заводського номера.")
        serial_number = "Невідомо"

    # Получение модели РРО
    # Открытие порта
    command = "get_soft_version"
    if ecr.t400me(command):
        response = ecr.get_last_result.strip()  # Убираем лишние пробелы и символы переноса строки
        if response.startswith("0;"):
            response = response[2:]  # Убираем ведущий '0;'

        # Разбиваем строку на части и извлекаем модель кассы
        parts = response.split(";")
        model = parts[0][:12] if parts else "Невідомо"  # Берём первые 11 символов

        log_message(f"Модель каси: {model}")
    else:
        log_message("Помилка відкриття порту.")
        model = "Невідомо"

    # Получение фискального номера
    command = "read_fm_table;1;5;"
    if ecr.t400me(command):
        response = ecr.get_last_result.strip()
        # Регулярное выражение для фискального номера, начинающегося на 300
        match = re.search(r"\b300\d{7}\b", response)
        fiscal_number = match.group(0) if match else "Невідомо"
        log_message(f"Фіскальний номер: {fiscal_number}")
    else:
        log_message("Помилка отримання фіскального номера.")
        fiscal_number = "Невідомо"

    # Получение шапки чека
        # Получение шапки чека
    command = "get_header;"
    if ecr.t400me(command):
        response = ecr.get_last_result.strip()  # Убираем лишние пробелы и символы переноса строки
        if response.startswith("0;"):
            response = response[2:]  # Убираем ведущий '0;'

        # Разбиваем строку на части и отбираем только нужные значения
        parts = response.split(";")
        header_lines = parts[::3]  # Берём каждый четвёртый элемент, начиная с первого

        # Первая строка - название фирмы
        firm_name = header_lines[0] if header_lines else "Невідомо"
        # Остальные строки объединяются
        receipt_header = " ".join(header_lines[1:]).strip() if len(header_lines) > 1 else "Невідомо"

        log_message("Шапка чека отримана.")
    else:
        log_message("Помилка отримання шапки чека.")
        firm_name = "Невідомо"
        receipt_header = "Невідомо"

    master = master_combo.get()
    write_to_google_sheet(serial_number, model, fiscal_number, firm_name, receipt_header, master)

    # Закрытие порта
    command = "close_port;"
    if not ecr.t400me(command):
        log_message(f"Помилка виконання команди: {command}")

######################################################################################
######################################################################################
# Графическая часть:
# Создание окна
root = Window(themename="superhero")
root.title("Скасування реєстрації")
root.geometry("600x500")
root.resizable(False, False)

# Создание меню
menu_bar = tk.Menu(root)

# Подменю для выбора темы
themes_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Налаштування", menu=themes_menu)

themes = root.style.theme_names()
for theme in themes:
    themes_menu.add_command(label=theme, command=lambda t=theme: change_theme(t))

# Подменю для информации о программе
about_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Про програму", menu=about_menu)
about_menu.add_command(label="Автор: @FiRmado")
about_menu.add_command(label="Дата релізу: 28.12.2024")

root.config(menu=menu_bar)

# Основной фрейм для размещения элементов
main_frame = ttk.Frame(root)
main_frame.pack(expand=True)

# Вложенный фрейм для центрирования элементов
center_frame = ttk.Frame(main_frame)
center_frame.pack(expand=True)

# Выпадающие списки с подписями
port_label = ttk.Label(center_frame, text="Оберіть сом-порт:")
port_label.grid(row=0, column=0, padx=10, pady=5, sticky="n")

ports = get_com_ports()
port_combo = ttk.Combobox(center_frame, values=ports, state="readonly", width=25)
port_combo.grid(row=1, column=0, padx=10, pady=5)
if ports:
    port_combo.set(ports[0])

master_label = ttk.Label(center_frame, text="Оберіть майстра:")
master_label.grid(row=0, column=1, padx=10, pady=5, sticky="n")

masters = get_masters()
master_combo = ttk.Combobox(center_frame, values=masters, state="readonly", width=25)
master_combo.grid(row=1, column=1, padx=10, pady=5)
master_combo.set(masters[0])

# Кнопки
buttons = [
    ("Синхронізація часу", "warning", sync_time_now),
    ("Заповнити гугл таблицю", "light", fill_google_sheet),
    ("Кількість пакетів в РРО", "info", packet_count),
    ("Надіслати дані в ДПС", "primary", send_data),
    ("X-звіт", "success", x_report),
    ("Фіскальні звіти", "danger", cancel_report)
]

# Размещение кнопок в сетке
for i, (text, style, command) in enumerate(buttons):
    button = ttk.Button(center_frame, text=text, style=style, command=command, width=25)
    button.grid(row=2 + i // 2, column=i % 2, padx=10, pady=5)

# Текстовое окно для вывода сообщений
output_text = scrolledtext.ScrolledText(center_frame, width=60, height=15, wrap=tk.WORD, state="normal")
output_text.grid(row=5, column=0, columnspan=2, padx=10, pady=10)

# Запуск приложения
root.mainloop()
