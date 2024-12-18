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


# Функция для смены темы
def change_theme(theme_name):
    try:
        root.style.theme_use(theme_name)
        log_message(f"Тема змінена на: {theme_name}")
    except Exception as e:
        log_message(f"Помилка зміни теми: {e}")


# Функция для получения списка COM-портов
def get_com_ports():
    ports = serial.tools.list_ports.comports()
    sorted_ports = sorted(ports, key=lambda p: p.device)  # Сортировка по имени устройства
    return [port.device for port in sorted_ports]

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


# Функция для логирования сообщений в текстовое окно
def log_message(message):
    output_text.insert(tk.END, message + "\n")
    output_text.see(tk.END)

##############################################################
# Функция для X-звіту
def x_report():
    port = port_combo.get()  # Получаем выбранный COM-порт
    if not port:
        log_message("Помилка: Оберіть СОМ-порт.")
        return

    # Извлекаем только цифру порта
    try:
        port_number = ''.join(filter(str.isdigit, port))  # Оставляем только цифры
        if not port_number:
            log_message("Помилка: Невірний формат СОМ-порту.")
            return
    except Exception as e:
        log_message(f"Помилка обробки СОМ-порту: {e}")
        return

    log_message(f"Вибраний СОМ-порт: {port_number}")
    # Дальнейшая работа с `port_number`

    try:
        # Подключение к OLE-серверу
        ecr = win32com.client.Dispatch("ecrmini.t400")
        log_message("Підключення до .dll бібліотеки встановлено!")
    except Exception as e:
        log_message(f"Помилка підключення: {e}")
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
############################################################
# Функция для Скасування
def cancel_report():    # СНЯТИЕ ФИСКАЛЬНЫХ ОТЧЁТОВ
    port = port_combo.get()  # Получаем выбранный COM-порт
    if not port:
        log_message("Помилка: Оберіть СОМ-порт.")
        return

    # Извлекаем только цифру порта
    try:
        port_number = ''.join(filter(str.isdigit, port))  # Оставляем только цифры
        if not port_number:
            log_message("Помилка: Невірний формат СОМ-порту.")
            return
    except Exception as e:
        log_message(f"Помилка обробки СОМ-порту: {e}")
        return

    log_message(f"Вибраний СОМ-порт: {port_number}")
    # Дальнейшая работа с `port_number`

    try:
        # Подключение к OLE-серверу
        ecr = win32com.client.Dispatch("ecrmini.t400")
        log_message("Підключення до .dll бібліотеки встановлено!")
    except Exception as e:
        log_message(f"Ошибка подключения: {e}")
        return

    # Выполнение команд
    commands = [
        f"open_port;{port_number};115200",
        "cashier_registration;1;0",
        "execute_x_report;12321",
        "execute_report;703;36963;01/01/2015;31/12/2035",
        "send_cmd; vp;4F 43 15 63 90 00 00;",
        "cut_paper;",
    ]

    # Первый цикл выполнения команд
    for command in commands:
        if not execute_command(command, ecr):
            log_message(f"Помилка виконання команди: {command}")
            return

    log_message("Перший комплект звітів готовий!")

    # Пауза 10 секунд
    #log_message("Очікування 10 секунд...")
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

######################################################################
# КІЛЬКІСТЬ ПАКЕТІВ В РРО
def packet_count():
    port = port_combo.get()  # Получаем выбранный COM-порт
    if not port:
        log_message("Помилка: Оберіть СОМ-порт.")
        return

    # Извлекаем только цифру порта
    try:
        port_number = ''.join(filter(str.isdigit, port))  # Оставляем только цифры
        if not port_number:
            log_message("Помилка: Невірний формат СОМ-порту.")
            return
    except Exception as e:
        log_message(f"Помилка обробки СОМ-порту: {e}")
        return

    log_message(f"Вибраний СОМ-порт: {port_number}")

    try:
        # Подключение к OLE-серверу
        ecr = win32com.client.Dispatch("ecrmini.t400")
        log_message("Підключення до .dll бібліотеки встановлено!")
    except Exception as e:
        log_message(f"Помилка підключення: {e}")
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

################################################################
# ПРИМУСОВА ПЕРЕДАЧА ДАНИХ ДО ПОДАТКОВОЇ
def send_data():    # ПРИНУДИТЕЛЬНАЯ ОТПРАВКА ДАННЫХ В НАЛОГОВУЮ
    port = port_combo.get()  # Получаем выбранный COM-порт
    if not port:
        log_message("Помилка: Оберіть СОМ-порт.")
        return

    # Извлекаем только цифру порта
    try:
        port_number = ''.join(filter(str.isdigit, port))  # Оставляем только цифры
        if not port_number:
            log_message("Помилка: Невірний формат СОМ-порту.")
            return
    except Exception as e:
        log_message(f"Помилка обробки СОМ-порту: {e}")
        return

    log_message(f"Вибраний СОМ-порт: {port_number}")
    # Дальнейшая работа с `port_number`

    try:
        # Подключение к OLE-серверу
        ecr = win32com.client.Dispatch("ecrmini.t400")
        log_message("Підключення до .dll бібліотеки встановлено!")
    except Exception as e:
        log_message(f"Помилка підключення: {e}")
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
##############################################################################
##############################################################################
# СІНХРОНІЗАЦІЯ ЧАСУ 
def sync_time_now():
    port = port_combo.get()  # Получаем выбранный COM-порт
    if not port:
        log_message("Помилка: Оберіть СОМ-порт.")
        return

    # Извлекаем только цифру порта
    try:
        port_number = ''.join(filter(str.isdigit, port))  # Оставляем только цифры
        if not port_number:
            log_message("Помилка: Невірний формат СОМ-порту.")
            return
    except Exception as e:
        log_message(f"Помилка обробки СОМ-порту: {e}")
        return

    log_message(f"Вибраний СОМ-порт: {port_number}")

    try:
        # Подключение к OLE-серверу
        ecr = win32com.client.Dispatch("ecrmini.t400")
        log_message("Підключення до .dll бібліотеки встановлено!")
    except Exception as e:
        log_message(f"Помилка підключення: {e}")
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
            log_message(f"Час на касі синхронізовано: {hours}:{minutes}:{seconds}. Відповідь: {response}")
            messagebox.showinfo("Синхронізація часу", f"Час успішно синхронізовано: {hours}:{minutes}:{seconds}")
        else:
            log_message(f"Помилка виконання команди: {set_time_command}")
    except Exception as e:
        log_message(f"Помилка синхронізації часу: {e}")

    # Закрытие порта
    command = "close_port;"
    if not ecr.t400me(command):
        log_message(f"Помилка виконання команди: {command}")

##############################################################################        
# Создание интерфейса с использованием ttkbootstrap
root = Window(themename="superhero")  
root.title("Скасування реєстрації")
root.geometry("600x500")
root.resizable(False, False)

# Создание меню
menu_bar = tk.Menu(root)

# Подменю для выбора темы
themes_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Налаштування", menu=themes_menu)

themes = root.style.theme_names()  # Получение всех доступных тем
for theme in themes:
    themes_menu.add_command(label=theme, command=lambda t=theme: change_theme(t))

# Подменю для информации о программе
about_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Про програму", menu=about_menu)
about_menu.add_command(label="Автор: @FiRmado")
about_menu.add_command(label="Дата релізу: 16.12.2024")

# Устанавливаем меню
root.config(menu=menu_bar)

# Выпадающий список для выбора порта
port_label = ttk.Label(root, text="Оберіть сом-порт: ")
port_label.pack(pady=5)

ports = get_com_ports()  # Получаем список доступных COM-портов
port_combo = ttk.Combobox(root, values=ports, state="readonly", width=25)
port_combo.pack(pady=5)
if ports:
    port_combo.set(ports[0])  # Устанавливаем первый порт по умолчанию

clock_button = ttk.Button(root, text="Синхронізація часу", style="warning", command=sync_time_now, width=25)
clock_button.pack(pady=10)

# Создание фрейма для кнопок
button_frame = ttk.Frame(root)
button_frame.pack(pady=10)

# Кнопки
x_report_button = ttk.Button(button_frame, text="X-звіт", style="success", command=x_report, width=25)
x_report_button.grid(row=0, column=0, padx=10, pady=5)

cancel_button = ttk.Button(button_frame, text="Фіскальні звіти", style="danger", command=cancel_report, width=25)
cancel_button.grid(row=0, column=1, padx=10, pady=5)

# Новые кнопки
packet_count_button = ttk.Button(button_frame, text="Кількість пакетів в РРО", style="info", command=packet_count, width=25)
packet_count_button.grid(row=1, column=0, padx=10, pady=5)

send_data_button = ttk.Button(button_frame, text="Надіслати дані в ДПС", style="primary", command=send_data, width=25)
send_data_button.grid(row=1, column=1, padx=10, pady=5)


# Текстовое окно для вывода сообщений
output_text = scrolledtext.ScrolledText(root, width=60, height=15, wrap=tk.WORD, state="normal")
output_text.pack(pady=10)

# Запуск приложения
root.mainloop()