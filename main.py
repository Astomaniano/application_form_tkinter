import tkinter as tk
import pandas as pd
from pandas import ExcelWriter
import datetime
import os

def get_next_request_number():
    filename = 'заявки.xlsx'  # Имя файла в кавычках
    if not os.path.isfile(filename):
        return 1
    else:
        df = pd.read_excel(filename)  # Имя файла в кавычках
        if df.empty:
            return 1
        else:
            return df['Номер'].max() + 1

def save_to_excel():

    n = get_next_request_number()

    now = datetime.datetime.now().strftime("%Y.%m.%d %H:%M:%S")

    # Собираем данные из полей ввода
    name = entry_name.get()
    theme = entry_theme.get()
    comment = entry_comment.get("1.0", "end-1c").strip()

    # Создаем DataFrame из собранных данных
    df = pd.DataFrame({'Номер': [n], 'Дата и время': [now], 'ФИО': [name], 'Тема заявки': [theme], 'Комментарии': [comment]})

    # Проверяем, существует ли файл
    if not os.path.isfile('заявки.xlsx'):
        # Если файл не существует, создаем его и записываем данные с заголовками
        df.to_excel('заявки.xlsx', index=False)
    else:
        # Если файл существует, добавляем данные без записи заголовков
        with ExcelWriter('заявки.xlsx', mode='a', if_sheet_exists='overlay') as writer:
            df.to_excel(writer, index=False, header=False, startrow=writer.sheets['Sheet1'].max_row)

    # Очищаем поля ввода
    clear()


def clear():
    entry_name.delete(0, 'end')
    entry_theme.delete(0, 'end')
    entry_comment.delete(1.0, 'end')

n = 1
now = datetime.datetime.now().strftime("%Y.%m.%d %H:%M:%S")

root = tk.Tk()
root.title('форма для заявок')
root.configure(background='gray19')
root.geometry('350x550+200+200')

label_text = tk.Label(root, text='Добро пожаловать!\n Оставьте свою заявку здесь!', font=("Arial", 12, "italic"), bg='gray19', fg='white')
label_text.pack(pady=40)

text_1 = tk.Label(root, text='Ваше ФИО', bg='gray19', fg='white')
text_1.pack(pady=5)
entry_name = tk.Entry(root, width=40)
entry_name.pack(pady=5)

text_2 = tk.Label(root, text='Тема заявки:', bg='gray19', fg='white')
text_2.pack(pady=5)
entry_theme = tk.Entry(root, width=40)
entry_theme.pack(pady=5)

text_3 = tk.Label(root, text='Коментарии к заявке:', bg='gray19', fg='white')
text_3.pack(pady=5)
entry_comment = tk.Text(root, width=30, height=10)  # width и height задаются в символах и строках соответственно
entry_comment.pack(pady=5)

send_button = tk.Button(root, text='Создать заявку', command=save_to_excel, bg='darkseagreen4', fg='white')
send_button.pack(pady=10)

clear_button = tk.Button(root, text='Очистить все поля', command=clear, bg='orangered4', fg='white')
clear_button.pack(pady=5)

root.mainloop()