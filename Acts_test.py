from tkinter.ttk import Combobox
#import numpy as np
from tkinter import *
from tkinter import filedialog
import pandas as pd
from docxtpl import DocxTemplate
#from tkinter.messagebox import showinfo
from tkinter import font
import sys
import os

from Check.check_word_list import fill_table_project
from Open.spec_project import spec_parce
from Info import info




#при запуске exe'шника все внутренности программы распаковываюся
# во временную папку Windows. Соответственно, в скрипте нужно обращаться к ним
#Для обращения к этим файлам и нужна функция ниже
#https://ru.stackoverflow.com/questions/1472473/%d0%90%d0%b2%d1%82%d0%be%d0%bd%d0%be%d0%bc%d0%bd%d1%8b%d0%b9-%d0%b8%d1%81%d0%bf%d0%be%d0%bb%d0%bd%d1%8f%d0%b5%d0%bc%d1%8b%d0%b9-%d1%84%d0%b0%d0%b9%d0%bb-%d0%b2-python
# По ссылке смотри описание
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# функция сохранения всех актов сразу
def safe_all_acts():
    # Проверка если спецификация была не загружена, то нет возможности создать акт
    if len(lst_xl) <= 1:
        info.info_act()
        return
    global enabled
    enabled=1
    global document_spec
    global type_var
    global dir_path
    # Если спецификация загружена и выбран файл для сохранения акта, то выполняется код ниже
    info.info_all_acts()
    dir_path = filedialog.asksaveasfilename()
    dir_path = dir_path[:dir_path.rfind('/')]

    for i in range(0,5):
        type_var=type_acts[i]
        type_choies.set(type_var)
        safe_act()

    type_var=type_acts[0]
    type_choies.set(type_var)

# Функция создания акта
def safe_act():
    # Проверка если спецификация была не загружена, то нет возможности создать акт
    if len(lst_xl) <= 1:
        info.info_act()
        return
    global document_spec

    # Данные для заполнения шаблона
    context = {
        'station': station.get(),
        'calc': calc.get(),
        'company': company.get(),
        'object': obj.get(),
        'address': address.get(),
        # 'number': number.get(),
        # 'name': name.get(),
        'data': data.get()
    }

    if type_choies.get() == 'входного контроля основного оборудования':
        context['number'] = '1.1'
        context['name'] = 'ВХОДНОГО КОНТРОЛЯ ОСНОВНОГО ОБОРУДОВАНИЯ'
    elif type_choies.get() == 'установки основного оборудования':
        context['number'] = '7'
        context['name'] = 'УСТАНОВКИ ОСНОВНОГО ОБОРУДОВАНИЯ'
    elif type_choies.get() == 'входного контроля арматуры':
        context['number'] = '1.2'
        context['name'] = 'ВХОДНОГО КОНТРОЛЯ АРМАТУРЫ'
    elif type_choies.get() == 'монтажа арматуры':
        context['number'] = '8'
        context['name'] = 'МОНТАЖА АРМАТУРЫ'
    elif type_choies.get() == 'входного контроля оборудования КИПиА':
        context['number'] = '1.3'
        context['name'] = 'ВХОДНОГО КОНТРОЛЯ ОБОРУДОВАНИЯ КИПиА'

    # Если спецификация загружена и выбран файл для сохранения акта, то выполняется код ниже
    filepath=''
    #Проверяется нажата ли кнопка все акты сразу
    if enabled==1:
        filepath=dir_path+'/Акт '+str(context['name']).lower()+'.docx'

    elif enabled==0:
        filepath = filedialog.asksaveasfilename(defaultextension='docx', initialfile='Акт '+str(context['name']).lower())
    if filepath != "" and len(lst_xl)>1:
        # Заголовки таблицы (при использовании шаблона они не нужны и используюся только для подсчета
        # столбцов
        headers = ('№ ', 'Поз.', 'Наименование', 'Тип, марка\nматериал\nТехническая\nдокументация',
                   'Завод -\nизготовитель', 'Кол-\nво,\nшт')

        # Заполнение шаблона данными
        document.render(context)
        # Получение списка таблиц из файла шаблона
        all_tables = document.tables
        # Поиск таблицы с одной строкой в шаблоне
        new_table = all_tables[0]
        # Количество колонок таблицы
        cols_number = len(headers)

        new_table=fill_table_project(type_choies.get(), cols_number, lst_xl, new_table)

        document.save(filepath)

# функция открытия файла с полями шапки
def open_head():
    # Информационное окно напоминающее, что нужно выбрать для открытия файл спецификации формата xlsx
    info.info_head()
    head_path=filedialog.askopenfilename()
    global name_station
    global number_calc
    global company_name
    global obj_name
    global addr_obj
    global dt
    # Если файл выбран, то данные из листа с названием Table 1 читаются датафрейм пандас
    if head_path !="":
       df = pd.read_excel(head_path)
       name_station.set(df[df.columns[1]].iloc[0])
       number_calc.set(df[df.columns[1]].iloc[1])
       company_name.set(df[df.columns[1]].iloc[2])
       obj_name.set(df[df.columns[1]].iloc[3])
       addr_obj.set(df[df.columns[1]].iloc[4])
       dt.set(df[df.columns[1]].iloc[5])


#Функция открытия xlsx файла с помощью диалогового окна выбора файла в системе
def open_table():
    # Информационное окно напоминающее, что нужно выбрать для открытия файл спецификации формата xlsx
    info.info_spec()
    table_path=filedialog.askopenfilename()
    global lst_xl
    lst_xl=spec_parce(table_path)

# Создается оконное приложение
root=Tk()
# Заголовок
root.title('Подготовка Актов')
# Размер окна
root.geometry('800x500+100+100')
# Иконка в титуле приложения
root.iconbitmap(default=resource_path('res/_brend.ico'))

# Стили
font1 = font.Font(family= "Times New Roman", size=11, weight="bold", slant="roman", underline=False, overstrike=False)
font2 = font.Font(family= "Times New Roman", size=11, weight="normal", slant="roman", underline=False, overstrike=False)

#Текстовая метка
name_form=Label(root, text='Заполните данные шапки акта', font=("Arial", 11, "bold"))
name_form.place(x=20, y=20)

#привязываем переменную name_station к полю ввода названия установки
name_station = StringVar()

#Поле ввода текста
station =Entry(root, font=font1, textvariable=name_station)
station.place(x=20, y=60, width=650)
station.insert(0,'Введите название установки')

#привязываем переменную number_cal к полю ввода номер расчета
number_calc = StringVar()

calc =Entry(root, font=font1, textvariable=number_calc)
calc.place(x=20, y=100, width=650)
calc.insert(0,'Введите номер расчета')

#привязываем переменную company_name к полю ввода название компании
company_name = StringVar()

company =Entry(root, font=font2, textvariable=company_name)
company.place(x=20, y=140, width=650)
company.insert(0,'Введите название компании')

#привязываем переменную obj_name к полю ввода название обьекта
obj_name = StringVar()

obj =Entry(root, font=font2, textvariable=obj_name)
obj.place(x=20, y=180, width=650)
obj.insert(0,'Введите название обьекта')

#привязываем переменную addr_obj к полю ввода адресс обьекта
addr_obj = StringVar()

address =Entry(root, font=font2, textvariable=addr_obj)
address.place(x=20, y=220, width=650)
address.insert(0,'Введите название адреса')

"""""
number =Entry(root, font=font1)
number.place(x=20, y=260, width=650)
number.insert(0,'Введите номер акта')

name =Entry(root, font=font1)
name.place(x=20, y=300, width=650)
name.insert(0,'Введите название акта')

"""

#привязываем переменную dt к полю ввода дата
dt = StringVar()

data =Entry(root, font=font2, textvariable=dt )
data.place(x=20, y=260, width=650)
data.insert(0,'Введите дату')

acttype=Label(root, text='Выберите тип акта', font=("Arial", 11, "bold"))
acttype.place(x=20, y=300)

type_acts=['входного контроля основного оборудования', 'входного контроля арматуры',
           'входного контроля оборудования КИПиА', 'установки основного оборудования', 'монтажа арматуры']
# по умолчанию будет выбран первый элемент из languages
type_var = StringVar(value=type_acts[0])

# Ниспадающий список
type_choies=Combobox(textvariable=type_var, values=type_acts, state="readonly")
type_choies.place(x=20, y=330, width=350)

#enabled=IntVar()

#Чек бокс для сохранения всех актов сразу
#all_checkbox=Checkbutton(text='Сделать все акты сразу', variable=enabled, command=safe_all_acts)
#all_checkbox.place(x=560, y=400)

#Кнопка открытия спецификации
spec_button=Button(text='Открыть спец', command=open_table, font=("Arial", 12, "bold"))
spec_button.place(x=400, y=20)

#Кнопка открытия файла xlsx c полями шапки
head_button=Button(text='Поля шапки', command=open_head, font=("Arial", 12, "bold"))
head_button.place(x=560, y=20)

#Кнопка создания актов
act_button=Button(text='Создать Акт', command=safe_act, font=("Arial", 12, "bold"))
act_button.place(x=560, y=320)

enabled=0
#Кнопка сохранения всех актов сразу
all_acts_button=Button(text='Все акты сразу', command=safe_all_acts, font=("Arial", 12, "bold"))
all_acts_button.place(x=560, y=360)



# Загрузка шаблона
document_spec = DocxTemplate(resource_path('res\Шаблон.docx'))

lst_xl=[]
dir_path=''


root.mainloop()