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
#from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

from Info.info import info_act, info_spec, info_head
#from Open.solid_spec import solid_parce
from Open.solid_spec_sale import solid_parce_sale
#from Check.check_solid_acts import fill_table
from Check.check_solid_acts_sale import fill_table_sale



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

# Функция создания акта
def safe_act():
    # Проверка если спецификация была не загружена, то нет возможности создать акт
    if len(tube_list) <= 1:
        info_act()
        return
    global document

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

    if type_choies.get() == 'входного контроля элементов трубопровода':
        context['number'] = '2'
        context['name'] = 'ВХОДНОГО КОНТРОЛЯ ЭЛЕМЕНТОВ ТРУБОПРОВОДА'
    elif type_choies.get() == 'входного контроля материалов':
        context['number'] = '3'
        context['name'] = 'ВХОДНОГО КОНТРОЛЯ МАТЕРИАЛОВ'
    elif type_choies.get() == 'проверки чистоты труб':
        context['number'] = '4'
        context['name'] = 'ПРОВЕРКИ ЧИСТОТЫ ТРУБ'
    elif type_choies.get() == 'монтажа опор под трубопроводы':
        context['number'] = '5'
        context['name'] = 'МОНТАЖА ОПОР ПОД ТРУБОПРОВОДЫ'
    elif type_choies.get() == 'монтажа трубопроводов':
        context['number'] = '6'
        context['name'] = 'МОНТАЖА ТРУБОПРОВОДОВ'
    elif type_choies.get() == '*закупка*':
        context['number'] = '100'
        context['name'] = 'ЗАКУПКА'

    # Если спецификация загружена и выбран файл для сохранения акта, то выполняется код ниже
    filepath = filedialog.asksaveasfilename(defaultextension='docx', initialfile='Акт '+str(context['name']).lower())
    if filepath != "" and len(tube_list)>1:
        # Заголовки таблицы (при использовании шаблона они не нужны и используюся только для подсчета
        # столбцов
        headers = ('№ ', 'Наименование', 'Размеры, материал',
                   'Техническая характеристика', 'Кол-во', 'Ед.')

        # Заполнение шаблона данными
        document.render(context)

        # Получение списка таблиц из файла шаблона
        all_tables = document.tables
        # Поиск таблицы с одной строкой в шаблоне
        new_table = all_tables[0]
        # Количество колонок таблицы
        cols_number = len(headers)

        new_table=fill_table_sale(type_choies.get(), cols_number, elements_list, pad_list, tube_list, support_list, sales_tube_list, new_table)

    document.save(filepath)

# функция открытия файла с полями шапки
def open_head():
    # Информационное окно напоминающее, что нужно выбрать для открытия файл спецификации формата xlsx
    info_head()
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
    info_spec()
    table_path=filedialog.askopenfilename()
    global tube_list
    global pad_list
    global support_list
    global elements_list
    global sales_tube_list

    tube_list=solid_parce_sale(table_path)[0]
    pad_list=solid_parce_sale(table_path)[1]
    support_list=solid_parce_sale(table_path)[2]
    elements_list=solid_parce_sale(table_path)[3]
    sales_tube_list=solid_parce_sale(table_path)[4]


# Создается оконное приложение
root=Tk()
# Заголовок
root.title('Подготовка актов из спецификации Solidworks')
root["bg"] = "aquamarine4"
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

type_acts=['входного контроля элементов трубопровода', 'входного контроля материалов',
           'проверки чистоты труб', 'монтажа опор под трубопроводы', 'монтажа трубопроводов', '*закупка*']
# по умолчанию будет выбран первый элемент из languages
type_var = StringVar(value=type_acts[0])

# Ниспадающий список
type_choies=Combobox(textvariable=type_var, values=type_acts, state="readonly")
type_choies.place(x=20, y=330, width=350)

#Кнопка открытия спецификации
file_button=Button(text='Открыть спец', command=open_table, font=("Arial", 12, "bold"))
file_button.place(x=400, y=20)

#Кнопка открытия файла xlsx c полями шапки
head_button=Button(text='Поля шапки', command=open_head, font=("Arial", 12, "bold"))
head_button.place(x=560, y=20)

#Кнопка создания актов
btn=Button(text='Создать Акт', command=safe_act, font=("Arial", 12, "bold"))
btn.place(x=560, y=320)


# Загрузка шаблона
document = DocxTemplate(resource_path('res\Шаблон_solid.docx'))

tube_list=[]
pad_list=[]
support_list=[]
elements_list=[]
sales_tube_list=[]

#df_fo_dic = pd.read_excel(resource_path('res\Словарь.xlsx'))
#model_dict=dict(zip(df_fo_dic['id'], df_fo_dic['value']))

root.mainloop()