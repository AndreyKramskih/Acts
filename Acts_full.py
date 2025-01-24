from email.utils import format_datetime
from tkinter.ttk import Combobox
from tkinter import ttk
#import numpy as np
from tkinter import *
from tkinter import filedialog
import pandas as pd
from docxtpl import DocxTemplate
#from tkinter.messagebox import showinfo
#from tkinter import font
import sys
import os
from tkinter.ttk import Style

from Check.check_word_list import fill_table_project
from Open.spec_project import spec_parce
from Open.spec_project_new_separ import spec_parce_new_separ
from Info import info
from View.specific_window import SpecificFrame, SolidFrame
from Open.solid_spec import solid_parce
from Check.check_solid_acts import fill_table_solid




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

#функция смены фона при выборе разных вкладок

def tab_index(event):
    if notebook.index('current')==0:
        test.theme_use("my_tables_1")
    elif notebook.index('current')==1:
        test.theme_use("my_tables_2")




# функция сохранения всех актов сразу
def safe_all_acts(event):
    # Проверка если спецификация была не загружена, то нет возможности создать акт
    if len(lst_xl) <= 1:
        info.info_act()
        return
    global all_safe_enabled
    enabled=1
    global document_spec

    global dir_path
    # Если спецификация загружена и выбран файл для сохранения акта, то выполняется код ниже
    info.info_all_acts()
    dir_path = filedialog.asksaveasfilename()
    dir_path = dir_path[:dir_path.rfind('/')]

    for i in range(0,5):
        project_frame.type_var=project_frame.type_acts[i]
        project_frame.type_choies.set(project_frame.type_var)
        safe_act_project(event)

    project_frame.type_var=project_frame.type_acts[0]
    project_frame.type_choies.set(project_frame.type_var)

    enabled=0


# функция сохранения всех актов из солида сразу
def safe_all_acts_solid(event):
    # Проверка если спецификация была не загружена, то нет возможности создать акт
    if len(tube_list) <= 1:
        info.info_act()
        return
    global all_safe_enabled
    enabled=1
    global document_solid

    global dir_path
    # Если спецификация загружена и выбран файл для сохранения акта, то выполняется код ниже
    info.info_all_acts()
    dir_path = filedialog.asksaveasfilename()
    dir_path = dir_path[:dir_path.rfind('/')]

    for i in range(0,5):
        solid_frame.type_var=solid_frame.type_acts[i]
        solid_frame.type_choies.set(solid_frame.type_var)
        safe_act_solid(event)

    solid_frame.type_var=solid_frame.type_acts[0]
    solid_frame.type_choies.set(solid_frame.type_var)
    enabled=0



# Функция создания акта спецификации
def safe_act_project(event):
    # Проверка если спецификация была не загружена, то нет возможности создать акт
    if len(lst_xl) <= 1:
        info.info_act()
        return
    global document_spec

    # Данные для заполнения шаблона
    context = {
        'station': project_frame.station.get(),
        'calc': project_frame.calc.get(),
        'company': project_frame.company.get(),
        'object': project_frame.obj.get(),
        'address': project_frame.address.get(),
        # 'number': number.get(),
        # 'name': name.get(),
        'data': project_frame.data.get()
    }

    if project_frame.type_choies.get() == 'входного контроля основного оборудования':
        context['number'] = '1.1'
        context['name'] = 'ВХОДНОГО КОНТРОЛЯ ОСНОВНОГО ОБОРУДОВАНИЯ'
    elif project_frame.type_choies.get() == 'установки основного оборудования':
        context['number'] = '7'
        context['name'] = 'УСТАНОВКИ ОСНОВНОГО ОБОРУДОВАНИЯ'
    elif project_frame.type_choies.get() == 'входного контроля арматуры':
        context['number'] = '1.2'
        context['name'] = 'ВХОДНОГО КОНТРОЛЯ АРМАТУРЫ'
    elif project_frame.type_choies.get() == 'монтажа арматуры':
        context['number'] = '8'
        context['name'] = 'МОНТАЖА АРМАТУРЫ'
    elif project_frame.type_choies.get() == 'входного контроля оборудования КИПиА':
        context['number'] = '1.3'
        context['name'] = 'ВХОДНОГО КОНТРОЛЯ ОБОРУДОВАНИЯ КИПиА'

    # Если спецификация загружена и выбран файл для сохранения акта, то выполняется код ниже
    filepath=''
    #Проверяется нажата ли кнопка все акты сразу
    if all_safe_enabled==1:
        filepath=dir_path+'/Акт '+str(context['name']).lower()+'.docx'

    elif all_safe_enabled==0:
        filepath = filedialog.asksaveasfilename(defaultextension='docx', initialfile='Акт '+str(context['name']).lower())
    if filepath != "" and len(lst_xl)>1:
        # Заголовки таблицы (при использовании шаблона они не нужны и используюся только для подсчета
        # столбцов
        headers = ('№ ', 'Поз.', 'Наименование', 'Тип, марка\nматериал\nТехническая\nдокументация',
                   'Завод -\nизготовитель', 'Кол-\nво,\nшт')

        # Заполнение шаблона данными
        document_spec.render(context)
        # Получение списка таблиц из файла шаблона
        all_tables = document_spec.tables
        # Поиск таблицы с одной строкой в шаблоне
        new_table = all_tables[0]
        # Количество колонок таблицы
        cols_number = len(headers)

        new_table=fill_table_project(project_frame.type_choies.get(), cols_number, lst_xl, new_table)

        document_spec.save(filepath)


# Функция создания акта из спецификации солида
def safe_act_solid(event):
    # Проверка если спецификация была не загружена, то нет возможности создать акт
    if len(tube_list) <= 1:
        info.info_act()
        return
    global document_solid

    # Данные для заполнения шаблона
    context = {
        'station': solid_frame.station.get(),
        'calc': solid_frame.calc.get(),
        'company': solid_frame.company.get(),
        'object': solid_frame.obj.get(),
        'address': solid_frame.address.get(),
        # 'number': number.get(),
        # 'name': name.get(),
        'data': solid_frame.data.get()
    }

    if solid_frame.type_choies.get() == 'входного контроля элементов трубопровода':
        context['number'] = '2'
        context['name'] = 'ВХОДНОГО КОНТРОЛЯ ЭЛЕМЕНТОВ ТРУБОПРОВОДА'
    elif solid_frame.type_choies.get() == 'входного контроля материалов':
        context['number'] = '3'
        context['name'] = 'ВХОДНОГО КОНТРОЛЯ МАТЕРИАЛОВ'
    elif solid_frame.type_choies.get() == 'проверки чистоты труб':
        context['number'] = '4'
        context['name'] = 'ПРОВЕРКИ ЧИСТОТЫ ТРУБ'
    elif solid_frame.type_choies.get() == 'монтажа опор под трубопроводы':
        context['number'] = '5'
        context['name'] = 'МОНТАЖА ОПОР ПОД ТРУБОПРОВОДЫ'
    elif solid_frame.type_choies.get() == 'монтажа трубопроводов':
        context['number'] = '6'
        context['name'] = 'МОНТАЖА ТРУБОПРОВОДОВ'
 # Если спецификация загружена и выбран файл для сохранения акта, то выполняется код ниже
    filepath = ''
    # Проверяется нажата ли кнопка все акты сразу
    if all_safe_enabled == 1:
        filepath = dir_path + '/Акт ' + str(context['name']).lower() + '.docx'

    elif all_safe_enabled == 0:
        filepath = filedialog.asksaveasfilename(defaultextension='docx',
                                                    initialfile='Акт ' + str(context['name']).lower())
    if filepath != "" and len(tube_list) > 1:
        # Заголовки таблицы (при использовании шаблона они не нужны и используюся только для подсчета
        # столбцов
        headers = ('№ ', 'Наименование', 'Размеры, материал',
                   'Техническая характеристика', 'Кол-во', 'Ед.')

        # Заполнение шаблона данными
        document_solid.render(context)

        # Получение списка таблиц из файла шаблона
        all_tables = document_solid.tables
        # Поиск таблицы с одной строкой в шаблоне
        new_table = all_tables[0]
        # Количество колонок таблицы
        cols_number = len(headers)

        new_table=fill_table_solid(solid_frame.type_choies.get(), cols_number, elements_list, pad_list, tube_list, support_list, new_table)

    document_solid.save(filepath)

# функция открытия файла с полями шапки для актов проекта и солида
def open_head(event):
    # Информационное окно напоминающее, что нужно выбрать для открытия файл спецификации формата xlsx
    info.info_head()
    head_path=filedialog.askopenfilename()



    # Если файл выбран, то данные из листа с названием Table 1 читаются датафрейм пандас
    if head_path !="":
       df = pd.read_excel(head_path)
       solid_frame.name_station.set(df[df.columns[1]].iloc[0])
       solid_frame.number_calc.set(df[df.columns[1]].iloc[1])
       solid_frame.company_name.set(df[df.columns[1]].iloc[2])
       solid_frame.obj_name.set(df[df.columns[1]].iloc[3])
       solid_frame.addr_obj.set(df[df.columns[1]].iloc[4])
       solid_frame.dt.set(df[df.columns[1]].iloc[5])

       project_frame.name_station.set(df[df.columns[1]].iloc[0])
       project_frame.number_calc.set(df[df.columns[1]].iloc[1])
       project_frame.company_name.set(df[df.columns[1]].iloc[2])
       project_frame.obj_name.set(df[df.columns[1]].iloc[3])
       project_frame.addr_obj.set(df[df.columns[1]].iloc[4])
       project_frame.dt.set(df[df.columns[1]].iloc[5])





#Функция открытия xlsx файла с помощью диалогового окна выбора файла в системе
def open_table_project(event):
    # Информационное окно напоминающее, что нужно выбрать для открытия файл спецификации формата xlsx
    info.info_spec()
    table_path=filedialog.askopenfilename()
    global lst_xl
    #lst_xl=spec_parce_new_separ(table_path, 1)
    lst_xl = spec_parce(table_path)


#Функция открытия xlsx файла с помощью диалогового окна выбора файла в системе
def open_table_solid(event):
    # Информационное окно напоминающее, что нужно выбрать для открытия файл спецификации формата xlsx
    info.info_spec()
    table_path=filedialog.askopenfilename()
    global tube_list
    global pad_list
    global support_list
    global elements_list

    tube_list = solid_parce(table_path)[0]
    pad_list = solid_parce(table_path)[1]
    support_list = solid_parce(table_path)[2]
    elements_list = solid_parce(table_path)[3]


# Создается оконное приложение
root=Tk()
# Заголовок
root.title('Подготовка Актов')
# Размер окна
root.geometry('800x500+100+100')
# Иконка в титуле приложения
root.iconbitmap(default=resource_path('res/_brend.ico'))

# создаем набор вкладок
notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill=BOTH)


# наполняем вкладки элементами через соответствущие классы
project_frame=SpecificFrame(notebook)
solid_frame=SolidFrame(notebook)


# добавляем картинку во вкладки
project_frame.pack(fill=BOTH, expand=True)
spec_logo = PhotoImage(file=resource_path('res/Spec_21.png'))
notebook.add(project_frame, text="Cпецификация", image=spec_logo, compound=LEFT)

solid_frame.pack(fill=BOTH, expand=True)
solid_logo = PhotoImage(file=resource_path('res/Spec_21.png'))
notebook.add(solid_frame, text="SolidWorks", image=solid_logo, compound=LEFT)


### Цвета вкладок

# Stiles
Mybackground_1 = "#8c968e"
Mybackground_2="#a2f2bc"
# фон для вкладки проекта
MyRed = "#fbf4c0"
# фон для вкладки солида
MyYellow = "#ea400a"
test = Style()
#тема для вкладки проекта
test.theme_create("my_tables_1", parent="alt", settings={
    "TFrame": {"configure": {"background": Mybackground_1}},
    "TNotebook": {"configure": {"tabmargins": [2, 0, 2, 0]}},
    "TNotebook": {
        "configure": {"background": Mybackground_1}},
    "TNotebook.Tab": {
         "configure": {"padding": [80, 1], "background": MyRed},
         "map": {"background": [("selected", MyYellow)]}}})

#тема для вкладки солида
test.theme_create("my_tables_2", parent="alt", settings={
    "TFrame": {"configure": {"background": Mybackground_2}},
    "TNotebook": {"configure": {"tabmargins": [2, 0, 2, 0]}},
    "TNotebook": {
        "configure": {"background": Mybackground_2}},
    "TNotebook.Tab": {
         "configure": {"padding": [80, 1], "background": MyRed},
         "map": {"background": [("selected", MyYellow)]}}})

test.theme_use("my_tables_1")

###


# обработка события смены вкладок, вызывается функция в которой меняем цвет фона tab_index
root.bind ("<Expose>", tab_index)


#Обработка нажатий на кнопки открыть поля шапки актов
project_frame.head_button.bind ("<ButtonRelease-1>", open_head)
solid_frame.head_button.bind ("<ButtonRelease-1>", open_head)

# обработка нажатия на кнопку спецификация проекта
project_frame.spec_button.bind ("<ButtonRelease-1>", open_table_project)
# обработка нажатия на кнопку создать все акты проекта сразу
project_frame.all_acts_button.bind("<ButtonRelease-1>", safe_all_acts)
# обработка нажатия на кнопку создать акт проекта
project_frame.act_button.bind("<ButtonRelease-1>", safe_act_project)

# обработка нажатия на кнопку спецификация cолида
solid_frame.spec_button.bind ("<ButtonRelease-1>", open_table_solid)
# обработка нажатия на кнопку создать акт из солида
solid_frame.act_button.bind("<ButtonRelease-1>", safe_act_solid)
# обработка нажатия на кнопку создать все акты проекта сразу
solid_frame.all_acts_button.bind("<ButtonRelease-1>", safe_all_acts_solid)

# Загрузка шаблона
document_spec = DocxTemplate(resource_path('res\Шаблон.docx'))
document_solid = DocxTemplate(resource_path('res\Шаблон_solid.docx'))
all_safe_enabled=0
lst_xl=[]
dir_path=''
tube_list=[]
pad_list=[]
support_list=[]
elements_list=[]


root.mainloop()