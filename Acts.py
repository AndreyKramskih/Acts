from tkinter.ttk import Combobox
import numpy as np
from tkinter import *
from tkinter import filedialog
import pandas as pd
from docxtpl import DocxTemplate
from tkinter.messagebox import showinfo
from tkinter import font
import sys
import os

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

# Функция вызывающее информационное окно для напоминания о выборе нужного файла спецификации
def open_spec():
    showinfo(title="Информация", message="Выберете спецификацию в формате xlsx")
# Функция вызывающая информационное окно когда не загружена спецификация при попытке создания акта
def make_act():
    showinfo(title="Информация", message="Проверьте что спецификация загружена!")

# Функция создания акта
def safe_act():
    # Проверка если спецификация была не загружена, то нет возможности создать акт
    if len(lst_xl) <= 1:
        make_act()
        return
    global document
    #global context

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
        #Если выбран тип акта основное оборудование
        if type_choies.get() == 'входного контроля основного оборудования':

            # Создается пустой список размера как список из таблицы xlsx и потом он очищается от мусора в памяти

            f_list = np.empty((1, len(lst_xl[0]))).tolist()

            f_list.clear()

            #Производится проверка полного списка по выбранным критериям и заполняется список акта основного оборудования
            f_list += [x for x in lst_xl if 'теплообменник' in str(x).lower()]
            f_list += [x for x in lst_xl if 'насос' in str(x).lower()]
            f_list += [x for x in lst_xl if 'регулирующий' in str(x).lower()]
            f_list += [x for x in lst_xl if 'регулятор давления' in str(x).lower()]
            f_list += [x for x in lst_xl if 'частотный преобразователь' in str(x).lower()]


            # Добавляются в список столбец номеров по порядку в начало
            j = 1

            for i in range(0, len(f_list)):
                f_list[i].insert(0, j)

                j += 1



            # Заполняется таблица шаблона списком основного оборудования
            for row in f_list:
                row_cells = new_table.add_row().cells
                for i in range(cols_number):
                    row_cells[i].text = str(row[i])

            # Убираем первый столбец с нумерацией
            for i in range(0, len(f_list)):
                f_list[i].pop(0)


        # Далее аналогично двум другим выборкам
        elif type_choies.get() =='установки основного оборудования':

            # Создается пустой список размера как список из таблицы xlsx и потом он очищается от мусора в памяти

            ff_list = np.empty((1, len(lst_xl[0]))).tolist()

            ff_list.clear()

            # Производится проверка полного списка по выбранным критериям и заполняется список акта основного оборудования
            ff_list += [x for x in lst_xl if 'теплообменник' in str(x).lower()]
            ff_list += [x for x in lst_xl if 'насос' in str(x).lower()]
            ff_list += [x for x in lst_xl if 'регулирующий' in str(x).lower()]
            ff_list += [x for x in lst_xl if 'регулятор давления' in str(x).lower()]
            ff_list += [x for x in lst_xl if 'частотный преобразователь' in str(x).lower()]



            # Добавляются в список столбец номеров по порядку в начало
            j = 1
            for i in range(0, len(ff_list)):
                ff_list[i].insert(0, j)
                j += 1



            # Заполняется таблица шаблона списком основного оборудования
            for row in ff_list:
                row_cells = new_table.add_row().cells
                for i in range(cols_number):
                    row_cells[i].text = str(row[i])

            for i in range(0, len(ff_list)):
                ff_list[i].pop(0)

        elif type_choies.get() == 'входного контроля арматуры':

            s_list = np.empty((1, len(lst_xl[0]))).tolist()
            s_list.clear()


            s_list += [x for x in lst_xl if 'фильтр' in str(x).lower()]
            s_list += [x for x in lst_xl if 'обратный' in str(x).lower()]
            #s_list += [x for x in lst_xl if 'вентиль' in str(x).lower()]
            #s_list += [x for x in lst_xl if 'шаровой' in str(x).lower()]
            s_list += [x for x in lst_xl if 'конденсатоотвод' in str(x).lower()]
            s_list += [x for x in lst_xl if 'балансировоч' in str(x).lower()]
            s_list += [x for x in lst_xl if 'прерыватель' in str(x).lower()]
            s_list += [x for x in lst_xl if 'бак' in str(x).lower()]
            s_list += [x for x in lst_xl if 'гидроаккумулятор' in str(x).lower()]
            s_list += [x for x in lst_xl if 'затвор' in str(x).lower()]
            s_list += [x for x in lst_xl if 'предохранительный' in str(x).lower()]
            s_list += [x for x in lst_xl if 'соленоидный' in str(x).lower()]
            s_list += [x for x in lst_xl if 'накип' in str(x).lower()]
            s_list += [x for x in lst_xl if 'запорный' in str(x).lower()]




            j = 1
            for i in range(0, len(s_list)):
                s_list[i].insert(0, j)
                j += 1


            for row in s_list:
                row_cells = new_table.add_row().cells
                for i in range(cols_number):
                    row_cells[i].text = str(row[i])

            for i in range(0, len(s_list)):
                s_list[i].pop(0)

        elif type_choies.get() =='монтажа арматуры':


            ss_list = np.empty((1, len(lst_xl[0]))).tolist()
            ss_list.clear()


            ss_list += [x for x in lst_xl if 'фильтр' in str(x).lower()]
            ss_list += [x for x in lst_xl if 'обратный' in str(x).lower()]
            #ss_list += [x for x in lst_xl if 'вентиль' in str(x).lower()]
            #ss_list += [x for x in lst_xl if 'шаровой' in str(x).lower()]
            ss_list += [x for x in lst_xl if 'конденсатоотвод' in str(x).lower()]
            ss_list += [x for x in lst_xl if 'балансировоч' in str(x).lower()]
            ss_list += [x for x in lst_xl if 'прерыватель' in str(x).lower()]
            ss_list += [x for x in lst_xl if 'бак' in str(x).lower()]
            ss_list += [x for x in lst_xl if 'гидроаккумулятор' in str(x).lower()]
            ss_list += [x for x in lst_xl if 'затвор' in str(x).lower()]
            ss_list += [x for x in lst_xl if 'предохранительный' in str(x).lower()]
            ss_list += [x for x in lst_xl if 'соленоидный' in str(x).lower()]
            ss_list += [x for x in lst_xl if 'накип' in str(x).lower()]
            ss_list += [x for x in lst_xl if 'запорный' in str(x).lower()]




            j = 1
            for i in range(0, len(ss_list)):
                ss_list[i].insert(0, j)
                j += 1


            for row in ss_list:
                row_cells = new_table.add_row().cells
                for i in range(cols_number):
                    row_cells[i].text = str(row[i])

            for i in range(0, len(ss_list)):
                ss_list[i].pop(0)

        elif type_choies.get() == 'входного контроля оборудования КИПиА':
            #создаем пустой список th_list с 1 строкой и количеством столбцов как в глобальном
            # списке lst_xl
            th_list = np.empty((1, len(lst_xl[0]))).tolist()
            # отчищаем его от мустора в памяти
            th_list.clear()
            #наполняем список по условиям нужных слов
            th_list += [x for x in lst_xl if 'манометр'  in str(x).lower()]
            th_list += [x for x in lst_xl if 'термометр' in str(x).lower()]
            th_list += [x for x in lst_xl if 'термостат ' in str(x).lower()]
            th_list += [x for x in lst_xl if 'датчик' in str(x).lower()]
            th_list += [x for x in lst_xl if 'реле' in str(x).lower()]
            th_list += [x for x in lst_xl if 'прессостат' in str(x).lower()]
            th_list += [x for x in lst_xl if 'трехходовой' in str(x).lower()]
            th_list += [x for x in lst_xl if 'одновентильный' in str(x).lower()]
            th_list += [x for x in lst_xl if 'охладитель' in str(x).lower()]
            th_list += [x for x in lst_xl if 'трубка' in str(x).lower()]
            th_list += [x for x in lst_xl if 'расходомер' in str(x).lower()]
            th_list += [x for x in lst_xl if 'счетчик' in str(x).lower()]


            #создаем столбец сквозной нумерации
            j = 1
            for i in range(0, len(th_list)):
                th_list[i].insert(0, j)
                j += 1
            # создание копии листа
            th_list_new=th_list.copy()
            # Замена 0.0 на пробел в списке th_list_new
            for i in range(0, len(th_list_new)):
                th_list_new[i] = list(map(lambda x: " " if x =='0.0' else x, th_list_new[i]))

            # заполняем таблицу шаблона списком th_list_new
            for row in th_list_new:
                row_cells = new_table.add_row().cells
                for i in range(cols_number):
                    row_cells[i].text = str(row[i])

            for i in range(0, len(th_list)):
                th_list[i].pop(0)
        # Если выборки не сделано, то таблица в шаблоне наполняется полным списком из спецификации
        else:

            xl_list = lst_xl.copy()

            j = 1
            for i in range(0, len(xl_list)):
                xl_list[i].insert(0, j)
                j += 1


            for row in xl_list:
                row_cells = new_table.add_row().cells
                for i in range(cols_number):
                    row_cells[i].text = str(row[i])

            for i in range(0, len(xl_list)):
                xl_list[i].pop(0)

        document.save(filepath)

#Функция открытия xlsx файла с помощью диалогового окна выбора файла в системе
def open_table():
    # Информационное окно напоминающее, что нужно выбрать для открытия файл спецификации формата xlsx
    open_spec()
    table_path=filedialog.askopenfilename()

    # Если файл выбран, то данные из листа с названием Table 1 читаются датафрейм пандас
    if table_path !="":
        df = pd.read_excel(table_path, sheet_name='Table 1', skiprows=2)
        #Из датафрейма удаляются все строки содержащие пустые ячейки
        #df_cleaned=df.dropna()
        # Из полученного датафрейма получаем двумерный массив средствами numpy
        #xl_arr=df_cleaned.to_numpy()
        #df.fillna(0.0)
        #print(df.columns)
        df['Кол-во']=df['Кол-во'].astype(float)
        #из датафрейма df создаем двумерный массив
        xl_arr = df.to_numpy()

        #из двумерного массива создаем список lst строк string
        # столбцы массива помещаем в строку через разделитель *
        lst=[]
        for i in range(0, xl_arr.shape[0]):
            string = ""
            for j in range(0,xl_arr.shape[1]):
                # Проверяем на NaN если значение не равно самому себе то это NaN
                # Заменяем NaN на 0.0 в строке
                if xl_arr[i][j]==xl_arr[i][j]:
                    string += "*"+str(xl_arr[i][j]).lstrip()
                else:
                    string+="*"+"0.0"

            lst.append(string)
        # создаем вспомогательный список lst_add из списка lst
        #из которого удалям элементы с конца строки до разделителя * (количество оборудования)
        lst_add=[]
        for i in range(0, len(lst)):
            lst_add.append(lst[i][:lst[i].rfind('*')])
        # глобальный список lst_xl
        global lst_xl
        # удаляем из списка lst_add все повторяющиеся строки и создаем список lst_xl
        lst_xl=list(set(lst_add))
       # проходимся по строкам списка lst_xl
        for i in range(0, len(lst_xl)):
            # обьявляем переменную которая считает количество оборудования и обнуляем ее здесь
            s=0.0
            # сравниваем строки списка lst_xl со строками из списка lst, в которых также обрезан конец до
            # первой *
            for j in range(0,len(lst)):
                if lst_xl[i]== lst[j][:lst[j].rfind('*')]:
                   #если строки совпадают то увеличиваем s на величину обрезка приведенного к float
                   s=s+float(lst[j][(lst[j].rfind('*')+1):])
            # добавляем к строке списка lst_xl величину s преобразованную в строку через разделитель *
            # затем раздеяем строку по * и получаем список списков
            lst_xl[i]=(lst_xl[i] + "*" + str(s)).split('*')
            # удаляем в списке по строчно первый столбец и получаем
            # финальный список для обработки в типах актов
            lst_xl[i].pop(0)


"""""
#Функция создания таблицы здесь не требуется так как мы заполняем существующую в шаблоне
def create_table(document, headers, rows, style='Table Grid'):
    cols_number = len(headers)

    table = document.add_table(rows=1, cols=cols_number)
    table.style = style

    hdr_cells = table.rows[0].cells
    for i in range(cols_number):
        hdr_cells[i].text = headers[i]

    for row in rows:
        row_cells = table.add_row().cells
        for i in range(cols_number):
            row_cells[i].text = str(row[i])

    return table
"""
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

#Поле ввода текста
station =Entry(root, font=font1)
station.place(x=20, y=60, width=650)
station.insert(0,'Введите название установки')

calc =Entry(root, font=font1)
calc.place(x=20, y=100, width=650)
calc.insert(0,'Введите номер расчета')

company =Entry(root, font=font2)
company.place(x=20, y=140, width=650)
company.insert(0,'Введите название компании')

obj =Entry(root, font=font2)
obj.place(x=20, y=180, width=650)
obj.insert(0,'Введите название обьекта')

address =Entry(root, font=font2)
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

data =Entry(root, font=font2)
data.place(x=20, y=2600, width=650)
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

#Кнопка открытия спецификации
file_button=Button(text='Открыть спец', command=open_table, font=("Arial", 12, "bold"))
file_button.place(x=400, y=20)

#Кнопка создания актов
btn=Button(text='Создать Акт', command=safe_act, font=("Arial", 12, "bold"))
btn.place(x=500, y=320)


# Загрузка шаблона
document = DocxTemplate(resource_path('res\Шаблон.docx'))

lst_xl=[]
#context={}

root.mainloop()