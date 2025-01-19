from tkinter import *
import tkinter.ttk as ttk
from tkinter import font
from tkinter.ttk import Combobox
from tkinter.ttk import Style




class SpecificFrame(ttk.Frame):

    def __init__(self,  *args, **kwargs):
        super().__init__(*args, **kwargs)

        font1 = font.Font(family= "Times New Roman", size=11, weight="bold", slant="roman", underline=False, overstrike=False)
        font2 = font.Font(family= "Times New Roman", size=11, weight="normal", slant="roman", underline=False,
                          overstrike=False)

        ### Цвета вкладок

        # Stiles
        Mybackground = "#a2f2bc"
        MyGreen = "#fbf4c0"
        MyYellow = "#ea400a"
        test = Style()
        test.theme_create("my_tables", parent="alt", settings={
            "TFrame": {"configure": {"background": Mybackground}},
            "TNotebook": {"configure": {"tabmargins": [2, 0, 2, 0]}},
            "TNotebook": {
                "configure": {"background": Mybackground}},
            "TNotebook.Tab": {
                "configure": {"padding": [80, 1], "background": MyGreen},
                "map": {"background": [("selected", MyYellow)]}}})
        test.theme_use("my_tables")

        ###

        # Текстовая метка
        self.name_form = Label(self, text='Заполните данные шапки акта', font=("Arial", 11, "bold"))
        self.name_form.place(x=20, y=30)

        # привязываем переменную name_station к полю ввода названия установки
        self.name_station = StringVar()

        # Поле ввода текста
        self.station = Entry(self, font=font1, textvariable=self.name_station)
        self.station.place(x=20, y=60, width=650)
        self.station.insert(0, 'Введите название установки')

        # привязываем переменную number_cal к полю ввода номер расчета
        self.number_calc = StringVar()

        self.calc = Entry(self, font=font1, textvariable=self.number_calc)
        self.calc.place(x=20, y=100, width=650)
        self.calc.insert(0, 'Введите номер расчета')

        # привязываем переменную company_name к полю ввода название компании
        self.company_name = StringVar()

        self.company = Entry(self, font=font2, textvariable=self.company_name)
        self.company.place(x=20, y=140, width=650)
        self.company.insert(0, 'Введите название компании')

        # привязываем переменную obj_name к полю ввода название обьекта
        self.obj_name = StringVar()

        self.obj = Entry(self, font=font2, textvariable=self.obj_name)
        self.obj.place(x=20, y=180, width=650)
        self.obj.insert(0, 'Введите название обьекта')

        # привязываем переменную addr_obj к полю ввода адресс обьекта
        self.addr_obj = StringVar()

        self.address = Entry(self, font=font2, textvariable=self.addr_obj)
        self.address.place(x=20, y=220, width=650)
        self.address.insert(0, 'Введите название адреса')

        # привязываем переменную dt к полю ввода дата
        self.dt = StringVar()

        self.data = Entry(self, font=font2, textvariable=self.dt)
        self.data.place(x=20, y=260, width=650)
        self.data.insert(0, 'Введите дату')

        self.acttype = Label(self, text='Выберите тип акта', font=("Arial", 11, "bold"))
        self.acttype.place(x=20, y=300)

        self.type_acts = ['входного контроля основного оборудования', 'входного контроля арматуры',
                     'входного контроля оборудования КИПиА', 'установки основного оборудования', 'монтажа арматуры']

        # по умолчанию будет выбран первый элемент из languages
        self.type_var = StringVar(value=self.type_acts[0])

        # Ниспадающий список
        self.type_choies = Combobox(self, textvariable=self.type_var, values=self.type_acts, state="readonly")
        self.type_choies.place(x=20, y=330, width=350)

        # Кнопка открытия спецификации
        self.spec_btn=IntVar()
        self.spec_button = Button(self, text='Открыть спец', font=("Arial", 12, "bold"))
        self.spec_button.place(x=400, y=20)

        # Кнопка открытия файла xlsx c полями шапки

        self.head_button = Button(self, text='Поля шапки', font=("Arial", 12, "bold"))
        self.head_button.place(x=560, y=20)

        # Кнопка создания актов

        self.act_button = Button(self, text='Создать Акт', font=("Arial", 12, "bold"))
        self.act_button.place(x=560, y=320)

        # Кнопка сохранения всех актов сразу

        self.all_acts_button = Button(self, text='Все акты сразу', font=("Arial", 12, "bold"))
        self.all_acts_button.place(x=560, y=360)


    def open_table(self)->int:
        self.spec_btn.set(1)
        return self.spec_btn.get()

class SolidFrame(ttk.Frame):

    def __init__(self,  *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Стили
        font1 = font.Font(family="Times New Roman", size=11, weight="bold", slant="roman", underline=False,
                          overstrike=False)
        font2 = font.Font(family="Times New Roman", size=11, weight="normal", slant="roman", underline=False,
                          overstrike=False)

        # Текстовая метка
        self.name_form = Label(self, text='Заполните данные шапки акта', font=("Arial", 11, "bold"))
        self.name_form.place(x=20, y=30)

        # привязываем переменную name_station к полю ввода названия установки
        self.name_station = StringVar()

        # Поле ввода текста
        self.station = Entry(self, font=font1, textvariable=self.name_station)
        self.station.place(x=20, y=60, width=650)
        self.station.insert(0, 'Введите название установки')

        # привязываем переменную number_cal к полю ввода номер расчета
        self.number_calc = StringVar()

        self.calc = Entry(self, font=font1, textvariable=self.number_calc)
        self.calc.place(x=20, y=100, width=650)
        self.calc.insert(0, 'Введите номер расчета')

        # привязываем переменную company_name к полю ввода название компании
        self.company_name = StringVar()

        self.company = Entry(self, font=font2, textvariable=self.company_name)
        self.company.place(x=20, y=140, width=650)
        self.company.insert(0, 'Введите название компании')

        # привязываем переменную obj_name к полю ввода название обьекта
        self.obj_name = StringVar()

        self.obj = Entry(self, font=font2, textvariable=self.obj_name)
        self.obj.place(x=20, y=180, width=650)
        self.obj.insert(0, 'Введите название обьекта')

        # привязываем переменную addr_obj к полю ввода адресс обьекта
        self.addr_obj = StringVar()

        self.address = Entry(self, font=font2, textvariable=self.addr_obj)
        self.address.place(x=20, y=220, width=650)
        self.address.insert(0, 'Введите название адреса')

        # привязываем переменную dt к полю ввода дата
        self.dt = StringVar()

        self.data = Entry(self, font=font2, textvariable=self.dt)
        self.data.place(x=20, y=260, width=650)
        self.data.insert(0, 'Введите дату')

        self.acttype = Label(self, text='Выберите тип акта', font=("Arial", 11, "bold"))
        self.acttype.place(x=20, y=300)

        self.type_acts = ['входного контроля элементов трубопровода', 'входного контроля материалов',
                     'проверки чистоты труб', 'монтажа опор под трубопроводы', 'монтажа трубопроводов']

        # по умолчанию будет выбран первый элемент из type_acts
        self.type_var = StringVar(value=self.type_acts[0])

        # Ниспадающий список
        self.type_choies = Combobox(self, textvariable=self.type_var, values=self.type_acts, state="readonly")
        self.type_choies.place(x=20, y=330, width=350)


        # Кнопка открытия спецификации
        self.spec_button = Button(self, text='Открыть спец', font=("Arial", 12, "bold"))
        self.spec_button.place(x=400, y=20)


        # Кнопка открытия файла xlsx c полями шапки
        # self.head_button = Button(text='Поля шапки', command=self.open_head, font=("Arial", 12, "bold"))
        self.head_button = Button(self, text='Поля шапки', font=("Arial", 12, "bold"))
        self.head_button.place(x=560, y=20)



        # Кнопка создания актов
        # self.act_button = Button(text='Создать Акт', command=self.safe_act, font=("Arial", 12, "bold"))
        self.act_button = Button(self, text='Создать Акт', font=("Arial", 12, "bold"))
        self.act_button.place(x=560, y=320)

        # Кнопка сохранения всех актов сразу
        # self.all_acts_button = Button(text='Все акты сразу', command=self.safe_all_acts, font=("Arial", 12, "bold"))
        self.all_acts_button = Button(self, text='Все акты сразу', font=("Arial", 12, "bold"))
        self.all_acts_button.place(x=560, y=360)



