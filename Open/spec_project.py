import pandas as pd

#Функция открытия xlsx файла с помощью диалогового окна выбора файла в системе
def spec_parce(path:str) ->list:

    # Если файл выбран, то данные из листа с названием Table 1 читаются датафрейм пандас
    if path !="":
        #df = pd.read_excel(table_path, sheet_name='Table 1', skiprows=2)
        df = pd.read_excel(path, skiprows=2)
        # Получение количества столбцов датафрейма
        a=df.shape[1]
        # если столбцов больше 5 то остальные удаляются
        if a>5:
            #df = df.drop(df.columns[[5, a-1]], axis=1)
            df = df[[df.columns[0], df.columns[1], df.columns[2], df.columns[3], df.columns[4]]]


            #удаление строк из датарейма если в двух столбцах у них Nan
        df = df.drop(df[(df[df.columns[2]].isnull()) & (df[df.columns[3]].isnull())].index)

        #Из датафрейма удаляются все строки содержащие пустые ячейки
        #df_cleaned=df.dropna()
        # Из полученного датафрейма получаем двумерный массив средствами numpy
        #xl_arr=df_cleaned.to_numpy()
        #df.fillna(0.0)
        #print(df.columns)
        #df['Кол-во']=df['Кол-во'].astype(float)
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

        return lst_xl
    else:

       empty_list=[]
       return empty_list


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