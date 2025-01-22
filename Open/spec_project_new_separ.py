import pandas as pd

# Функция обработки спецификации проекта по новой форме
#Функция открытия xlsx файла с помощью диалогового окна выбора файла в системе
def spec_parce_new_separ(path:str, k:int=0) ->list:

    # Если файл выбран, то данные из листа с названием Table 1 читаются датафрейм пандас
    if path !="":
        df = pd.read_excel(path)
        #print(df.head())

        # удаляем лишние столбцы из исходного датафрейма
        df = df.drop(df.columns[[3, 4, 5, 6, 7, 8, 9, 11, 12, 13, 14, 15, 16, 17, 19, 21, 22, 23, 24]], axis=1)
        #print(df.head())
        # cоздаем список из названий столбцов полученного выше датафрейм
        columns_names_df = df.columns.tolist()
        #print(columns_names_df)
        # удаляем все строки которые имеют Nan во втором стоблце датафрейм
        df_clean = df.dropna(subset=columns_names_df[1])
        # удаляем все строки которые имеют Nan во 4 стоблце датафрейм
        df_clean = df.dropna(subset=columns_names_df[3])

        # удаление строк из датарейма если во втором столбце есть фраза "наименование" при
        # приведении в нижний регистр
        #df_clean = df_clean.drop(df_clean[df_clean[df.columns[1]].str.lower().str.contains('наименование')].index)
        #print(df_clean.shape[0])

        # сбрасываем индексы датафрейма и обновляем их
        df_clean.reset_index(drop=True, inplace=True)

        #print(df_clean.shape[0])

        # находим строки где в первом столбце есть наименование
        ind=df_clean[df_clean[df.columns[1]].str.lower().str.contains('наименование')].index
        #print(ind)

        #создаем пустой список
        list_of_df=list()
        list_of_df_new=list()


        #for i in range(len(ind)-1):
            #list_of_df_new[i]=pd.DataFrame(columns=columns_names_df)


        # заполняем список выше копиями df_clean и выбираем в каждой копии строки с индексами найденными в ind
        for i in range(len(ind)):
            list_of_df.append(df_clean.copy())

            list_of_df[i] = list_of_df[i].iloc[ind[i]:]
            #print(list_of_df[i].shape[0])
            #print('Edfsf')

        # выделяем в каждый датафрейм list_of_df_new количество строк только для каждого блока спецификации
        for i in range(len(list_of_df)-1):
            # датафрейм list_of_df_new содержит строки list_of_df[i] которых нет в list_of_df[i+1
            mask = list_of_df[i].isin(list_of_df[i+1].to_dict(orient='list')).all(axis=1)
            list_of_df_new.append(list_of_df[i][~mask])
            #print(list_of_df_new[i].shape[0])

        list_of_df_new.append(list_of_df[len(list_of_df)-1])
        #print(list_of_df_new[len(list_of_df)-1].shape[0])

        for i in range(len(list_of_df_new)):
            # удалить пробелы из всех строк датафреймоф list_of_df_new
            list_of_df_new[i] = list_of_df_new[i].map(lambda x: x.strip() if isinstance(x, str) else x)
            #print(df_clean.shape[0])
            #print(df_tipinul_new.shape[0])

            # заменяем nan в первом столбце на пробел
            list_of_df_new[i] = list_of_df_new[i].fillna(' ')

        #Проеферяем если к>len(list_of_df_new) то делаем его равным len(list_of_df_new)
        if k>len(list_of_df_new):
            k=len(list_of_df_new)

        # из датафрейма df создаем двумерный массив
        new_arr = list_of_df_new[k].to_numpy()

        # создаем список
        new_list = new_arr.tolist()


        return new_list
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