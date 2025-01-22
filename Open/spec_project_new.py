import pandas as pd

# Функция обработки спецификации проекта по новой форме
#Функция открытия xlsx файла с помощью диалогового окна выбора файла в системе
def spec_parce_new(path:str) ->list:

    # Если файл выбран, то данные из листа с названием Table 1 читаются датафрейм пандас
    if path !="":
        df = pd.read_excel(path, skiprows=1)
        #print(df.head())

        # удаляем лишние столбцы из исходного датафрейма
        df = df.drop(df.columns[[3, 4, 5, 6, 7, 8, 9, 11, 12, 13, 14, 15, 16, 17, 19, 21, 22, 23, 24]], axis=1)
        #print(df.head())
        # cоздаем список из названий столбцов полученного выше датафрейм
        columns_names_df = df.columns.tolist()
        #print(columns_names_df)
        # удаляем все строки которые имеют Nan в третьем стоблце датафрейм
        df_clean = df.dropna(subset=columns_names_df[2])

        # удаление строк из датарейма если во втором столбце есть фраза "наименование" при
        # приведении в нижний регистр
        df_clean = df_clean.drop(df_clean[df_clean[df.columns[1]].str.lower().str.contains('наименование')].index)
        #print(df_clean.shape[0])
        # создаем список для термометров TI, манометров PI
        list_of_tipi = ['TI', 'PI']

        # создаем датафрейм df_tipi в котором только манометры, термометры
        df_tipi = df_clean[df_clean[columns_names_df[0]].isin(list_of_tipi)]
        # cоздаем датафрейм df_unname где все строки в первом столбце где номер позиции не имеют его
        df_unnamed = df_clean[df_clean[columns_names_df[0]].isnull()]


        # обьединяем датафрейм манометров, термометров и не имеющих номера позиции
        df_tipi_nul = pd.concat([df_tipi, df_unnamed], axis=0)  # объединение по вертикали


        # удаляем пробелы в строках 1 и 2 столбцах
        df_tipi_nul[columns_names_df[1]] = df_tipi_nul[columns_names_df[1]].str.strip()
        df_tipi_nul[columns_names_df[2]] = df_tipi_nul[columns_names_df[2]].str.strip()
        # удалить пробелы из всех строк датафрейм
        df_tipi_nul = df_tipi_nul.map(lambda x: x.strip() if isinstance(x, str) else x)

        # создаем пустой датафрейм с количеством стобцов датафрейма df_cleann
        df_tipinul_new = pd.DataFrame(columns=columns_names_df)
        # добавим в него 1 строку с числами
        df_tipinul_new.loc[0] = range(0, len(columns_names_df))

        # список уникальных значений в 2 третьем столбце df_tipinul
        uniq_val_tipinul = df_tipi_nul[columns_names_df[1]].unique()

        #print(uniq_val_tipinul)
        #print(df_tipinul_new.shape[0])
        #print(df_tipi_nul.shape[0])

        # проходимся по циклу через все уникальные значения 1 столбца датафрейма df_tipinul
        for i in range(0, len(uniq_val_tipinul)):
            df_tipinul_i = df_tipi_nul[df_tipi_nul[columns_names_df[1]] == uniq_val_tipinul[i]]

            # список уникальных значений в 3 третьем столбце для каждого уникального во 2 столбце
            uniq_val_tipinul_i = df_tipinul_i[columns_names_df[2]].unique()
            #print(uniq_val_tipinul_i)
            # проходимся по циклу через все уникальные значения в третьем столбце каждого уникального
            # во втором
            for j in range(0, len(uniq_val_tipinul_i)):
                df_tipinul_j = df_tipinul_i[df_tipinul_i[columns_names_df[2]] == uniq_val_tipinul_i[j]]

                # cчитаем сумму  в 5 столбце полученного датафрейма df_tipinul_j

                sum_j = df_tipinul_j[columns_names_df[4]].sum(axis=0)
                #print(sum_j)
                # удаляем дубликаты то есть по сути все строки кроме 1 строки в df_tube_j
                df_tipinul_drop_j = df_tipinul_j.drop_duplicates(subset=[columns_names_df[2]])


                # заносим в первую и единственную строку в 5 столбец значение sum_i
                df_tipinul_drop_j.iat[0, 4] = sum_j
                #print(df_tipi_nul_j.iat[0, 4])


                df_tipinul_new = pd.concat([df_tipinul_new, df_tipinul_drop_j], axis=0)  # объединение по вертикали


        #print(df_tipinul_new.shape[1])
        # удаляем первую строку из датафрейма df_tipinul_new
        df_tipinul_new = df_tipinul_new.iloc[1:]
        #print(df_tipinul_new.shape[1])

        # удалить пробелы из всех строк датафрейм df_clean
        df_clean = df_clean.map(lambda x: x.strip() if isinstance(x, str) else x)
        #print(df_clean.shape[0])
        #print(df_tipinul_new.shape[0])

        # датафрейм df_withnumber содержит строки df_clean которых нет в df_tipinul_new
        mask = df_clean.isin(df_tipi_nul.to_dict(orient='list')).all(axis=1)
        df_withnumber = df_clean[~mask]
        #print(df_withnumber.shape[0])

        # обьединяем датафрейм труб и элементов трубопроводов
        df_new = pd.concat([df_withnumber, df_tipinul_new], axis=0)  # объединение по вертикали

        # заменяем nan в первом столбце на NN
        df_new = df_new.fillna(' ')

        # из датафрейма df создаем двумерный массив
        new_arr = df_new.to_numpy()

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