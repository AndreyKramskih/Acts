import pandas as pd

def solid_parce(table_path:str)->list:
    # Если файл выбран, то данные из листа с названием Table 1 читаются датафрейм пандас
    if table_path != "":
        # global model_dict
        # df = pd.read_excel(table_path, sheet_name='Table 1', skiprows=2)
        df = pd.read_excel(table_path)
        # удаляем лишние столбцы из исходного датафрейма
        df = df.drop(df.columns[[0, 1, 3, 5, 6, 7, 11, 12, 15, 16, 17, 18]], axis=1)
        # cоздаем список из названий столбцов полученного выше датафрейм
        columns_names_df = df.columns.tolist()
        # удаляем все строки которые имеют Nan в первом стоблце датафрейм
        df_clean = df.dropna(subset=columns_names_df[0])
        # создаем список номеров из словаря Котенко для труб
        list_of_tube = list(range(7, 13))

        # создаем датафрейм df_tube в котором только трубы
        df_tube = df_clean[df_clean[columns_names_df[1]].isin(list_of_tube)]

        # создадим пустой датафрейм
        df_tube_new = pd.DataFrame(columns=columns_names_df)
        # добави в него 1 строку с числами
        df_tube_new.loc[0] = range(0, len(columns_names_df))

        # список уникальных значений в 2 третьем столбце df_tube
        uniq_val_tube = df_tube[columns_names_df[1]].unique()
        # проходимся по циклу через все уникальные значения 2 столбца датафрейма труб
        for i in range(0, len(uniq_val_tube)):
            df_tube_i = df_tube[df_tube[columns_names_df[1]] == uniq_val_tube[i]]
            # список уникальных значений в 3 третьем столбце для каждого уникального во 2 столбце
            uniq_val_tube_i = df_tube_i[columns_names_df[2]].unique()
            # проходимся по циклу через все уникальные значения в третьем столбце каждого уникального
            # во втором
            for j in range(0, len(uniq_val_tube_i)):
                df_tube_j = df_tube_i[df_tube_i[columns_names_df[2]] == uniq_val_tube_i[j]]
                # cчитаем сумму длинн в 6 столбце полученного датафрейма df_tube_j и оставляем 2 знака после
                # запятой
                sum_j = round(df_tube_j[columns_names_df[5]].sum(axis=0), 2)
                # удаляем дубликаты то есть по сути все строки кроме 1 строки в df_tube_j
                df_tube_j = df_tube_j.drop_duplicates(subset=[columns_names_df[1]])
                # заносим в первую и единственную строку в 6 столбец значение sum_j
                df_tube_j.iat[0, 5] = sum_j
                # в 4 столбце удаляем значение длины участка в скобках которое было в xlsx файле
                df_tube_j.iat[0, 3] = df_tube_j.iat[0, 3][:df_tube_j.iat[0, 3].find('(')]
                # в 6 столбце меняем м на п.м
                df_tube_j.iat[0, 6] = 'п.м.'

                df_tube_new = pd.concat([df_tube_new, df_tube_j], axis=0)  # объединение по вертикали

        # удаляем лишние столбцы из  датафрейма df_tube_new
        df_tube_new = df_tube_new.drop(df.columns[[1, 4]], axis=1)

        # удаляем первую строку из датафрейма df_tube_new
        df_tube_new = df_tube_new.iloc[1:]

        # из датафрейма df создаем двумерный массив
        tube_arr = df_tube_new.to_numpy()


        # создаем список
        tube_list = tube_arr.tolist()
        # print(tube_list)

        # создаем столбец сквозной нумерации
        j = 1
        for i in range(0, len(tube_list)):
            tube_list[i].insert(0, j)
            j += 1

        # номер соотвествующий прокладке
        numb_pad = 29
        # создаем датафрейм df_pad в котором только прокладки
        df_pad = df_clean[df_clean[columns_names_df[1]] == numb_pad]
        # удаляем лишние столбцы из  датафрейма df_pad
        df_pad = df_pad.drop(df.columns[[1, 5]], axis=1)
        # из датафрейма df создаем двумерный массив
        pad_arr = df_pad.to_numpy()


        # создаем список
        pad_list = pad_arr.tolist()

        # создаем столбец сквозной нумерации
        j = 1
        for i in range(0, len(pad_list)):
            pad_list[i].insert(0, j)
            j += 1

        # создаем список номеров из словаря Котенко для опор
        list_of_support = list(range(68, 70))
        list_of_support.append(6)
        list_of_profil=list(range(2, 4))

        ####
        # создаем датафрейм df_tube в котором только трубы
        df_profil = df_clean[df_clean[columns_names_df[1]].isin(list_of_profil)]

        # создадим пустой датафрейм
        df_profil_new = pd.DataFrame(columns=columns_names_df)
        # добави в него 1 строку с числами
        df_profil_new.loc[0] = range(0, len(columns_names_df))

        # список уникальных значений в 2 третьем столбце df_profil
        uniq_val_profil = df_profil[columns_names_df[1]].unique()
        # проходимся по циклу через все уникальные значения 2 столбца датафрейма профилей
        for i in range(0, len(uniq_val_profil)):
            df_profil_i = df_profil[df_profil[columns_names_df[1]] == uniq_val_profil[i]]
            # список уникальных значений в 3 третьем столбце для каждого уникального во 2 столбце
            uniq_val_profil_i = df_profil_i[columns_names_df[2]].unique()
            # проходимся по циклу через все уникальные значения в третьем столбце каждого уникального
            # во втором
            for j in range(0, len(uniq_val_profil_i)):
                df_profil_j = df_profil_i[df_profil_i[columns_names_df[2]] == uniq_val_profil_i[j]]
                # cчитаем сумму длинн в 6 столбце полученного датафрейма df_profil_j и оставляем 2 знака после
                # запятой
                sum_j = round(df_profil_j[columns_names_df[5]].sum(axis=0), 2)
                # удаляем дубликаты то есть по сути все строки кроме 1 строки в df_profil_j
                df_profil_j = df_profil_j.drop_duplicates(subset=[columns_names_df[1]])
                # заносим в первую и единственную строку в 6 столбец значение sum_j
                df_profil_j.iat[0, 5] = sum_j
                # в 4 столбце удаляем значение длины участка в скобках которое было в xlsx файле
                df_profil_j.iat[0, 3] = df_profil_j.iat[0, 3][:df_profil_j.iat[0, 3].find('(')]
                # в 6 столбце меняем м на п.м
                df_profil_j.iat[0, 6] = 'п.м.'

                df_profil_new = pd.concat([df_profil_new, df_profil_j], axis=0)  # объединение по вертикали

        # удаляем лишние столбцы из  датафрейма df_profil_new
        df_profil_new = df_profil_new.drop(df.columns[[1, 4]], axis=1)


        # удаляем первую строку из датафрейма df_profil_new
        df_profil_new = df_profil_new.iloc[1:]

        ####

        # создаем датафрейм df_support в котором только опоры и хомуты
        df_support = df_clean[df_clean[columns_names_df[1]].isin(list_of_support)]

        # удаляем лишние столбцы из  датафрейма df_support
        df_support = df_support.drop(df.columns[[1, 5]], axis=1)



        # cоздаем список из названий столбцов полученного
        columns_names_support = df_support.columns.tolist()

        #заменяем Nan на пробел
        df_support[columns_names_support[1]] =df_support[columns_names_support[1]] .fillna('Опора')
        df_support[columns_names_support[4]] = df_support[columns_names_support[4]].fillna('шт.')
        # переименовываем столбец 4 чтобы был как у труб
        df_support = df_support.rename(columns={columns_names_support[3]: columns_names_df[5]})

        # обьединяем датафрейм профилей и опор
        df_support_new = pd.concat([df_profil_new, df_support], axis=0)  # объединение по вертикали


        # из датафрейма df создаем двумерный массив
        support_arr = df_support_new.to_numpy()


        # создаем список
        support_list = support_arr.tolist()

        # создаем столбец сквозной нумерации
        j = 1
        for i in range(0, len(support_list)):
            support_list[i].insert(0, j)
            j += 1

        # создаем список номеров из словаря Котенко для материалов элементов трубопроводов
        list_of_elements = list(range(13, 29))
        list_of_elements_2 = list(range(63, 68))
        list_of_elements_3 = list(range(71, 74))
        list_of_anker=list(range(30, 36))
        # обьединяем списки
        list_of_elements.extend(list_of_elements_2)
        list_of_elements.extend(list_of_elements_3)

        #####
        # создаем датафрейм df_tube в котором только крепеж
        df_anker = df_clean[df_clean[columns_names_df[1]].isin(list_of_anker)]

        # создадим пустой датафрейм
        df_anker_new = pd.DataFrame(columns=columns_names_df)
        # добавим в него 1 строку с числами
        df_anker_new.loc[0] = range(0, len(columns_names_df))

        # cоздаем список из названий столбцов полученного
        columns_names_anker = df_anker.columns.tolist()


        # в 3 столбце удаляем значение  ду в скобках которое было в xlsx файле
        for i in range(0, df_anker.shape[0]):
            # в 4 столбце удаляем значение длины участка в скобках которое было в xlsx файле
            df_anker.iat[i, 2] = df_anker.iat[i, 2][:df_anker.iat[i, 2].find('(')]
            df_anker.iat[i, 2] = df_anker.iat[i, 2].strip('\n')
            df_anker.iat[i, 2] = df_anker.iat[i, 2].strip()

        # список уникальных значений в 2 третьем столбце df_anker
        uniq_val_anker = df_anker[columns_names_df[1]].unique()
        # проходимся по циклу через все уникальные значения 2 столбца датафрейма крепежа
        for i in range(0, len(uniq_val_anker)):
            df_anker_i = df_anker[df_anker[columns_names_df[1]] == uniq_val_anker[i]]
            # список уникальных значений в 3 третьем столбце для каждого уникального во 2 столбце
            uniq_val_anker_i = df_anker_i[columns_names_df[2]].unique()
            #print(uniq_val_anker_i)
            # проходимся по циклу через все уникальные значения в третьем столбце каждого уникального
            # во втором
            for j in range(0, len(uniq_val_anker_i)):
                df_anker_j = df_anker_i[df_anker_i[columns_names_df[2]] == uniq_val_anker_i[j]]
                # cчитаем сумму длинн в 5 столбце полученного датафрейма df_anker_j и оставляем 2 знака после
                # запятой
                #print(df_anker_j.iat[0,4])
                sum_j = round(df_anker_j[columns_names_df[4]].sum(axis=0), 2)
                #print(sum_j)
                # удаляем дубликаты то есть по сути все строки кроме 1 строки в df_anker_j
                df_anker_j = df_anker_j.drop_duplicates(subset=[columns_names_df[1]])
                #print(df_anker_j.head())
                # заносим в первую и единственную строку в 5 столбец значение sum_j
                df_anker_j.iat[0, 4] = sum_j
                #print(df_anker_j.iat[0,4])

                df_anker_new = pd.concat([df_anker_new, df_anker_j], axis=0)  # объединение по вертикали

        # удаляем лишние столбцы из  датафрейма df_anker_new
        df_anker_new = df_anker_new.drop(df.columns[[1, 5]], axis=1)
        #print(df_anker_new.head())

        # удаляем первую строку из датафрейма df_anker_new
        df_anker_new = df_anker_new.iloc[1:]

        #####



        # создаем датафрейм df_elements в котором только элементы трубопроводов
        df_elements = df_clean[df_clean[columns_names_df[1]].isin(list_of_elements)]



        # удаляем лишние столбцы из  датафрейма df_elements
        df_elements = df_elements.drop(df.columns[[1, 5]], axis=1)
        # cоздаем список из названий столбцов полученного
        columns_names_elements = df_elements.columns.tolist()



        # обьединяем датафрейм крепежа и элементов трубопроводов
        df_elements_1 = pd.concat([df_elements, df_anker_new], axis=0)  # объединение по вертикали

        # переименовываем столбец 4 чтобы был как у труб
        df_elements_1 = df_elements_1.rename(columns={columns_names_elements[3]: columns_names_df[5]})

        # обьединяем датафрейм труб и элементов трубопроводов и крепежа
        df_elements_new = pd.concat([df_tube_new, df_elements_1], axis=0)  # объединение по вертикали

        # из датафрейма df создаем двумерный массив
        elements_arr = df_elements_new.to_numpy()


        # создаем список
        elements_list = elements_arr.tolist()

        # создаем столбец сквозной нумерации
        j = 1
        for i in range(0, len(elements_list)):
            elements_list[i].insert(0, j)
            j += 1

        parse_lists=[tube_list, pad_list, support_list, elements_list]

        return parse_lists
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