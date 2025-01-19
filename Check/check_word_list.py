import numpy as np
from docxtpl import DocxTemplate

def fill_table_project(choice:str, col_number:int, list_for_check:list, table_for_fill:DocxTemplate) -> DocxTemplate:

#Если выбран тип акта основное оборудование
        if choice == 'входного контроля основного оборудования':

            # Создается пустой список размера как список из таблицы xlsx и потом он очищается от мусора в памяти

            f_list = np.empty((1, len(list_for_check[0]))).tolist()

            f_list.clear()

            #Производится проверка полного списка по выбранным критериям и заполняется список акта основного оборудования
            f_list += [x for x in list_for_check if 'теплообменник' in str(x).lower()]
            f_list += [x for x in list_for_check if 'насос' in str(x).lower()]
            f_list += [x for x in list_for_check if 'регулирующий' in str(x).lower()]
            f_list += [x for x in list_for_check if 'регулятор давления' in str(x).lower()]
            f_list += [x for x in list_for_check if 'частотный преобразователь' in str(x).lower()]


            # Добавляются в список столбец номеров по порядку в начало
            j = 1

            for i in range(0, len(f_list)):
                f_list[i].insert(0, j)

                j += 1



            # Заполняется таблица шаблона списком основного оборудования
            for row in f_list:
                row_cells = table_for_fill.add_row().cells
                for i in range(col_number):
                    row_cells[i].text = str(row[i])

            # Убираем первый столбец с нумерацией
            for i in range(0, len(f_list)):
                f_list[i].pop(0)

            return table_for_fill

        # Далее аналогично двум другим выборкам
        elif  choice =='установки основного оборудования':

            # Создается пустой список размера как список из таблицы xlsx и потом он очищается от мусора в памяти

            ff_list = np.empty((1, len(list_for_check[0]))).tolist()

            ff_list.clear()

            # Производится проверка полного списка по выбранным критериям и заполняется список акта основного оборудования
            ff_list += [x for x in list_for_check if 'теплообменник' in str(x).lower()]
            ff_list += [x for x in list_for_check if 'насос' in str(x).lower()]
            ff_list += [x for x in list_for_check if 'регулирующий' in str(x).lower()]
            ff_list += [x for x in list_for_check if 'регулятор давления' in str(x).lower()]
            ff_list += [x for x in list_for_check if 'частотный преобразователь' in str(x).lower()]



            # Добавляются в список столбец номеров по порядку в начало
            j = 1
            for i in range(0, len(ff_list)):
                ff_list[i].insert(0, j)
                j += 1



            # Заполняется таблица шаблона списком основного оборудования
            for row in ff_list:
                row_cells = table_for_fill.add_row().cells
                for i in range(col_number):
                    row_cells[i].text = str(row[i])

            for i in range(0, len(ff_list)):
                ff_list[i].pop(0)

            return table_for_fill

        elif choice == 'входного контроля арматуры':

            s_list = np.empty((1, len(list_for_check[0]))).tolist()
            s_list.clear()


            s_list += [x for x in list_for_check if 'фильтр' in str(x).lower()]
            s_list += [x for x in list_for_check if 'обратный' in str(x).lower()]

            s_list += [x for x in list_for_check if 'конденсатоотвод' in str(x).lower()]
            s_list += [x for x in list_for_check if 'балансировоч' in str(x).lower()]
            s_list += [x for x in list_for_check if 'прерыватель' in str(x).lower()]
            s_list += [x for x in list_for_check if 'бак' in str(x).lower()]
            s_list += [x for x in list_for_check if 'гидроаккумулятор' in str(x).lower()]
            s_list += [x for x in list_for_check if 'затвор' in str(x).lower()]
            s_list += [x for x in list_for_check if 'предохранительный' in str(x).lower()]
            s_list += [x for x in list_for_check if 'соленоидный' in str(x).lower()]
            s_list += [x for x in list_for_check if 'накип' in str(x).lower()]
            s_list += [x for x in list_for_check if 'запорный' in str(x).lower()]
            s_list += [x for x in list_for_check if 'сепаратор' in str(x).lower()]




            j = 1
            for i in range(0, len(s_list)):
                s_list[i].insert(0, j)
                j += 1


            for row in s_list:
                row_cells = table_for_fill.add_row().cells
                for i in range(col_number):
                    row_cells[i].text = str(row[i])

            for i in range(0, len(s_list)):
                s_list[i].pop(0)

            return table_for_fill

        elif choice =='монтажа арматуры':


            ss_list = np.empty((1, len(list_for_check[0]))).tolist()
            ss_list.clear()


            ss_list += [x for x in list_for_check if 'фильтр' in str(x).lower()]
            ss_list += [x for x in list_for_check if 'обратный' in str(x).lower()]

            ss_list += [x for x in list_for_check if 'конденсатоотвод' in str(x).lower()]
            ss_list += [x for x in list_for_check if 'балансировоч' in str(x).lower()]
            ss_list += [x for x in list_for_check if 'прерыватель' in str(x).lower()]
            ss_list += [x for x in list_for_check if 'бак' in str(x).lower()]
            ss_list += [x for x in list_for_check if 'гидроаккумулятор' in str(x).lower()]
            ss_list += [x for x in list_for_check if 'затвор' in str(x).lower()]
            ss_list += [x for x in list_for_check if 'предохранительный' in str(x).lower()]
            ss_list += [x for x in list_for_check if 'соленоидный' in str(x).lower()]
            ss_list += [x for x in list_for_check if 'накип' in str(x).lower()]

            ss_list += [x for x in list_for_check if 'сепаратор' in str(x).lower()]




            j = 1
            for i in range(0, len(ss_list)):
                ss_list[i].insert(0, j)
                j += 1


            for row in ss_list:
                row_cells = table_for_fill.add_row().cells
                for i in range(col_number):
                    row_cells[i].text = str(row[i])

            for i in range(0, len(ss_list)):
                ss_list[i].pop(0)

            return table_for_fill

        elif choice == 'входного контроля оборудования КИПиА':
            #создаем пустой список th_list с 1 строкой и количеством столбцов как в глобальном
            # списке lst_xl
            th_list = np.empty((1, len(list_for_check[0]))).tolist()
            # отчищаем его от мустора в памяти
            th_list.clear()
            #наполняем список по условиям нужных слов
            th_list += [x for x in list_for_check if 'манометр'  in str(x).lower()]
            th_list += [x for x in list_for_check if 'термометр' in str(x).lower()]
            th_list += [x for x in list_for_check if 'термостат ' in str(x).lower()]
            #th_list += [x for x in list_for_check if 'датчик' in str(x).lower()]


            th_list += [x for x in list_for_check if ('датчик' in str(x).lower()) and ('погруж' in str(x).lower())]
            th_list += [x for x in list_for_check if ('датчик' in str(x).lower()) and ('уровня' in str(x).lower())]
            th_list += [x for x in list_for_check if ('датчик' in str(x).lower()) and ('давления' in str(x).lower())]
            th_list += [x for x in list_for_check if ('датчик' in str(x).lower()) and ('нар' in str(x).lower())]

            th_list += [x for x in list_for_check if ('датчик' in str(x).lower()) and ('помещ' in str(x).lower())]

            th_list += [x for x in list_for_check if ('датчик' in str(x).lower()) and ('гильза' in str(x).lower())]
            th_list += [x for x in list_for_check if ('бобышка' in str(x).lower())]


            th_list += [x for x in list_for_check if 'реле' in str(x).lower()]
            th_list += [x for x in list_for_check if 'прессостат' in str(x).lower()]
            th_list += [x for x in list_for_check if 'трехходовой' in str(x).lower()]
            th_list += [x for x in list_for_check if 'одновентильный' in str(x).lower()]
            th_list += [x for x in list_for_check if 'охладитель' in str(x).lower()]
            th_list += [x for x in list_for_check if 'трубка' in str(x).lower()]
            th_list += [x for x in list_for_check if 'расходомер' in str(x).lower()]
            th_list += [x for x in list_for_check if 'счетчик' in str(x).lower()]
            th_list += [x for x in list_for_check if ('гильза' in str(x).lower()) and not ('датчик' in str(x).lower())]


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
                row_cells = table_for_fill.add_row().cells
                for i in range(col_number):
                    row_cells[i].text = str(row[i])

            for i in range(0, len(th_list)):
                th_list[i].pop(0)

            return table_for_fill
        # Если выборки не сделано, то таблица в шаблоне наполняется полным списком из спецификации
        else:

            xl_list = list_for_check.copy()

            j = 1
            for i in range(0, len(xl_list)):
                xl_list[i].insert(0, j)
                j += 1


            for row in xl_list:
                row_cells = table_for_fill.add_row().cells
                for i in range(col_number):
                    row_cells[i].text = str(row[i])

            for i in range(0, len(xl_list)):
                xl_list[i].pop(0)

            return table_for_fill