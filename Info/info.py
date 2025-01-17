from tkinter.messagebox import showinfo

# Функция вызывающее информационное окно для напоминания о выборе нужного файла спецификации
def info_spec():
    showinfo(title="Информация", message="Выберете спецификацию в формате xlsx")

# Функция вызывающее информационное окно для напоминания о выборе нужного файла спецификации
def info_head():
    showinfo(title="Информация", message="Выберете файл с полями акта в формате xlsx")

# Функция вызывающая информационное окно когда не загружена спецификация при попытке создания акта
def info_act():
    showinfo(title="Информация", message="Проверьте что спецификация загружена!")
