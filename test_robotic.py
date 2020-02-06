from win32com.client import Dispatch
from test_selenium import start_selenium_for_fssp, start_selenium_for_sudrf
import datetime
import os

local_dir = r'D:\project\robotic_for_MTS_bank'

# Функция проверки людей на площадке http://fssprus.ru/
def check_peoples_for_fssp():
    app_excel = Dispatch("Excel.Application")
    app_excel.Visible = True
    # Открываем эксель файл
    wb = app_excel.Workbooks.Open(local_dir + r'\test_fssp.xlsx')
    # Выбираем активный лист
    sheet = wb.ActiveSheet
    # перебираем все строки
    for index, row in enumerate(sheet.UsedRange.Rows()):
        # пропускаем первую строк, т.к. она не содержит данных, а только заголовки столбиков
        if index == 0:
            continue
        # разбиваем данные строки на части
        last_name = row[0]
        first_name = row[1]
        middle_mane = row[2]
        birthday = row[3]
        birthday = birthday.strftime('%d.%m.%Y')
        # Отправляем в функцию данные из строки и ожидаем либо True, если ничего нет.
        # либо список с данными, которые позже запишем
        result = start_selenium_for_fssp(last_name, first_name, middle_mane, birthday)

        # Если данных нет, то запишем в основной эксель файл, что исполнительных производств по этому физ.лицу нет.
        if result is True:
            sheet.Cells(index + 1, 5).Value = 'Нет'
        # Если данные есть, то создадим новый эксель файл, название файла будет содержать ФИО
        # а внутри будет храниться данные по исполнительным производствам
        else:
            # Так же в основной файл запишем информацию о том что исполнительные производства(И.П.) есть
            sheet.Cells(index + 1, 5).Value = 'Есть'
            filename = last_name + ' ' + first_name[0] + '.' + middle_mane[0]
            # Создадим новый эксель файл
            local_wb = app_excel.Workbooks.Add()
            local_sheet = local_wb.ActiveSheet
            # С помощью цикла запишем данные о И.П.
            for index, res in enumerate(result):
                for i, xx in enumerate(res):
                    local_sheet.Cells(index + 1, i + 1).Value = xx
            # Проверяем, если ли уже файл по этому человеку, если есть то удалим и сохраним новый
            if os.path.isfile(local_dir + r'\{} И.П.xlsx'.format(filename)):
                os.remove(local_dir + r'\{} И.П.xlsx'.format(filename))
            # сохраняем файл
            local_wb.SaveAs(local_dir + r'\{} И.П.xlsx'.format(filename))
            # закрываем эйсель файл
            local_wb.Close()
    # сохраняем файл
    wb.Save()
    # закрываем эйсель файл
    wb.Close()

# Функция проверки людей на площадке http://sudrf.ru/
def check_peoples_for_sudrf():
    app_excel = Dispatch("Excel.Application")
    app_excel.Visible = True
    # Открываем эксель файл
    wb = app_excel.Workbooks.Open(local_dir + r'\test_sudrf.xlsx')
    # Выбираем активный лист
    sheet = wb.ActiveSheet

    # перебираем все строки
    for index, row in enumerate(sheet.UsedRange.Rows()):
        # пропускаем первую строк, т.к. она не содержит данных, а только заголовки столбиков
        if index == 0:
            continue
        # разбиваем данные строки на части
        last_name = row[0]
        first_name = row[1]
        middle_mane = row[2]
        # Отправляем в функцию данные из строки и ожидаем либо True, если ничего нет.
        # либо список с данными, которые позже запишем
        result = start_selenium_for_sudrf(last_name, first_name, middle_mane)

        # Если данных нет, то запишем в основной эксель файл, что судебных дел по этому физ.лицу нет.
        if result is True:
            sheet.Cells(index + 1, 4).Value = 'Нет'
        # Если данные есть, то создадим новый эксель файл, название файла будет содержать ФИО
        # а внутри будет храниться данные о судебных делах
        else:
            # Так же в основной файл запишем информацию о том что судебные дела(С.Д.) есть
            sheet.Cells(index + 1, 4).Value = 'Есть'
            filename = last_name + ' ' + first_name[0] + '.' + middle_mane[0]
            # Создадим новый эксель файл
            local_wb = app_excel.Workbooks.Add()
            local_sheet = local_wb.ActiveSheet
            # С помощью цикла запишем данные о С.Д.
            for index, res in enumerate(result):
                for i, xx in enumerate(res):
                    local_sheet.Cells(index + 1, i + 1).Value = xx
            # Проверяем, если ли уже файл по этому человеку, если есть то удалим и сохраним новый
            if os.path.isfile(local_dir + r'\{} С.Д.xlsx'.format(filename)):
                os.remove(local_dir + r'\{} С.Д.xlsx'.format(filename))
            # сохраняем файл
            local_wb.SaveAs(local_dir + r'\{} С.Д.xlsx'.format(filename))
            # закрываем эйсель файл
            local_wb.Close()
    # сохраняем фай
    wb.Save()
    # закрываем эйсель файл
    wb.Close()



if __name__ == '__main__':
    check_peoples_for_sudrf()
    check_peoples_for_fssp()
