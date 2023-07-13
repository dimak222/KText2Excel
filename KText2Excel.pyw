#-------------------------------------------------------------------------------
# Author:      dimak222
#
# Created:     05.07.2023
# Copyright:   (c) dimak222 2023
# Licence:     No
#-------------------------------------------------------------------------------

title = "KText2Excel"
ver = "v1.0.0.0"

#------------------------------Настройки!---------------------------------------

dict_settings = {
"check_update" : [True, '# проверять обновление программы ("True" - да; "False" или "" - нет)'],
"beta" : [False, '# скачивать бета версии программы ("True" - да; "False" или "" - нет)'],
"unselect" : [True, '# снимать выделение записанного текста ("True" - да; "False" или "" - нет)'],
}

#------------------------------Импорт модулей-----------------------------------

import psutil # модуль вывода запущеных процессов
import os # работа с файовой системой
from sys import exit # для выхода из приложения без ошибки

from threading import Thread # библиотека потоков
import tkinter as tk # модуль окон
import tkinter.messagebox as mb # окно с сообщением

import pythoncom # модуль для запуска без IDLE
from win32com.client import Dispatch, gencache # библиотека API Windows

from pynput import mouse
from pynput import keyboard

import time # модуль времени

#-------------------------------------------------------------------------------

def DoubleExe(): # проверка на уже запущеное приложение

    global program_directory # значение делаем глобальным

    list = [] # список найденых названий программы

    filename = psutil.Process().name() # имя запущеного файла
    filename2 = title + ".exe" # имя запущеного файла

    if filename == "python.exe" or filename == "pythonw.exe": # если программа запущена в IDE/консоли
        pass # пропустить

    else: # запущено не в IDE/консоли

        for process in psutil.process_iter(): # перебор всех процессов

            try: # попытаться узнать имя процесса
                proc_name = process.name() # имя процесса

            except psutil.NoSuchProcess: # в случае ошибки
                pass # пропускаем

            else: # если есть имя
                if proc_name == filename or proc_name == filename2: # сравниваем имя
                    list.append(process.cwd()) # добавляем в список название программы
                    if len(list) > 2: # если найдено больше двух названий программы (два процесса)
                        Message("Приложение уже запущено!") # сообщение, поверх всех окон и с автоматическим закрытием
                        exit() # выходим из программы

    if list == []: # если нет найденых названий программы
        program_directory = os.path.dirname(os.path.abspath(__file__)) # путь рядом с файлом

    else: # если путь найден
        program_directory = os.path.dirname(psutil.Process().exe()) # директория файла

def Resource_path(relative_path): # для сохранения картинки внутри exe файла

    try: # попытаться определить путь к папке
        base_path = sys._MEIPASS # путь к временной папки PyInstaller

    except Exception: # если ошибка
        base_path = os.path.abspath(".") # абсолютный путь

    return os.path.join(base_path, relative_path) # объеденяем и возващаем полный путь

def Message(text = "Ошибка!", counter = 4): # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

    def Message_Thread(text, counter): # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

        if counter == 0: # время до закрытия окна (если 0)
            counter = 1 # закрытие через 1 сек
        window_msg = tk.Tk() # создание окна
        try: # попытаться использовать значёк
            window_msg.iconbitmap(default = Resource_path("cat.ico")) # значёк программы
        except: # если ошибка
            pass # пропустить
        window_msg.attributes("-topmost", True) # окно поверх всех окон
        window_msg.withdraw() # скрываем окно "невидимое"
        time = counter * 1000 # время в милисекундах
        window_msg.after(time, window_msg.destroy) # закрытие окна через n милисекунд
        if mb.showinfo(title, text, parent = window_msg) == "": # информационное окно закрытое по времени
            pass # пропустить
        else: # если не закрыто по времени
            window_msg.destroy() # окно закрыто по кнопке
        window_msg.mainloop() # отображение окна

    msg_th = Thread(target = Message_Thread, args = (text, counter)) # запуск окна в отдельном потоке
    msg_th.start() # запуск потока

    msg_th.join() # ждать завершения процесса, иначе может закрыться следующие окно

def ExcelAPI(): # подключение API Excel (https://memotut.com/en/150745ae0cc17cb5c866/)

    global wb # значение делаем глобальным
    global ws # значение делаем глобальным
    global max_col # значение делаем глобальным

    xlUp      = -4162 # константа Excel
    xlDown    = -4121 # константа Excel
    xlToLeft  = -4159 # константа Excel
    xlToRight = -4161 # константа ExceL

    xlCenter = -4108 # константа Excel

    def SettingsExcel(): # настройки считанные с Excel

        try: # попытаться открыть файл Excel

            ws_name = "Настройки" # название листа
            ws = wb.Worksheets(ws_name) # выбираем лист

            iCountRows = ws.Rows.Count # максимальное число строк в Excel файле
            iCells = ws.Cells(iCountRows, 1) # все ячейки в строке
            iCellsEnd = iCells.End(xlUp) # последняя строка со значением
            max_row = iCellsEnd.Row # максимальное кол-во строк по 1-й колонке

            for parameter in ws.Range(f"A1:C{max_row}").Value: # считывание значений из Excel

                if parameter[1] != None: # если параметр не пустой

                    if str(parameter[1]).find("True") != -1: # если есть параметр со словом True, обрабатываем его
                        dict_settings[parameter[0]] = [True, parameter[2]] # словарь параметров

                    elif str(parameter[1]).find("False") != -1 or parameter[1].strip() == "": # если есть параметр со словом False или "", обрабатываем его
                        dict_settings[parameter[0]] = [False, parameter[2]] # словарь параметров

                    elif str(parameter[1]).find(";") != -1: # если есть параметр с ";", обрабатываем его
                        parameter[1] = parameter[1].split(";") # разделяем параметр по ";", создаёться список
                        dict_settings[parameter[0]] = [parameter[1], parameter[2]] # словарь параметров

        except: # если лист Excel не найден
            Message("В Excel не найден лист \"Настройки\"\nИспользуються настройки по умолчанию") # сообщение, поверх всех окон с автоматическим закрытием

    def ClearRows(ws): # очистка строк если уже заполнен Excel

        global start_row # значение делаем глобальным

        iCountRows = ws.Rows.Count # максимальное число строк в Excel файле
        iCells = ws.Cells(iCountRows, 1) # все ячейки в строке
        iCellsEnd = iCells.End(xlUp) # последняя строка со значением
        max_row = iCellsEnd.Row # максимальное кол-во строк по 1-й колонке

        if max_row > 1: # если строк больше чем одна

            if AskYesNo("Продолжить заполнение строк?\n \"Нет\" - очистить строки"): # вопросительное сообщение, поверх всех окон

                start_row = max_row + 1 # строка с которой начинать заполнение

            else: # если не продолжать заполнение

                try: # попытаться выделить ячейки Excel

                    iUsedRange = ws.UsedRange # используемые колонки

                    iUsedRange.GetOffset(RowOffset = 1, ColumnOffset = 0).Delete() # очистка записанных данных

                    start_row = 2 # строка с которой начинать заполнение

                except: # если лист Excel не найден

                    Message("Ошибка записи Excel!\nВозможно, Excel закрыт или ячейка в режиме редактирования!") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)
                    exit() # выходим из программы

        else: # нет записанных строк

            start_row = 2 # строка с которой начинать заполнение

            Message("Программа запущена!", 1) # сообщение, поверх всех окон с автоматическим закрытие

    def CreateExcel(): # создание нового листа Excel

        wb = Excel.Workbooks.Add() # создать книгу
        ws = wb.Worksheets.Add() # создать лист
        ws = wb.Worksheets(1) # выбираем первый лист (создается автоматом)
        ws.Name = "v1.0" # переимненовать лист

        ws.Range("A1:D1").Value = ("Тег кабеля", "Откуда", "Куда", "Маркировка кабеля") # заголовки
        ws.Range("A1:D1").ColumnWidth = (11, 11, 11, 20) # ширина колонок
        ws.Range("A1:D1").HorizontalAlignment = xlCenter # выравнивание текста в колонках

        ws.Cells.Font.Name = "Times New Roman" # стиль текста всех ячеек
        ws.Cells.Font.Size = 12 # высота шрифта всех ячеек

        CreateSettingsExcel(wb) # создать лист с настройками в Excel

        try: # попробовать сохранить файл
            wb.SaveAs(name_txt_file) # сохранить файл Excel

        except: # не удалось сохранить

            Message("Возможно, открыт файл с тем же именем, закройте его и запустите программу повторно!", 8) # сообщение, поверх всех окон с автоматическим закрытием
            exit() # выходим из программы

        Message("Введите необходимое кол-во заголовков в первых\nстроках Excel и запустите программу заново!", 8) # сообщение, поверх всех окон с автоматическим закрытием
        exit() # выходим из программы

    def CreateSettingsExcel(wb): # сохдать лист с настройками в Excel

        ws_settings = wb.Worksheets(2) # выбираем первый лист (создается автоматом)
        ws_settings.Name = "Настройки" # переимненовать лист

        n = 1 # отсчёт строки

        for parameter, val in dict_settings.items(): # считывание значений из Excel

            ws_settings.Range(f"A{n}:C{n}").Value = (parameter, val[0], val[1]) # прописываем значения

            ws_settings.Range(f"A{n}:C{n}").EntireColumn.AutoFit() # автоширина ячейки

            n += 1 # отсчёт строки

    try: # попытаться подключиться к Excel

        Excel = Dispatch("Excel.Application") # подключение к открытому Excel

    except: # если лист Excel не найден

        Message("Excel не найден!\nУстановите или переустановите Excel!") # сообщение, поверх всех окон с автоматическим закрытием
        exit() # выходим из программы

    try: # попытаться сделать Excel видимым

        Excel.Visible = True # делаем Excel видимым

    except: # если лист Excel не найден

        Message("Excel в режиме редактирования ячейки!\nВыйдете из него или сохраните Excel!") # сообщение, поверх всех окон с автоматическим закрытие
        exit() # выходим из программы

    name_txt_file = os.path.join(program_directory, title + ".xlsx") # путь к файлу Excel

    if os.path.exists(name_txt_file): # если есть txt файл использовать его

        try: # попытаться открыть файл Excel

            wb = Excel.Workbooks.Open(name_txt_file) # открытие файла Excel

        except: # если лист Excel не найден

            Message(f"Ошибка открытия Excel файла!\n{name_txt_file}") # сообщение, поверх всех окон с автоматическим закрытием
            exit() # выходим из программы

        if wb == None: # если уже открыт такой файл
            Message(f"Уже открыт Excel файл с таким же именем, закройте его!") # сообщение, поверх всех окон с автоматическим закрытием
            exit() # выходим из программы

        SettingsExcel() # настройки считанные с Excel

        try: # попытаться открыть файл Excel

            ws_name = "v1.0" # название листа
            ws = wb.Worksheets(ws_name) # выбираем лист
            ws.Activate() # делаем его активным

        except: # если лист Excel не найден

            Message("Лист в файле Excel отсутствует или имеет несоответствующию версию!") # сообщение, поверх всех окон с автоматическим закрытием
            exit() # выходим из программы

        iCountColumns = ws.Columns.Count # максимальное число колонок в Excel файле
        iCells = ws.Cells(1, iCountColumns) # все ячейки в строке
        iCellsEnd = iCells.End(xlToLeft) # последняя колонка со значением
        max_col = iCellsEnd.Column # максимальное кол-во колонок 1-й строчки

        ClearRows(ws) # очистка строк если уже заполнен Excel

    else: # если нет Excel файла

        CreateExcel() # создание нового листа Excel

def Settings(): # присвоене значений параметров

    global check_update # значение делаем глобальным
    global beta # значение делаем глобальным
    global unselect # значение делаем глобальным

    check_update = dict_settings["check_update"][0] # опция проверки обновления программы
    beta = dict_settings["beta"][0] # опция скачивания бета версии программы
    unselect = dict_settings["unselect"][0] # опция снятия выделения записанного текста

    for parameter, val in dict_settings.items(): # для каждой строки производим обработку
        print(f"{parameter} = {val[0]}") # выводим прочитание параметры

def CheckUpdate(): # проверить обновление приложение

    global url # значение делаем глобальным

    if check_update: # если проверка обновлений включена

        try: # попытаться импортировать модуль обновления

            from Updater import Updater # импортируем модуль обновления

            if "url" not in globals(): # если нет ссылки на программу
                url = "" # нет ссылки

            Updater.Update(title, ver, beta, url, Resource_path("cat.ico")) # проверяем обновление (имя программы, версия программы, скачивать бета версию, ссылка на программу, путь к иконке)

        except SystemExit: # если закончили обновление (exit в Update)
            exit() # выходим из программы

        except: # не удалось
            pass # пропустить

def AskYesNo(text): # вопросительное сообщение, поверх всех окон

    ask = tk.Tk() # создание окна
    ask.iconbitmap(default = Resource_path("cat.ico")) # значок программы
    ask.attributes("-topmost",True) # окно поверх всех окон
    ask.withdraw() # скрываем окно "невидимое"
    ask_mb = mb.askyesno(title, text) # задаём вопрос
    ask.destroy() # закрываем окно
    ask.mainloop() # отображение окна

    return ask_mb # возвращаем результат вопроса

def KompasAPI(): # подключение API КОМПАСа

    try: # попытаться подключиться к КОМПАСу

        global KompasAPI7 # значение делаем глобальным
        global iApplication # значение делаем глобальным
        global iKompasObject # значение делаем глобальным

        KompasAPI5 = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0) # API5 КОМПАСа
        iKompasObject = Dispatch("Kompas.Application.5", None, KompasAPI5.KompasObject.CLSID) # интерфейс API КОМПАС

        KompasAPI7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0) # API7 КОМПАСа
        iApplication = Dispatch("Kompas.Application.7") # интерфейс приложения КОМПАС-3D.

        if iApplication.Visible == False: # если компас невидимый
            iApplication.Visible = True # сделать КОМПАС-3D видемым

    except: # если не получилось подключиться к КОМПАСу

        Message("КОМПАС-3D не найден!\nУстановите или переустановите КОМПАС-3D!") # сообщение, поверх всех окон с автоматическим закрытием
        exit() # выходим из программы

def ListeningCycle(): # цикл мониторинга

    global click # значение делаем глобальным

    def Click(x, y, button, pressed): # проверка нажатия кнопки мышки

        global click # значение делаем глобальным

        if button == mouse.Button.left: # ели нажата левая кнопка мышки

            if not pressed: # если кнопка отжата
                click = True # тригер нажатия кнопки мышки

    def KeyExit(key): # проверка нажатия кнопки Esc

        global stop # значение делаем глобальным

        if key == keyboard.Key.esc: # если Esc нажат
            stop = True # тригер остановки

    mouse_listener = mouse.Listener(on_click = Click) # мониторинг нажатия кнопки мышки
    mouse_listener.start() # запуск в потокеq

    keyboard_listener = keyboard.Listener(on_release = KeyExit) # мониторинг нажатия кнопки клавиатуры
    keyboard_listener.start() # запуск в потоке

    while True: # цикл определения нажатия кнопки

        if click: # тригер нажатия кнопки мышки
            TextSelection() # обработка выделенного текта
            click = False # выключаем тригер нажатия кнопки мышки

        if stop: # тригер остановки
            break # завершаем цикл

        time.sleep(0.05) # сон в секундах

    mouse_listener.stop() # останавливаем мониторинг мышки
    keyboard_listener.stop() # останавливаем мониторинг клавиатуры

def TextSelection(): # обработка выделенного текта

    def ReadText(iSelectedObject): # прочитаем выделеный текст

        if iSelectedObject: # если выделено

            try: # попытаться определить тип выделенного объекта

                if iSelectedObject.DrawingObjectType == 4: # если это текст

                    iKompasDocument2D = KompasAPI7.IKompasDocument2D(iKompasDocument) # базовый класс графических документов КОМПАС
                    iViewsAndLayersManager = iKompasDocument2D.ViewsAndLayersManager # менеджер видов и слоев документа
                    iViews = iViewsAndLayersManager.Views # коллекция видов

                    iDocument2D = iKompasObject.ActiveDocument2D() # указатель на интерфейс текущего графического документа
                    iReference = iSelectedObject.Reference # указатель объекта
                    iNumber = iDocument2D.ksGetViewNumber(iReference) # номер вида по выделеному объекту

                    iView = iViews.ViewByNumber(iNumber) # вид, заданный по номеру

                    iDrawingContainer = KompasAPI7.IDrawingContainer(iView) # интерфейс контейнера объектов вида графического докумен
                    iDrawingTexts = iDrawingContainer.DrawingTexts # указатель на интерфейс коллекции текстов на чертеже

                    iDrawingText = iDrawingTexts.DrawingText(iReference) # интерфейс текста на чертеже
                    iText = KompasAPI7.IText(iDrawingText) # интерфейс текста

                    text = iText.Str # прочитанный текст

                    Record2Excel(text) # запись в Excel

            except: # если ошибка определения
                print("Ошибка чтения из КОМПАСа!") # пропускаем

        else: # не выделено
            print("Выделите текст!")

    iKompasDocument = iApplication.ActiveDocument # получить текущий активный документ

    if iKompasDocument == None or iKompasDocument.DocumentType not in (1, 2): # если не открыт док. или не 2D док., выдать сообщение (1-чертёж; 2- фрагмент; 3-СП; 4-модель; 5-СБ; 6-текст. док.; 7-тех. СБ;)
        print("Откройте чертёж!")

    else: # если открыт 2D документ
        iKompasDocument2D1 = KompasAPI7.IKompasDocument2D1(iKompasDocument) # дополнительный интерфейс IKompasDocument2D

        iSelectionManager = iKompasDocument2D1.SelectionManager # менеджер выделенных объектов
        iSelectedObjects = iSelectionManager.SelectedObjects # массив выделенных объектов в виде SAFEARRAY | VT_DISPATCH

        if isinstance(iSelectedObjects, tuple): # если выбрано несколько объектов (кортеж объектов)
            for iSelectedObject in iSelectedObjects: # перебор всех выделеных объектов
                ReadText(iSelectedObject) # прочитаем выделеный текст

        else:  # если выбран один объект
            iSelectedObject = iSelectedObjects # если один объект
            ReadText(iSelectedObject) # прочитаем выделеный текст

        if unselect: # если снимать выделенный текст
            time.sleep(0.1) # сон в секундах
            iUnselectAll = iSelectionManager.UnselectAll() # снять выделение со всех объектов

def Record2Excel(text): # запись в Excel

    global start_row # значение делаем глобальным
    global start_col # значение делаем глобальным

    try: # попытаться записать значение в Excel

        ws.Cells(start_row, start_col).Value = text # запись значения в ячейку
        print(text)

        start_col += 1 # колонка с которой начинать заполнение

        if start_col == max_col + 1: # если дошли до последней колонки
            start_row += 1 # строка с которой начинать заполнение
            start_col = 1 # колонка с которой начинать заполнение

    except: # если ошибка определения
        Message("Ошибка записи Excel!\nВозможно, Excel закрыт или ячейка в режиме редактирования!", 2) # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

def ExcelSave(): # сохранение Excel

    try: # попытаться сохранить Excel
        wb.Save() # сохранение Excel
##        wb.Close() # закрыть книгу
##        Excel.Quit() # завершить работу Excel

    except: # не удалось сохранить
        Message("Возникла ошибка при сохранении Excel!") # сообщение, поверх всех окон с автоматическим закрытием

#-------------------------------------------------------------------------------

start_col = 1 # колонка с которой начинать заполнение
click = False # тригер нажатия кнопки мышки
stop = False # тригер остановки

DoubleExe() # проверка на уже запущеное приложени

ExcelAPI() # подключение API Excel

Settings() # присвоене значений параметров

CheckUpdate() # проверить обновление приложение

KompasAPI() # подключение API компаса

ListeningCycle() # цикл мониторинга

Message("Программа остановлена!", 2) # сообщение, поверх всех окон с автоматическим закрытием

ExcelSave() # сохранение Excel