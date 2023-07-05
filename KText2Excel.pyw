#-------------------------------------------------------------------------------
# Author:      dimak222
#
# Created:     27.02.2022
# Copyright:   (c) dimak222 2023
# Licence:     No
#-------------------------------------------------------------------------------

title = "Получение выделенного текста из чертежа и запись в Excel"
ver = "v0.1.0.0"

import win32api # библиотека API Windows
import time # модуль времени

def KompasAPI(): # подключение API компаса

    import pythoncom # модуль для запуска без IDLE
    from win32com.client import Dispatch, gencache # библиотека API Windows

    try: # попытаться подключиться к КОМПАСу

        global iKompasObject # значение делаем глобальным
        global KompasAPI7 # значение делаем глобальным
        global iApplication # значение делаем глобальным

        KompasAPI5 = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0) # API5 КОМПАСа
        iKompasObject = Dispatch("Kompas.Application.5", None, KompasAPI5.KompasObject.CLSID) # интерфейс API КОМПАС

        KompasAPI7 = gencache.EnsureModule('{69AC2981-37C0-4379-84FD-5DD2F3C0A520}', 0, 1, 0) # API7 КОМПАСа
        iApplication = Dispatch('Kompas.Application.7') # интерфейс приложения КОМПАС-3D.

        if iApplication.Visible == False: # если компас невидимый
            iApplication.Visible = True # сделать КОМПАС-3D видемым

    except: # если не получилось подключиться к КОМПАСу

        print("КОМПАС-3D не найден!\nУстановите или переустановите КОМПАС-3D!") # сообщение, поверх всех окон с автоматическим закрытием
        exit() # выходим из програмы

def ExcelAPI(path): # подключение API Excel (полный путь к файлу Excel)

    from win32com.client import Dispatch, gencache # библиотека API Windows

    try: # попытаться подключиться к Excel

        Excel = Dispatch("Excel.Application") # подключение к Excel
        wb = Excel.Workbooks.Open(path) # открытие файла Excel
        sheet = wb.ActiveSheet # получить текущий активный лист

        return sheet # возвращаем значение

    except: # если не получилось подключиться к Excel

        print("Файл Excel не найден!") # сообщение, поверх всех окон с автоматическим закрытием
        exit() # выходим из програмы

def Text_selection(): # обработка выделенного текта

    def Read_text(iSelectedObject): # прочитаем выделеный текст

        if iSelectedObject: # если выделено
            try: # попытаться определить тип выделенного объекта
                if iSelectedObject.DrawingObjectType == 4: # если это текст

                    iReference = iSelectedObject.Reference # указатель объекта

                    iNumber = iDocument2D.ksGetViewNumber(iReference) # номер вида по выделеному объекту
                    iView = iViews.ViewByNumber(iNumber) # вид, заданный по номеру

                    iDrawingContainer = KompasAPI7.IDrawingContainer(iView) # интерфейс контейнера объектов вида графического докумен
                    iDrawingTexts = iDrawingContainer.DrawingTexts # указатель на интерфейс коллекции текстов на чертеже

                    iDrawingText = iDrawingTexts.DrawingText(iReference) # интерфейс текста на чертеже
                    iText = KompasAPI7.IText(iDrawingText) # интерфейс текста
                    text = iText.Str # прочитанный текст
                    print(iText.Str)
                    Writing_in_Excel(text) # запись в Excel

            except: # если ошибка определения
                pass # пропускаем

        else: # не выделено
            print("Выделите текст!")

    if iApplication.ActiveDocument: # проверяем открыт ли файл в КОМПАСе

        iDocument2D = iKompasObject.ActiveDocument2D() # указатель на интерфейс текущего графического документа

        iKompasDocument = iApplication.ActiveDocument # получить текущий активный документ

        iKompasDocument2D = KompasAPI7.IKompasDocument2D(iKompasDocument) # базовый класс графических документов КОМПАС

        iViewsAndLayersManager = iKompasDocument2D.ViewsAndLayersManager # менеджер видов и слоев документа
        iViews = iViewsAndLayersManager.Views # коллекция видов

        iKompasDocument2D1 = KompasAPI7.IKompasDocument2D1(iKompasDocument) # дополнительный интерфейс IKompasDocument2D

        iSelectionManager = iKompasDocument2D1.SelectionManager # менеджер выделенных объектов
        iSelectedObjects = iSelectionManager.SelectedObjects # массив выделенных объектов в виде SAFEARRAY | VT_DISPATCH

        if isinstance(iSelectedObjects, tuple): # если выбрано несколько объектов (кортеж объектов)
            for iSelectedObject in iSelectedObjects: # перебор всех выделеных объектов
                Read_text(iSelectedObject) # прочитаем выделеный текст

        else:  # если выбран один объект
            iSelectedObject = iSelectedObjects # если один объект
            Read_text(iSelectedObject) # прочитаем выделеный текст

    else:
        print("Откройте чертёж!")

def Writing_in_Excel(text): # запись в Excel

    sheet.Cells(1,1).value = text # запись значения в ячейку

#-------------------------------------------------------------------------------

KompasAPI() # подключение API компаса

sheet = ExcelAPI(r"C:\Users\Каширских Дмитрий\Desktop\Дмитрий\ГОСТ\Прочее\Макросы\KText2Excel\Тест.xlsx") # подключение API Excel (полный путь к файлу Excel)

##val = sheet.Cells(1,2).value # значение ячейки Excel
##print(val)

while True: # цикл определения нажатия кнопки

    a = win32api.GetKeyState(0x01) # 0x02 правая / 0x01 левая
    b = win32api.GetKeyState(0x1B) # ESC key

    if a < 0: # если кнопка нажата
##        print("Кнопка нажата")
        Text_selection() # обработка выделенного текта

    if b < 0:
        print("ESC key")
        break # останавливаенм цикл

    time.sleep(0.08) # сон в секундах