import pandas as pd
import win32com.client as win32  # it seems to work only on Windows
from pathlib import Path


def main():
    #  csv -->> xlsx
    pd.read_csv("titanic.csv").to_excel("pt_example.xlsx", index=False)

    def clear_pt_sheet(ws):
        "Удалить сводную с листа"
        for pt in ws.PivotTables():
            pt.TableRange2.clear()

    xl = win32.Dispatch("Excel.Application")
    xl.Visible = True

    # открываем книгу и определяем лист с исходными данными
    wb = xl.Workbooks.Open(Path.cwd() / "pt_example.xlsx")
    ws_data = wb.Worksheets("Sheet1")

    # первый лист для сводной
    ws_pivot1 = wb.Worksheets.Add(After=ws_data)
    ws_pivot1.Name = "PivotTable_1"

    # второй лист для сводной
    ws_pivot2 = wb.Worksheets.Add(After=ws_pivot1)
    ws_pivot2.Name = "PivotTable_2"

    # очистка листа перед построением сводной
    clear_pt_sheet(ws_pivot1)
    clear_pt_sheet(ws_pivot2)

    # добавление данных в кэш
    pt_cache = wb.PivotCaches().Create(1, ws_data.Range("A1").CurrentRegion)

    # добавление сводной таблицы
    pt1 = pt_cache.CreatePivotTable(ws_pivot1.Range("A1"), "TitanicPivotTableBySex")
    pt2 = pt_cache.CreatePivotTable(
        ws_pivot2.Range("A1"), "TitanicPivotTableByManyColumns"
    )

    # итоги
    pt1.ColumnGrand = True
    pt1.RowGrand = False

    pt2.ColumnGrand = True
    pt2.RowGrand = False

    # расположения отчета
    pt1.RowAxisLayout(1)
    pt2.RowAxisLayout(1)

    # стиль таблицы
    pt1.TableStyle2 = "PivotStyleMedium9"
    pt2.TableStyle2 = "PivotStyleMedium10"

    def first_pt(pt):
        "Создание первой таблицы"
        # строки
        field_rows = {}
        field_rows["Пол"] = pt.PivotFields("Sex")

        # значения
        field_values = {}
        field_values["Кол. пассажиров"] = pt.PivotFields("PassengerId")
        field_values["Кол. выживших"] = pt.PivotFields("Survived")
        field_values["Max возраст"] = pt.PivotFields("Age")
        field_values["Min возраст"] = pt.PivotFields("Age")
        field_values["Mean возраст"] = pt.PivotFields("Age")

        field_rows["Пол"].Orientation = 1
        field_rows["Пол"].Position = 1

        # операции агрегации
        field_values["Кол. пассажиров"].Orientation = 4
        field_values["Кол. пассажиров"].Function = -4112
        field_values["Кол. пассажиров"].NumberFormat = "# ##0"
        field_values["Кол. пассажиров"].Caption = "Количество пассажиров"

        field_values["Кол. выживших"].Orientation = 4
        field_values["Кол. выживших"].Function = -4157
        field_values["Кол. выживших"].NumberFormat = "# ##0"
        field_values["Кол. выживших"].Caption = "Количество выживших"

        field_values["Max возраст"].Orientation = 4
        field_values["Max возраст"].Function = -4136
        field_values["Max возраст"].NumberFormat = "# ##0"
        field_values["Max возраст"].Caption = "Максимальный возраст"

        field_values["Min возраст"].Orientation = 4
        field_values["Min возраст"].Function = -4139
        field_values["Min возраст"].NumberFormat = "# ##0,00"
        field_values["Min возраст"].Caption = "Минимальный возраст"

        field_values["Mean возраст"].Orientation = 4
        field_values["Mean возраст"].Function = -4106
        field_values["Mean возраст"].NumberFormat = "# ##0"
        field_values["Mean возраст"].Caption = "Средний возраст"

        # полгон ширины ячеек по содержимому
        ws_pivot1.Columns.AutoFit()

    def second_pt(pt):
        "Создание второй таблицы"
        # строки
        field_rows = {}
        field_rows["Порт посадки"] = pt.PivotFields("Embarked")
        field_rows["Класс билета"] = pt.PivotFields("Pclass")
        field_rows["Выживаемость"] = pt.PivotFields("Survived")

        # значения
        field_values = {}
        field_values["Супруги"] = pt.PivotFields("SibSP")
        field_values["Дети"] = pt.PivotFields("Parch")
        field_values["Стоимость билета"] = pt.PivotFields("Fare")
        field_values["Mean возраст"] = pt.PivotFields("Age")

        field_rows["Порт посадки"].Orientation = 1
        field_rows["Порт посадки"].Position = 1

        field_rows["Класс билета"].Orientation = 1
        field_rows["Класс билета"].Position = 2

        field_rows["Выживаемость"].Orientation = 1
        field_rows["Выживаемость"].Position = 3

        # операции агрегации
        field_values["Супруги"].Orientation = 4
        field_values["Супруги"].Function = -4157
        field_values["Супруги"].NumberFormat = "# ##0"
        field_values["Супруги"].Caption = "Количество супругов"

        field_values["Дети"].Orientation = 4
        field_values["Дети"].Function = -4157
        field_values["Дети"].NumberFormat = "# ##0"
        field_values["Дети"].Caption = "Количество детей"

        field_values["Стоимость билета"].Orientation = 4
        field_values["Стоимость билета"].Function = -4157
        field_values["Стоимость билета"].NumberFormat = "# ##0"
        field_values["Стоимость билета"].Caption = "Стоимость билетов"

        # добавление условного форматирования
        col_range = field_values["Стоимость билета"].DataRange
        col_range.FormatConditions.AddColorScale(3)

        ws_pivot2.Columns.AutoFit()

    first_pt(pt1)
    second_pt(pt2)

    wb.Save()
    wb.Close()

    xl.Quit()


if __name__ == "__main__":
    main()
