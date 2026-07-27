from copy import copy
from openpyxl import load_workbook
import pandas as pd


def merge_data(workbook, source_file):
    header_row = int(input('С какой строки начинаются заголовки файла? Введите число, например, 2 >>> ')) - 1

    df = pd.read_excel(source_file, header=header_row)
    print(df.columns.tolist())

    # открываем исходный файл
    source_wb = load_workbook(source_file)
    source_ws = source_wb.active

    if "Export" not in workbook.sheetnames:
        ws = workbook.create_sheet("Export")
    else:
        ws = workbook["Export"]


    # берём вкладку Export из шаблона
    if "Export" not in workbook.sheetnames:
        ws = workbook.create_sheet("Export")
    else:
        ws = workbook["Export"]


    # очищаем Export если там что-то было
    for merged_range in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged_range))


    for row in ws.iter_rows():
        for cell in row:
            cell.value = None


    # копируем значения и стили
    for row in source_ws.iter_rows():

        for cell in row:

            target = ws[cell.coordinate]

            target.value = cell.value

            if cell.has_style:
                target._style = copy(cell._style)

            if cell.number_format:
                target.number_format = cell.number_format

            if cell.alignment:
                target.alignment = copy(cell.alignment)

            if cell.border:
                target.border = copy(cell.border)

            if cell.fill:
                target.fill = copy(cell.fill)

            if cell.font:
                target.font = copy(cell.font)


    # ширина колонок
    for col, dim in source_ws.column_dimensions.items():
        ws.column_dimensions[col].width = dim.width


    # высота строк
    for row, dim in source_ws.row_dimensions.items():
        ws.row_dimensions[row].height = dim.height


    print(
        f"Export заполнен: {source_ws.max_row} строк, "
        f"{source_ws.max_column} колонок"
    )


    return df