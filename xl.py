import xlrd
import xlwt
from io import BytesIO


def _get_cell(sheet, row_idx, col_idx):
    """
    Получить ячейку.

    :param sheet: страница (xlwt.Worksheet)
    :param col: столбец (координаты ячейки)
    :param row: строка (координаты ячейки)
    """
    row = sheet._Worksheet__rows.get(row_idx)
    if not row:
        return None
    cell = row._Row__cells.get(col_idx)
    return cell


def set_value_cell(sheet, row, col, value, pattern_row=None, pattern_col=None):
    """
    Изменение значения ячейки без изменения форматирования.
    Если pattern_col и pattern_row не заданы, стили текущей ячейки сохранятся теже,
    которые были заданы до записи нового значения в ячейку, если заданы, стили будут
    скопированны из указанной ячейки.

    :param sheet: страница (xlwt.Worksheet)
    :param col: столбец (координаты текущей ячейки)
    :param row: строка (координаты текущей ячейки)
    :param value: значение для записи в текущую ячейку
    :param pattern_col: столбец (координаты ячейки, от которой будут унаследованы стили текущей ячейки)
    :param pattern_row: строка (координаты ячейки, от которой будут унаследованы стили текущей ячейки)
    """
    col_width = row_height = None

    if pattern_col is not None and pattern_row is not None:
        prev_cell = _get_cell(sheet, pattern_row, pattern_col)
        col_width = sheet.col(pattern_col).width
        row_height = sheet.row(pattern_row).height
    else:
        prev_cell = _get_cell(sheet, row, col)

    # запись значения в ячейку
    sheet.write(row, col, value)
    # установка стилей
    if prev_cell:
        new_cell = _get_cell(sheet, row, col)
        if new_cell:
            new_cell.xf_idx = prev_cell.xf_idx
    # установка ширины столбца
    if col_width:
        sheet.col(col).width = col_width
    # установка высоты строки
    if row_height:
        sheet.row(row).height = row_height


def _delete_marge_cell(sheet_wt, value):
    """Удалить текущее объединение ячеек."""
    try:
        sheet_wt._Worksheet__merged_ranges.remove(value)
    except ValueError:
        pass


def _get_marged_cell(sheet_wt, row, col):
    """
    Возвращает координаты объединения текущей ячейки с другими ячейками,
    если она входит в объединение.
    """
    for merge in sheet_wt.merged_ranges:
        if row in range(merge[0], merge[1]+1) and col in range(merge[2], merge[3]+1):
            return merge


def _is_skip_cells(sheet_wt, row, col):
    """
    Проверка нужно ли вставить в текущую ячейку значение или ее следует пропустить,
    если это не верхняя левая ячейка входящая в объединение.
    """
    marge = _get_marged_cell(sheet_wt, row, col)
    if marge:
        if (row == marge[0] and col != marge[2]) or (row != marge[0]):
            return True
    return False


def insert_rows(workbook_wt: xlwt.Workbook, sheet_wt, idx, amount=1):
    """
    Вставка строки с возможностью указать кол-во вставляемых строк.
    При вставке строк объединения ячеек корректно перемещаются.
    """
    sheet_name = sheet_wt.get_name()

    buf_sheet = BytesIO()
    workbook_wt.save(buf_sheet)

    workbook_rd = xlrd.open_workbook(file_contents=buf_sheet.getvalue(), formatting_info=True)
    sheet_rd = workbook_rd.sheet_by_name(sheet_name)

    # кол-во строк и столбцов
    nrows = sheet_wt.last_used_row
    ncols = sheet_wt.last_used_col

    # список исходных объединенний ячеек
    merged_ranges = sheet_wt.merged_ranges.copy()

    # проверяем индекс и уменьшаем на единицу, чтобы он соответсвовал номеру строки
    if idx <= 0:
        idx = 1
    idx -= 1

    for row in range(nrows, idx-1, -1):
        for col in range(0, ncols+1):
            if not _is_skip_cells(sheet_wt, row, col):
                value = sheet_rd.cell_value(row, col)
                marged_cell = _get_marged_cell(sheet_wt, row, col)
                # если ячейка была объединена, делаем новое объединенние ячеек
                # с учетом смещения строк
                if marged_cell:
                    row1, row2, col1, col2 = marged_cell
                    sheet_wt.write_merge(row1+amount, row2+amount, col1, col2, '')
                # записываем данные в ячейку
                set_value_cell(sheet_wt, row+amount, col, value, row, col)

        # очистка и установка стандартной высоты смещенной строки
        row = sheet_wt.rows[row]
        row.write_blanks(0, ncols)
        row.height = 256

    # удаление старого объединения ячеек
    for merge in merged_ranges:
        # если строки с объединеними были смещены
        if idx <= merge[0]:
            _delete_marge_cell(sheet_wt, merge)
        # если вставка строки попала в конец или в серидину объединеной ячейки,
        # увеличиваем это объединение
        elif merge[0] < idx <= merge[1]:
            _delete_marge_cell(sheet_wt, merge)
            row1, row2, col1, col2 = merge
            sheet_wt.merge(row1, row2 + amount, col1, col2)

    buf_sheet.close()
