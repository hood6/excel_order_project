import xlrd as r
import os


def find_out_of_stock():
    item_name_ind = 2
    store_ware_ind = 7
    in_stock_ind = 7
    min_stock_ind = 8
    max_stock_ind = 9
    main_ware_ind = 12
    expected_ind = 13
    ret = {}
    loc = "rudd_markup.xlsx"
    wb = r.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    for i in range(1, sheet.nrows):
        if int(sheet.cell_value(i, in_stock_ind)) < int(sheet.cell_value(i, min_stock_ind))\
                and sheet.cell_value(i, expected_ind) == ""\
                and int(sheet.cell_value(i, main_ware_ind)) > 0:
            ret[i] = (str(sheet.cell_value(i, item_name_ind)))
    return ret
