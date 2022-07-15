import sys
import openpyxl

import SystemConfig
import userConfig

def read_sheet(sheet_name, last_column_name):
    wb_obj = openpyxl.load_workbook(userConfig.excelFileName)
    SystemConfig.sheetObj = wb_obj[sheet_name]
    SystemConfig.maxRows = SystemConfig.sheetObj.max_row + 1
    set_max_column(last_column_name)

def get_cell_location_of_string(expected_string):
    for current_row in range(1, SystemConfig.maxRows + 1):
        for current_col in range(1, SystemConfig.maxCol + 1):
            cell_value = get_cell_value(current_row, current_col)
            cell_value = str(cell_value).strip()
            if cell_value == str(expected_string):
                return [current_row, current_col]

    print "Expected String : {0} was not found in excel.".format(expected_string)
    print "Terminating Execution since Excel Template was tampered."
    sys.exit(-1)

def get_column_number_of_string(expected_string):
    return get_cell_location_of_string(expected_string)[1]

def get_row_number_of_string(expected_string):
    return get_cell_location_of_string(expected_string)[0]

def set_max_column(column_name):
    SystemConfig.maxCol = 100
    SystemConfig.maxCol = get_column_number_of_string(column_name)

def get_max_column():
    return SystemConfig.maxCol

def get_cell_value(row, col):
    return SystemConfig.sheetObj.cell(row, col).value
