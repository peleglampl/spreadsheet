import pytest
from sheet import Sheet
import pandas as pd

def test_spreadsheet_creation():
    rows, columns = 5, 3
    sheet = Sheet(rows, columns)
    assert sheet.rows == rows
    assert sheet.columns == columns
    assert len(sheet.df.columns) == columns
    assert len(sheet.df) == rows

def test_initial_cell_values():
    rows, columns = 4, 4
    sheet = Sheet(rows, columns)
    for col in sheet.df.columns:
        for row in sheet.df.index:
            assert sheet.get_cell_value(row, col) == '', f"Cell value at {col}{row} is not an empty string."

def test_enter_new_value():
    sheet = Sheet(3, 3)
    cell_label = 'A1'
    value = 'Hello'
    sheet.enter_new_value(cell_label, value)
    assert sheet.get_cell_value_by_label(cell_label) == value, "Failed to enter new value into the cell."

def test_circular_dependency_detection():
    sheet = Sheet(2, 2)
    # Creating a circular dependency: A1 -> B1 -> A1
    sheet.enter_new_value('A1', '=B1')
    sheet.enter_new_value('B1', '=A1')
    assert sheet.check_if_circular_dependency_from_cell('A1') == True, "Failed to detect circular dependency."


def test_change_reference_in_cell():
    sheet = Sheet(2, 2)
    sheet.enter_new_value('A1', '5')
    sheet.enter_new_value('B1', '=A1')
    assert sheet.evaluate_cell('B1') == 5.0, "Cell B1 did not correctly reference the value of A1."

    # Change the value of A1 and check if B1's value is updated correctly
    sheet.enter_new_value('A1', '10')
    sheet.change_reference_in_cell('B1')  # Assuming this should re-evaluate B1 based on A1's new value
    assert sheet.evaluate_cell('B1') == 10.0, "Cell B1 did not update based on the new value of A1."


def test_list_of_cells_labels():
    sheet = Sheet(4, 3)
    assert sheet.list_of_cells_labels() == ['A1', 'A2', 'A3', 'A4', 'B1', 'B2', 'B3', 'B4', 'C1', 'C2', 'C3', 'C4']

def test_get_cell_row_by_label():
    sheet = Sheet(4, 4)
    assert sheet.get_cell_row_by_label('A1') == 1
    assert sheet.get_cell_row_by_label('B2') == 2
    assert sheet.get_cell_row_by_label('C3') == 3
    assert sheet.get_cell_row_by_label('D4') == 4

def test_empty_sheet():
    sheet = Sheet(0, 0)
    assert sheet.df.empty == True

def test_get_cell_col_by_label():
    sheet = Sheet(0, 0)
    assert sheet.get_cell_col_by_label('A1') == 'A'
    assert sheet.get_cell_col_by_label('B2') == 'B'
    assert sheet.get_cell_col_by_label('C3') == 'C'
    assert sheet.get_cell_col_by_label('D4') == 'D'

def test_get_cell_value_by_label():
    dict_of_cells= {
        'A': ['=2/1', "2", '=A2*B1', "=2/0"],
        'B': ["3.14", '=SUM(1, 2, 3)', "=MIN(90)", "=SUM(3, 33, 7)"],
        'C': ["=SUM(A2, B1)", "=MIN(3, A1, A2, 0)", 'A1+A4', '=abf']
    }
    sheet = Sheet(4, 3)
    sheet.df = pd.DataFrame(dict_of_cells, index=range(1, len(max(dict_of_cells.values(), key=len)) + 1))
    assert sheet.get_cell_value_by_label('A1') == '=2/1'
    assert sheet.get_cell_value_by_label('B2') == '=SUM(1, 2, 3)'
    assert sheet.get_cell_value_by_label('C3') == 'A1+A4'
    assert sheet.get_cell_value_by_label('C4') == '=abf'

def test_get_original_formula():
    dict_of_cells= {
        'A': ['=2/1', "2", '=A2*B1', "=2/0"],
        'B': ["3.14", '=SUM(1, 2, 3)', "=MIN(90)", "=SUM(3, 33, 7)"],
        'C': ["=SUM(A2, B1)", "=MIN(3, A1, A2, 0)", 'A1+A4', '=abf']
    }
    sheet = Sheet(4, 3)
    sheet.original_df = pd.DataFrame(dict_of_cells, index=range(1, len(max(dict_of_cells.values(), key=len)) + 1))

    assert sheet.get_original_formula(1, 'A') == '=2/1'
    assert sheet.get_original_formula(2, 'B') == '=SUM(1, 2, 3)'
    assert sheet.get_original_formula(3, 'C') == 'A1+A4'
    assert sheet.get_original_formula(4, 'C') == '=abf'

def test_convert_label_to_row_col():
    sheet = Sheet(4, 3)
    assert sheet.convert_label_to_row_col('A1') == (1, 'A')
    assert sheet.convert_label_to_row_col('B2') == (2, 'B')
    assert sheet.convert_label_to_row_col('C3') == (3, 'C')
    assert sheet.convert_label_to_row_col('D4') == (4, 'D')
    assert sheet.convert_label_to_row_col('D5') == (5, 'D')

def test_is_range_formula():
    sheet = Sheet(4, 3)
    arguments_list = ['A1, B1', 'A2', 'B2', 'A3', 'B3', 'A4', 'B4']
    assert sheet.is_range_formula('=A1:B2') == False
    assert sheet.is_range_formula('A1') == False
    assert sheet.is_range_formula('A1:B') == False
    assert sheet.is_range_formula('A:B') == False
    assert sheet.is_range_formula('=SUM(A1:B2)') == True
    assert sheet.is_range_formula('=MIN(A1:B2)') == True
    assert sheet.is_range_formula('=MAX(B1:A2)') == True
    assert sheet.is_range_formula('=AVERAGE(A1:N100)') == False

def test_evaluate_range_function():
    dict_of_cells= {
        'A': ['2', "2", '5.666', "#DIV/0!"],
        'B': ["3.14", '5.789', "=MIN(A1:A2)", "=SUM(A1:B2)"],
        'C': ["=AVERAGE(A1:A4)", "=MAX(A1:B2)", 'A1+A4', '=abf']
    }
    sheet = Sheet(4, 3)
    sheet.df = pd.DataFrame(dict_of_cells, index=range(1, len(max(dict_of_cells.values(), key=len)) + 1))
    assert sheet.evaluate_range_function('=MIN(A1:A2)', 'B3') == 2
    assert sheet.evaluate_range_function('=SUM(A1:B2)', 'B4') == 12.929
    assert sheet.evaluate_range_function('=AVERAGE(A1:A4)', 'C1') == '#DIV/0!'
    assert sheet.evaluate_range_function('=MAX(A1:B2)', 'C2') == 5.789


def test_evaluate_formula():
    dict_of_cells= {
        'A': ['2', "2", '5.666', "#DIV/0!"],
        'B': ["3.14", '5.789', "=MIN(A1:A2)", "=SUM(A1:B2)"],
        'C': ["=AVERAGE(A1:A4)", "=MAX(A1:B2)", '=A1+A4', 'abf']
    }
    sheet = Sheet(4, 3)
    sheet.df = pd.DataFrame(dict_of_cells, index=range(1, len(max(dict_of_cells.values(), key=len)) + 1))
    assert sheet.evaluate_formula('2', 'A1') == '2'
    assert sheet.evaluate_formula('5.789','B2') == '5.789'
    assert sheet.evaluate_formula('=abf','C4') == ''
    assert sheet.evaluate_formula('=AVERAGE(A1:A4)', 'C1') == '#DIV/0!'
    assert sheet.evaluate_formula('=MIN(A1:A2)', 'B3') == 2.0

