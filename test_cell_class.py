import pytest
from cell import Cell


def test_check_if_function_1():
    cell = Cell("=SUM(1, 2+3, 3) / AVERAGE(4, 5)")
    assert cell.check_if_function(cell.get_cell_value()) == True


def test_check_if_function_2():
    cell = Cell("=SUN(1, 2)")
    assert cell.check_if_function(cell.get_cell_value()) == True


def test_check_if_function_3():
    cell = Cell("=1+2")
    assert cell.check_if_function(cell.get_cell_value()) == False


def test_check_if_function_4():
    cell = Cell("AB")
    assert cell.check_if_function(cell.get_cell_value()) == False


def test_check_if_function_5():
    cell = Cell("=AB()")
    assert cell.check_if_function(cell.get_cell_value()) == True


def test_calculate_simple_formula_1():
    cell = Cell("1+2-5")
    assert cell.calculate_simple_formula() == -2


def test_calculate_simple_formula2():
    cell = Cell("1+2*3")
    assert cell.calculate_simple_formula() == 7


def test_calculate_simple_formula_3():
    cell = Cell("1+2/2")
    assert cell.calculate_simple_formula() == 2


def test_calculate_simple_formula_zero_division():
    cell = Cell("1/0")
    assert cell.calculate_simple_formula() == "#DIV/0!"


def test_calculate_simple_formula_type_error():
    cell = Cell("1+'two'")
    assert cell.calculate_simple_formula() == "#TYPE_ERROR"


def test_calculate_simple_formula_test4():
    cell = Cell("1 + 2+ 3+5+6+8+9/22")
    assert cell.calculate_simple_formula() == 25.40909090909091


def test_valid_function_1():
    cell = Cell("=SUM(1, 2+3, 3) / AVERAGE(4, 5)")
    assert cell.evaluate_functions() == "#VALUE!"


def test_valid_function_2():
    cell = Cell("=SUN(1, 2)")
    assert cell.evaluate_functions() == "#INVALID_FUNCTION"


def test_valid_function_3():
    cell = Cell("=MIN()")
    assert cell.evaluate_functions() == "#N/A"


def test_sqrt_function():
    cell = Cell("=SQRT(4)")
    assert cell.evaluate_functions() == 2


def test_sqrt_function_test2():
    cell = Cell("=SQRT(9, 3)")
    assert cell.evaluate_functions() == "#N/A"


def test_sqrt_function_test3():
    cell = Cell("=SQRT()")
    assert cell.evaluate_functions() == "#N/A"


def test_mod_function():
    cell = Cell("=MOD(4, 2)")
    assert cell.evaluate_functions() == 0

def test_mod_function_test2():
    cell = Cell("=MOD(4, 3)")
    assert cell.evaluate_functions() == 1

def test_mod_function_test3():
    cell = Cell("=MOD(4)")
    assert cell.evaluate_functions() == "#N/A"

def test_mod_function_test4():
    cell = Cell("=MOD(4, 3, 2)")
    assert cell.evaluate_functions() == "#N/A"

def test_mod_function_test5():
    cell = Cell("=MOD()")
    assert cell.evaluate_functions() == "#N/A"

def test_mod_function_test6():
    cell = Cell("=MOD(4, 'two')")
    assert cell.evaluate_functions() == "#ERROR!"

def test_mod_function_test7():
    cell = Cell("=MOD(A1:B2)")
    assert cell.evaluate_functions() == "#VALUE!"

def test_valid_function_4():
    cell = Cell("=AB()")
    assert cell.evaluate_functions() == "#INVALID_FUNCTION"


def test_valid_function_5():
    cell = Cell("=SUM(1, 2, 3)")
    assert cell.evaluate_functions() == 6


def test_valid_function_6():
    cell = Cell("=AVERAGE(1, 2, 3, 4)")
    assert cell.evaluate_functions() == 2.5


def test1():
    cell = Cell("=SUM(1, 2+3, 3)")
    assert cell.main_evaluate_cell() == 9


def test_simple_calculation():
    cell = Cell("=2 + 3*4")
    assert cell.main_evaluate_cell() == 14


def test2():
    cell = Cell("=AVERAGE(4, 5)")
    assert cell.main_evaluate_cell() == 4.5


def test3():
    cell = Cell("=MAX(1, 2, 3)")
    assert cell.main_evaluate_cell() == 3


def test4():
    cell = Cell("=MIN(1, 2, 3, 4)")
    assert cell.main_evaluate_cell() == 1


def no_function():
    cell = Cell("1+2")
    assert cell.main_evaluate_cell() == "1+2"


def test5():
    cell = Cell("=abc")
    assert cell.main_evaluate_cell() == "#NAME?"


def test6():
    cell = Cell("abc")
    assert cell.main_evaluate_cell() == "abc"


def test7():
    cell = Cell("=SUN(1, 2, 3)")
    assert cell.main_evaluate_cell() == "#INVALID_FUNCTION"
