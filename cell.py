from sympy import sympify, SympifyError, oo, zoo
import re


class Cell:
    """This class represents a cell in a spreadsheet. The cell can contain a value or a formula. If the cell contains
    a formula, the formula is evaluated and the result is returned. If the formula contains a reference to another cell,
    the value of the referred cell is used in the evaluation. The class also supports custom functions like SUM,
     AVERAGE, MIN, and MAX."""
    def __init__(self, value: str) -> None:
        """Create a new cell with the given value."""
        self.value = value

    def get_cell_value(self) -> str:
        """Return the value of the cell."""
        return self.value

    def set_cell_value(self, new_value: str) -> None:
        """Set the value of the cell."""
        self.value = new_value

    def main_evaluate_cell(self) -> str:
        """Evaluate the cell's value. and replace the cell with the result"""
        if self.value is None or self.value == "":  # If the cell is empty,
            return ""
        if str(self.value)[0] != "=":  # If the value is not a formula (starts at "="), return it as is
            return self.value
        # If the value starts with "=", it is a formula
        if self.check_if_function(self.value):
            self.set_cell_value(self.evaluate_functions())
            return self.value
        else:
            self.value = self.value.strip(self.value[0])
            self.set_cell_value(self.calculate_simple_formula())
            return self.value

    def check_if_function(self, value: str):
        """Check if the input value is a function."""
        # Basic pattern that matches strings starting with "=" followed by a function name
        # (e.g., "SUM", "AVERAGE") and parentheses, or arithmetic operations
        pattern = r"^=\s*([A-Z]+)\(.*\)"
        # Check if the input matches the pattern
        return bool(re.match(pattern, value))


    def calculate_simple_formula(self) -> str:
        """Evaluate the cell's value if it contains a simple formula. (*, +, -, /)"""
        formula = self.get_cell_value()
        try:
            # Use SymPy to evaluate the formula
            result = sympify(formula).evalf()
            # Check if result is infinity, which indicates division # by zero
            if result == oo or result == -oo or result == zoo:
                return "#DIV/0!"
            elif result.is_Symbol:  # If the result is a symbol, return #NAME?
                return "#NAME?"
        # Handle exceptions:
        except ZeroDivisionError:  # Division by zero
            return "#DIV/0!"
        except SympifyError:  # Invalid formula
            return "#ERROR!"
        except TypeError:  # Type error
            return "#TYPE_ERROR"
        except ValueError:  # Value error
            return "#VALUE!"
        except Exception:  # Other exceptions
            return "#ERROR!"
        # If the result is a number, return it (This if is for not returning a numbers with decimal points)
        if result.is_Number:
            if result.is_Integer:
                return int(result)
            else:
                return float(result)
        return result

    def evaluate_functions(self):
        """Evaluate the cell's value if it contains a function."""
        function_name = self.value.split('(')[0][1:]  # Extract the function name
        arguments = self.value.split('(')[1].rstrip(')')  # Extract the arguments
        # If the function is not one of the supported functions, return #INVALID_FUNCTION
        if function_name not in ['MIN', 'MAX', 'SUM', 'AVERAGE', 'SQRT', 'MOD']:
            return "#INVALID_FUNCTION"
        else:  # If the function is one of the supported functions
            if arguments == "":  # If the function has no arguments, return #N/A (no arguments)
                return "#N/A"
            # If the argument is one number, and the function is SQRT, return the square root of the number:
            if arguments.isdigit() and function_name == 'SQRT':
                return float(arguments) ** 0.5
             # If the argument is one number, and the function is MOD, return #N/A:
            elif arguments.isdigit() and function_name == 'MOD':
                return "#N/A"
            # If the argument is one number, and the function is not SQRT, return the number:
            elif arguments.isdigit() and function_name != 'SQRT':
                return arguments

            if ',' in self.value:  # If the function has multiple arguments
                argument_list = [arg.strip() for arg in arguments.split(',')]  # Split the arguments
                argument_list = self.calculate_in_list(argument_list)  # Calculate the value of each argument
                # Check if any of the arguments is an error, using the check_error method:
                if self.check_error(argument_list):
                    return self.check_error(argument_list)
                argument_list = self.type_cast(argument_list)
                if function_name == "SUM":
                    return sum(argument_list)
                elif function_name == "AVERAGE":
                    try:
                        return sum(argument_list)/len(argument_list)
                    except ZeroDivisionError:
                        return "#DIV/0!"
                elif function_name == "MIN":
                    return min(argument_list)
                elif function_name == "MAX":
                    return max(argument_list)
                elif function_name == "SQRT":
                    return "#N/A"
                elif function_name == "MOD":
                    if len(argument_list) != 2:
                        return "#N/A"
                    return argument_list[0] % argument_list[1]
            else:
                return '#VALUE!'

    def check_error(self, argument_list):
        if "#VALUE!" in argument_list:
            return "#VALUE!"
        if "#DIV/0!" in argument_list:
            return "#DIV/0!"
        if "#NAME?" in argument_list:
            return "#NAME?"
        if "#ERROR!" in argument_list:
            return "#ERROR!"
        if "#TYPE_ERROR" in argument_list:
            return "#TYPE_ERROR"
        if "#N/A" in argument_list:
            return "#N/A"
        if "#INVALID_FUNCTION" in argument_list:
            return "#INVALID_FUNCTION"
        if "#CIRCULAR_DEPENDENCY" in argument_list:
            return "#CIRCULAR_DEPENDENCY"

    def type_cast(self, list_of_strings):
        """Convert a list of strings to a list of numbers, if possible. If a string cannot be converted to a number, it
        is left as is."""
        list_of_numbers = []
        for string in list_of_strings:
            try:
                list_of_numbers.append(float(string))
            except ValueError or TypeError:
                return list_of_strings
            except Exception:
                return list_of_strings
        return list_of_numbers

    def calculate_in_list(self, list_elements):
        """Calculate the value of each item in the list and return a new list with the calculated values."""
        new_list = []
        for item in list_elements:
            try:
                new_arg = sympify(item).evalf()
                if new_arg == zoo:
                    new_list.append("#DIV/0!")
                else:
                    new_list.append(new_arg)
            except ZeroDivisionError:
                new_list.append("#DIV/0!")
            except TypeError:
                new_list.append('#NAME?')
            except ValueError:
                new_list.append('#VALUE!')
            except Exception:
                new_list.append("#ERROR!")
        return new_list
