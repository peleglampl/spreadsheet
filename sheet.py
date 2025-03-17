import pandas as pd
from cell import Cell
import re
import json

class Sheet:
    """This class represents a spreadsheet with cells that can contain values or formulas. The class supports basic"""
    def __init__(self, rows, columns):
        """Create a new spreadsheet with the given number of rows and columns."""
        self.rows = rows
        self.columns = columns
        # Initialize a dictionary of cells with empty strings
        self.dict_of_cells = {chr(ord('A') + j): {i + 1: '' for i in range(rows)} for j in range(columns)}
        if self.dict_of_cells != {}:
            self.df = pd.DataFrame(self.dict_of_cells, index=range(1, len(max(self.dict_of_cells.values(), key=len)) + 1))
        else:
            self.df = pd.DataFrame(index=range(1, rows + 1), columns=range(1, columns + 1))
        # Keep a copy of the original DataFrame for reference
        self.original_df = self.df.copy()
        # Initialize a dictionary to store cell dependencies
        self.dependencies = {}
        
    def get_cell_value(self, row, col):
        """Return the value of the cell at the given row and column indices."""
        return self.df.at[row, col]

    def get_cell_value_by_label(self, cell_label):
        """Return the value of the cell with the given label."""
        row, col = self.convert_label_to_row_col(cell_label)
        return self.get_cell_value(row, col)

    def get_cell_row_by_label(self, cell_label):
        """Return the row index of the cell with the given label."""
        return int(re.search(r'\d+', cell_label).group())

    def get_cell_col_by_label(self, cell_label):
        """Return the column label of the cell with the given label."""
        return re.search(r'[A-Z]+', cell_label).group()

    def convert_label_to_row_col(self, cell_label):
        """Convert a cell label to row and column indices."""
        return self.get_cell_row_by_label(cell_label), self.get_cell_col_by_label(cell_label)

    def get_original_formula(self, row, col):
        """Return the original formula of the cell at the given row and column indices."""
        return self.original_df.at[row, col]

    def enter_new_value(self, cell_label, value):
        """Enter a new value into the cell."""
        self.original_df.at[self.convert_label_to_row_col(cell_label)] = value
        self.df.at[self.convert_label_to_row_col(cell_label)] = value
        # Update cell dependencies
        self.update_cell_dependencies(cell_label)

    def list_of_cells_labels(self):
        """Return a list of cell labels in the spreadsheet."""
        cell_labels = []
        columns = self.df.columns
        # Ensure row indices start from 1 for spreadsheet-like indexing
        rows = range(1, len(self.df) + 1)
        for column in columns:
            for row in rows:
                cell_labels.append(f"{column}{row}")
        return cell_labels

    def change_reference_in_cell(self, cell_label):
        """Change the reference in the cell to the actual value of the referenced cell."""
        row, col = self.convert_label_to_row_col(cell_label)
        value = self.original_df.at[row, col]
        cell = Cell(value)
        for cell_label_in_value in re.findall(r'[A-Z]+\d+', str(value)): # Find all cell references in the value
            if cell_label_in_value not in self.list_of_cells_labels():  # Check if the cell reference exists
                return Cell("#REF!")
            if cell_label_in_value == cell_label:  # Check for circular references
                return Cell("#REF!")
            else:
                # Replace the cell reference with the actual value
                cell.value = cell.value.replace(cell_label_in_value,
                                                str(self.get_cell_value_by_label(cell_label_in_value)))
        self.df.at[row, col] = cell.value # Update the cell value in the DataFrame
        return cell
    
    def parse_formula(self, cell_label):
        """Parse the formula of the cell to find cell dependencies."""
        formula = self.get_cell_value_by_label(cell_label)
        list_of_depend=re.findall(r'[A-Z][1-9]\d*', formula) # Find all cell references in the formula
        self.dependencies[cell_label] = list_of_depend
        return list_of_depend

    def check_if_circular_dependency_from_cell(self, start_cell):
        """Check if there is a circular dependency starting from the given cell."""
        visited = set()
        rec_stack = set()
        def is_cyclic_util(cell):
            if cell not in self.dependencies:
                return False  # No dependencies, so no cycle from this cell

            visited.add(cell)
            rec_stack.add(cell)

            for dependent in self.dependencies[cell]:
                if dependent not in visited:
                    if is_cyclic_util(dependent):
                        return True
                elif dependent in rec_stack:
                    return True

            rec_stack.remove(cell)
            return False

        # Start the cycle check from the specific cell
        return is_cyclic_util(start_cell)

    def is_range_formula(self, value):
        """Check if the cell's value contains a range formula. (e.g., SUM(A1:A10))"""
        cell = Cell(value)
        if str(value)[0] != "=":
            return False
        if cell.check_if_function(value):
            # Check if the function is a range function (e.g., SUM, MIN, MAX, AVERAGE)
            function_name = cell.value.split('(')[0][1:]
            arguments = cell.value.split('(')[1].rstrip(')')
            if function_name not in ['MIN', 'MAX', 'SUM', 'AVERAGE']:
                return False
            if ':' not in arguments: # Check if there is a range argument
                return False
            if arguments.count(':') != 1: # Check if there is only one range argument
                return False
            argument_list = arguments.split(':')
            # Check if the range arguments are valid cell labels
            if (argument_list[0] not in self.list_of_cells_labels() or argument_list[1]
                    not in self.list_of_cells_labels()):
                return False
            return True
        return False

    def evaluate_range_function(self, value, cell_label):
        """Evaluate the cell's value if it contains a range function. (e.g., SUM, MIN, MAX, AVERAGE)"""
        cell = Cell(value)
        list_of_values = []
        function_name = cell.value.split('(')[0][1:]
        arguments = cell.value.split('(')[1].rstrip(')')
        argument_list = arguments.split(':')
        start = argument_list[0]
        end = argument_list[1]
        # Check if the range arguments are valid cell labels
        if start not in self.list_of_cells_labels() or end not in self.list_of_cells_labels():
            return 0
        # Get the row and column indices of the start and end cells
        start_row = self.get_cell_row_by_label(start)
        start_col = self.get_cell_col_by_label(start)
        end_row = self.get_cell_row_by_label(end)
        end_col = self.get_cell_col_by_label(end)
        # Ensure the start cell is before the end cell
        if start_row > end_row:
            start_row, end_row = end_row, start_row
        if ord(start_col) > ord(end_col):
            start_col, end_col = end_col, start_col

        for i in range(start_row, end_row + 1):
            for j in range(ord(start_col), ord(end_col) + 1):
                new_cell_label = chr(j) + str(i)
                # Check if the column label exists in this row
                if new_cell_label in self.list_of_cells_labels():
                    if new_cell_label == cell_label:
                        return '#REF!'
                    cell_value = str(self.get_cell_value_by_label(new_cell_label))
                    # Check if the cell is empty or starts with '#'
                    if cell_value.startswith('#'):
                        return cell_value  # Returns the first matching cell's value
                    if not cell_value:  # Skip empty cells
                        continue
                    list_of_values.append(cell_value)
                else:
                    # Handle case where column label does not exist (e.g., log or ignore)
                    return '0'

        # Convert the list of values to floats and apply the function
        list_of_values = [float(num) for num in list_of_values]
        if function_name == 'SUM':
            return sum(list_of_values)
        if function_name == 'MIN':
            return min(list_of_values)
        if function_name == 'MAX':
            return max(list_of_values)
        if function_name == 'AVERAGE':
            return sum(list_of_values) / len(list_of_values)
        else:
            return "#INVALID_FUNCTION"

    def update_cell_dependencies(self, cell_label):
        """Update the cell dependencies for the given cell."""
        if cell_label in self.dependencies.keys():
            self.dependencies[cell_label] = self.parse_formula(cell_label)
        else:
            self.dependencies[cell_label] = []
            self.dependencies[cell_label] = self.parse_formula(cell_label)
        return self.dependencies

    def update_dependent_cells(self, updated_cell_label):
        """Update the dependent cells of the given cell."""
        if self.check_if_circular_dependency_from_cell(updated_cell_label):
            self.df.at[self.convert_label_to_row_col(updated_cell_label)] = "#CIRCULAR_DEPENDENCY"
        else:
            for cell_label in self.dependencies.keys():
                if updated_cell_label in self.dependencies[cell_label] and cell_label != updated_cell_label:
                    row, col = self.convert_label_to_row_col(cell_label)
                    original_formula = self.get_original_formula(row, col)
                    result = self.evaluate_cell_by_original_formula(cell_label, original_formula)
                    self.df.at[row, col] = result
                    
    def evaluate_cell_by_original_formula(self, cell_label, original_formula):
        """Evaluate the cell using the original formula. We use it for dependent cells."""
        if cell_label not in self.list_of_cells_labels():
            return "#REF!"
        else:
            result = self.evaluate_formula(original_formula, cell_label)
            return result

    def evaluate_cell(self, cell_label):
        """Evaluate the cell with the given label."""
        formula = str(self.get_cell_value_by_label(cell_label))
        if cell_label not in self.list_of_cells_labels():
            return "#REF!"
        else:
            result = self.evaluate_formula(formula, cell_label)
            self.df.at[self.convert_label_to_row_col(cell_label)] = result
            # Trigger updates for dependent cells
            self.update_dependent_cells(cell_label)
            return result

    def evaluate_formula(self, formula, cell_label):
        """Evaluate the formula of the cell and return the result."""
        if not formula or not formula.startswith('='):
            return formula
        if self.is_range_formula(formula):
            cell_value = self.evaluate_range_function(formula, cell_label)
            self.df.at[self.convert_label_to_row_col(cell_label)] = cell_value
            return cell_value
        else:
            cell = self.change_reference_in_cell(cell_label)
            cell_value = cell.main_evaluate_cell()
            self.df.at[self.convert_label_to_row_col(cell_label)] = cell_value
            return cell_value

    def clear_sheet(self):
        """Clear the sheet by resetting all cell values to empty strings."""
        # Initialize a DataFrame with the correct dimensions and empty strings
        data = {chr(ord('A') + j): ['' for _ in range(self.rows)] for j in range(self.columns)}
        self.df = pd.DataFrame(data, index=pd.RangeIndex(start=1, stop=self.rows + 1))
        self.original_df = self.df.copy()
        self.dependencies = {}

    def load_json(self, json_file):
        """Load data from a JSON file into the spreadsheet. The json file can be a list of lists or a dictionary"""
        with open(json_file, 'r') as file:
            data = json.load(file)
        if all(isinstance(subitem, list) for subitem in data):
            # Handle list of lists
            dict_of_values = self.convert_list_to_dict(data)
            self.df = pd.DataFrame(dict_of_values, index=range(1, len(max(dict_of_values.values(), key=len)) + 1))
        elif isinstance(data, dict):
            # Handle dictionary of lists
            self.df = pd.DataFrame(data, index=range(1, len(max(str(data.values()), key=len)) + 1))
        else:
            return

        self.original_df = self.df.copy()
        self.dependencies = {}
        return self.df

    def convert_list_to_dict(self, data):
        """Convert a list of lists to a dictionary."""
        return {chr(ord('A') + j): [data[j][i] for i in range(len(data[j]))] for j in range(len(data))}
