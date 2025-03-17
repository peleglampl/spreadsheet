from tkinter import filedialog, messagebox
from sheet import Sheet
from gui import Window


class SpreadsheetController:
    """Controller class to manage the interaction between the view and the sheet."""
    def __init__(self, view: Window, sheet: Sheet):
        """Initialize the controller with the view and the sheet."""
        self.view = view
        self.sheet = sheet
        self.bind_events()
        self.view.on_clear_callback = self.clear_table

    def bind_events(self):
        """Bind return key events to all cells for evaluation and update."""
        for i in range(1, self.view.rows):  # Skip the label row
            for j in range(1, self.view.columns):  # Skip the label column
                cell_entry = self.view.get_cell(i, j)
                # Binding using a lambda function that captures i and j correctly
                cell_entry.bind('<Return>', lambda event, row=i, col=j: self.on_cell_edit(row, col))

    def on_cell_edit(self, row, col):
        """Handle cell edit event, evaluate the cell, and update the GUI."""
        cell_label = self.convert_row_col_to_label(row, col)
        cell_entry = self.view.get_cell(row, col).get()
        # First, set the cell's value. This might be a direct value or a formula.
        self.sheet.enter_new_value(cell_label, cell_entry)
        evaluated = self.sheet.evaluate_cell(cell_label)
        self.view.update_cell_display(row, col, evaluated)
        self.update_gui_for_dependent_cells(cell_label)
        self.refresh_gui()

    def convert_row_col_to_label(self, row, col):
        """Convert row and column indices to a cell label (e.g., A1)."""
        column_label = chr(ord('A') + col - 1)  # Convert column index to letter
        cell_label = f"{column_label}{row}"
        return cell_label

    def update_gui_for_dependent_cells(self, updated_cell_label):
        """Update the GUI for all cells that depend on the updated cell."""
        # Assuming there's a method in Sheet to get all cells that need updating
        self.sheet.update_dependent_cells(updated_cell_label)

    def convert_label_to_row_col(self, cell_label):
        """Convert a cell label (e.g., A1) to row and column indices."""
        col = ord(cell_label[0]) - ord('A') + 1
        row = int(cell_label[1:])
        return row, col

    def refresh_gui(self):
        """Refresh the GUI with the updated values from the sheet."""
        # Iterate over all cells in the sheet and update their displayed values
        for cell_label in self.sheet.list_of_cells_labels():
            value = self.sheet.get_cell_value_by_label(cell_label)  # Use the evaluated value
            row, col = self.convert_label_to_row_col(cell_label)
            self.view.update_cell_display(row, col, value)

    def clear_table(self):
        """Clear the sheet and refresh the GUI."""
        self.sheet.clear_sheet()
        self.bind_events()
        self.refresh_gui()

    def load_json(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if not file_path:
            return  # User cancelled the dialog
        try:
            df = self.sheet.load_json(file_path)
            if df is None:
                raise Exception("Invalid JSON file")
            if len(df.columns) > self.view.columns or len(df.index) > self.view.rows:
                raise Exception("JSON file dimensions exceed the current sheet size")
            self.refresh_gui_json()
            self.refresh_gui()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load JSON file: {e}")
            self.clear_table()

    def refresh_gui_json(self):
        """Refresh the GUI with the updated values from the sheet."""
        for col in self.sheet.df.columns:
            for row in self.sheet.df.index:
                cell_value = str(self.sheet.df.at[row, col])
                if cell_value == "nan":
                    cell_value = ""
                # Convert the row and column to GUI indices (e.g., 'A1' -> (1, 1))
                gui_row, gui_col = self.convert_label_to_row_col(f"{col}{row}")
                self.view.update_cell_display(gui_row, gui_col, cell_value)
