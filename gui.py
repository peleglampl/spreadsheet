
import tkinter as tk
from tkinter import Frame, Canvas, Scrollbar, Entry, Label, StringVar, simpledialog, messagebox, filedialog, font
from tkinter.simpledialog import askstring
from tkinter.colorchooser import askcolor
from pandas import DataFrame


class Window:
    """Class to represent the main window of the spreadsheet application."""
    def __init__(self, master_window, rows, columns, name):
        """Initialize the main window with the specified rows, columns, and name."""
        self.master = master_window
        self.master.geometry("800x800")
        self.rows = rows
        self.columns = columns
        self.name = name
        self.cells = []
        self.initialize_ui()

    def initialize_ui(self):
        """Initialize the UI components."""
        self.create_main_frame()
        self.create_top_frame()
        self.create_middle_frame()
        self.create_scrollbars()
        self.create_cells()
        self.create_menu(self.master)

    def create_main_frame(self):
        """Create the main frame to hold other widgets."""
        self.main = Frame(self.master)
        self.main.pack(expand=True, fill=tk.BOTH, side=tk.TOP)

    def create_top_frame(self):
        """Create the top frame to hold the menu and toolbar."""
        self.top = Frame(self.main)
        self.top.pack(fill=tk.X, pady=(0, 5))

    def create_middle_frame(self):
        """Create the middle frame to hold the spreadsheet cells."""
        self.middle = Frame(self.main)
        self.middle.pack(padx=5, pady=(5, 5), fill=tk.BOTH, expand=True)

    def create_scrollbars(self):
        """Create and set up horizontal and vertical scrollbars."""
        self.canvas = Canvas(self.middle, bg="white", width=800, height=800)
        self.scrollbar_y = Scrollbar(self.middle, orient="vertical", command=self.canvas.yview)
        self.scrollbar_x = Scrollbar(self.middle, orient="horizontal", command=self.canvas.xview)

        self.canvas.configure(yscrollcommand=self.scrollbar_y.set, xscrollcommand=self.scrollbar_x.set)
        self.canvas.bind('<Configure>', lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        self.scrollbar_y.pack(side="right", fill="y")
        self.scrollbar_x.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.middle = Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.middle, anchor="nw")

    def create_cells(self):
        """Create spreadsheet cells."""
        self.rows += 1  # Adjusting for label row
        self.columns += 1  # Adjusting for label column
        self.cells = [[None for _ in range(self.columns)] for _ in range(self.rows)]

        for i in range(self.rows):
            for j in range(self.columns):
                label_text = chr(ord('A') + (j - 1)) if i == 0 and j != 0 else str(i) if j == 0 and i != 0 else ""
                var = StringVar()
                cell = Entry(self.middle, width=7, textvariable=var) if i != 0 and j != 0 else\
                    Label(self.middle, text=label_text, width=7, fg='black', bg='white')
                cell.grid(row=i, column=j)
                if i != 0 and j != 0:
                    self.cells[i][j] = cell
                if label_text:  # For label cells
                    cell.config(text=label_text)

    def get_cell(self, row, col):
        """Get the cell widget at the specified row and column."""
        if row < 0 or row >= self.rows or col < 0 or col >= self.columns:
            return ''
        else:
            return self.cells[row][col]

    def create_menu(self, master):
        """Create the menu bar for the main window."""
        self.menu = tk.Menu(master)
        master.config(menu=self.menu)
        # File menu
        file_menu = tk.Menu(self.menu, tearoff=0)
        self.menu.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Save to JSON", command=self.save_to_json)
        file_menu.add_command(label="Load from JSON", command=self.load_json)
        file_menu.add_command(label="Import to excel", command=self.import_to_excel)
        file_menu.add_command(label="Errors Index", command=self.error_index)
        file_menu.add_command(label="Exit", command=master.quit)

        # Edit menu
        edit_menu = tk.Menu(self.menu, tearoff=0)
        self.menu.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="Rename", command=self.rename_sheet)
        edit_menu.add_command(label="Clear Sheet", command=self.clear_table)
        edit_menu.add_command(label="Search", command=self.search_value)
        edit_menu.add_command(label="Change Cell Color", command=self.change_cell_color)
        edit_menu.add_command(label="Change Cell Font", command=self.change_cell_font)

    def update_cell_display(self, row, column, value):
        """Update the display of the cell at the specified row and column with the given value."""
        if self.cells[row][column]:
            self.cells[row][column].delete(0, tk.END)
            self.cells[row][column].insert(0, value)

    def refresh_gui(self):
        """Refresh the GUI with the updated values from the sheet."""
        for i in range(1, self.rows):
            for j in range(1, self.columns):
                if self.cells[i][j]:
                    value = self.cells[i][j].get()
                    self.update_cell_display(i, j, value)

    def import_to_excel(self):
        """Export the data to an Excel file."""
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            pd = self.change_to_df()
            pd.to_excel(filepath, engine='openpyxl')
        else:
            return None

    def save_to_json(self):
        """Save the data to a JSON file."""
        filepath = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if filepath:
            print(f"Saving to JSON: {filepath}")
        pd = self.change_to_df()
        pd.to_json(filepath)

    # The following methods: set_controller, load_json are used for LOAD JSON:
    def set_controller(self, controller):
        """Set the controller for the view."""
        self.controller = controller

    def load_json(self):
        """Trigger the JSON loading process through the controller."""
        self.controller.load_json()

    def change_to_df(self):
        # Export the data to a DataFrame
        data = []
        for i in range(1, self.rows):
            row = []
            for j in range(1, self.columns):
                if self.cells[i][j]:
                    row.append(self.cells[i][j].get())
            data.append(row)
        columns = [chr(65 + j) for j in range(self.columns - 1)]  # 65 is ASCII for 'A'
        # Create DataFrame
        df = DataFrame(data, columns=columns)
        # Adjust the DataFrame index to start from 1 instead of the default 0
        df.index = range(1, len(df) + 1)
        return df

    def quit(self):
        """Quit the application."""
        self.master.quit()

    def rename_sheet(self):
        """Rename the sheet."""
        # This method uses `askstring` to prompt the user for a new name
        new_name = askstring("Rename Sheet", "Enter the new name of the sheet:")
        if new_name:
            self.name = new_name  # Update the internal name
            self.master.title(f"Spreadsheet App - {self.name}")

    def clear_table(self):
        """Clear the table and reset the GUI."""
        for i in range(1, self.rows):
            for j in range(1, self.columns):
                if self.cells[i][j]:
                    self.cells[i][j].config(bg='white')
                    self.cells[i][j].config(font=font.Font(family='Arial', size=8, weight='normal'))
                    self.update_cell_display(i, j, "")
        if hasattr(self, 'on_clear_callback'):
            self.on_clear_callback()

    def search_value(self):
        """Search for a value in the table."""
        # Prompt the user to enter a search string
        search_query = simpledialog.askstring("Search", "Enter the value to search:")
        if search_query:  # If the user entered a value
            self.find_in_table(search_query)

    def find_in_table(self, search_query):
        """Find and highlight the cells with the search query."""
        found = False
        for i, row in enumerate(self.cells):
            for j, cell in enumerate(row):
                if cell and cell.cget('state') == 'normal':  # Make sure the cell is editable (not a label)
                    if search_query.lower() in cell.get().lower():  # Case-insensitive search
                        cell.config(bg='yellow')  # Highlight the cell
                        found = True
                    else:
                        cell.config(bg='white')  # Reset the background
        if not found:
            messagebox.showinfo("Search", "Value not found.")

    def change_cell_color(self):
        """Change the background color of a cell."""
        cell_address = simpledialog.askstring("Change Cell Color", "Enter cell address (e.g., B2):")
        if cell_address:
            color = askcolor(title="Choose a color")[1]  # This returns a tuple, where the hex color is the second item
            if color:
                self.set_cell_color(cell_address, color)

    def set_cell_color(self, address, color):
        """Set the background color of a cell at the specified address."""
        if str(address) not in self.all_cell_labels():
            messagebox.showinfo("Error", "Invalid cell address.")
            return

        column, row = self.parse_address(address)
        if 0 < row <= self.rows and 0 < column <= self.columns:
            self.cells[row][column].config(bg=color)
        else:
            messagebox.showinfo("Error", "Invalid cell address.")

    def parse_address(self, address):
        """Parse a cell address (e.g., B2) into column and row indices."""
        column = ord(address[0].upper()) - ord('A') + 1
        row = int(address[1:])
        return column, row

    def change_cell_font(self):
        """Change the font of a cell."""
        cell_address = simpledialog.askstring("Change Cell Font", "Enter cell address (e.g., B2):")
        if str(cell_address) not in self.all_cell_labels():
            messagebox.showinfo("Error", "Invalid cell address.")
            return
        column, row = self.parse_address(cell_address)
        if row < 0 or row > self.rows or column < 0 or column > self.columns:
            messagebox.showinfo("Error", "Invalid cell address.")
        if cell_address:
            font_family = simpledialog.askstring("Font Family", "Enter font family (e.g., Arial, David):")
            if font_family not in font.families():
                messagebox.showerror("Font Error", "Invalid font family.")
                return
            font_size = simpledialog.askstring("Font Size", "Enter font size (e.g., 12-34):")
            if not font_size.isdigit() or int(font_size) <= 0 or int(font_size) > 34:
                messagebox.showerror("Font Error", "Invalid font size.")
                return
            font_style = simpledialog.askstring("Font Style", "Enter font style (e.g., bold/ normal):")
            if font_style.lower() not in ['bold', 'normal']:
                messagebox.showerror("Font Error", "Invalid font style.")
                return
            self.set_cell_font(cell_address, font_family, font_size, font_style)

    def set_cell_font(self, address, family, size, style):
        """Set the font of a cell at the specified address."""
        column, row = self.parse_address(address)
        if row > 0 and row < self.rows and column > 0 and column < self.columns:
            try:
                cell_font = font.Font(family=family, size=int(size), weight=style.lower())
                self.cells[row][column].config(font=cell_font)
            except ValueError as e:
                messagebox.showerror("Font Error", str(e))
        else:
            messagebox.showinfo("Error", "Invalid cell address.")

    def error_index(self):
        """Display the errors index."""
        # This method will be called when the button is clicked
        new_window = tk.Toplevel(self.master)
        new_window.title("Errors Index")
        new_window.geometry("300x300")
        # Add widgets to the new window
        label = tk.Label(new_window, text="Errors Index:")
        label.pack(pady=10)
        # Example of displaying a list of errors
        dict_of_errors = {
            '#ERROR': 'Not possible to calculate the formula.',
            '#DIV/0!': 'division in zero is not allowed.',
            '#VALUE!': 'Function argument is of the wrong type.',
            '#REF!': 'Reference is invalid.',
            '#TYPE_ERROR': 'Type error.',
            '#INVALID_FUNCTION': 'Invalid function. Supported functions: MIN, MAX, SUM, AVERAGE.',
            '#N/A': 'Wrong number of arguments to the function. For example:'
                    'for SQRT function, only one argument is allowed.'
                    'For MOD function, two arguments are allowed.'
                    'For SUM, AVERAGE, MIN, MAX functions, at least one argument is required.',
            '#CIRCULAR_DEPENDENCY': 'Circular dependency detected.'
        }
        for error, description in dict_of_errors.items():
            label = tk.Label(new_window, text=f"{error}: {description}")
            label.pack()

    def all_cell_labels(self):
        list_of_cells = []
        for i in range(1, self.rows):
            for j in range(1, self.columns):
                list_of_cells.append(f"{chr(64 + j)}{i}")
        return list_of_cells
