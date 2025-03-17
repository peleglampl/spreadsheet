import sys
from sheet import Sheet
from gui import Window
from spreadsheet_controller import SpreadsheetController
import tkinter as tk
from tkinter.simpledialog import askinteger, askstring
import argparse


def get_user_input(root):
    """Get the number of rows and columns for the spreadsheet from the user."""
    while True:
        rows = askinteger("Input", "How many rows do you want in your spreadsheet?", parent=root,
                          minvalue=1, maxvalue=100)
        if rows is None:
            # Handle the case where user presses Enter without input or cancels
            root.destroy()
            return None, None
        break  # Exit the loop if rows is not None

    while True:
        columns = askinteger("Input", "How many columns do you want in your spreadsheet?", parent=root,
                             minvalue=1, maxvalue=26)
        if columns is None:
            # Handle the case where user presses Enter without input or cancels
            root.destroy()
            return None, None
        break  # Exit the loop if columns is not None

    return rows, columns


def get_name_to_sheet(root):
    """Get the name of the sheet from the user."""
    name = askstring('Input', "Enter the name of the sheet: ", parent=root)
    return name


def help():
    print("Usage: main.py ")
    print("Launch the spreadsheet application.")
    print("Options:")
    print("  --helper    display this help and exit")
    print("For GUI mode, just run without any options.")
    sys.exit()


def controller():
    """Main controller function to set up the spreadsheet application."""
    root = tk.Tk()
    rows, columns = get_user_input(root)  # Example size, could be dynamic via user input
    if rows is None or columns is None:
        return
    name = get_name_to_sheet(root)
    root.title(f"Spreadsheet App - {name}")
    # Create a new Sheet instance with the specified rows and columns
    sheet = Sheet(rows, columns)
    sheet.original_df = sheet.df.copy()
    # this method sets up the GUI
    view = Window(root, rows, columns, name)
    controller_sheet = SpreadsheetController(view=view, sheet=sheet)
    view.set_controller(controller_sheet)
    controller_sheet.bind_events()
    root.mainloop()


def main():
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument('--help', action='store_true', help='Show custom help message and exit')
    args = parser.parse_args()

    if args.help:
        help()
        sys.exit()

    if len(sys.argv) >= 1:
        controller()


if __name__ == "__main__":
    main()
