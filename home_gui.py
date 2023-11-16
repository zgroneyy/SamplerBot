from tkinter import *
from tkinter import filedialog
import os
from methods import get_first_row_cells
import openpyxl

def browseFiles():
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select a File",
                                          filetypes=(("All files", "*.*"),
                                                     ("Text files", "*.txt"),
                                                     ("CSV files", "*.csv"),
                                                     ("Excel files", "*.xlsx"),
                                                     ("All files", "*.*")))

    label_file_explorer.configure(text="Opened: " + filename)

    # Read the first row cells of the selected file
    first_row_cells = get_first_row_cells(filename)

    # Update label_headings with the first row cells
    label_headings.configure(text=" | ".join(map(str, first_row_cells.values())))

    # Update dropdown menu
    update_dropdown_menu(first_row_cells)

def update_dropdown_menu(first_row_cells):
    global dropdown_menu  # Ensure dropdown_menu is a global variable

    # Clear the existing menu items
    dropdown_menu['menu'].delete(0, 'end')

    # Set the background color of the dropdown menu to gray
    dropdown_menu.configure(bg="gray")

    # Add "----------" as the default value
    dropdown_menu['menu'].add_command(label="----------", command=lambda: dropdown_variable.set("----------"))

    # Add elements of first_row_cells to the dropdown menu
    for key, value in first_row_cells.items():
        dropdown_menu['menu'].add_command(label=f"{key}: {value}", command=lambda v=value: dropdown_variable.set(v))

def on_exit():
    os._exit(0)

def create_gui():
    window = Tk()

    window.title('File Explorer')
    window.geometry("600x250")
    window.resizable(False, False)
    window.config(background="gray")

    global label_file_explorer
    label_file_explorer = Label(window,
                                text="Select a file to open:",
                                width=40,
                                height=2,
                                fg="black",
                                bg="white")

    button_explore = Button(window,
                            text="Browse",
                            command=browseFiles,
                            bg="gray",
                            fg="black",
                            highlightbackground="gray",
                            highlightcolor="gray")

    button_exit = Button(window,
                         text="Exit",
                         command=on_exit,
                         bg="gray",
                         fg="red",
                         highlightbackground="gray",
                         highlightcolor="gray")

    global label_headings
    label_headings = Label(window,
                          text="[Headings of your data to be shown here...]",
                          width=40,
                          height=2,
                          fg="black",
                          bg="gray")

    label_file_explorer.grid(row=1, column=0, sticky="w", padx=(5, 0), pady=(30, 0))
    button_explore.grid(row=1, column=1, sticky="e", padx=(0, 20), pady=(25, 0))
    label_headings.grid(row=2, column=0, sticky="w", padx=(5, 0), pady=(0, 0))

    # Create a StringVar to store the selected value from the dropdown menu
    global dropdown_variable
    dropdown_variable = StringVar()

    # Create a dropdown menu with "----------" as the default value and gray background
    global dropdown_menu
    dropdown_menu = OptionMenu(window, dropdown_variable, "----------")
    dropdown_menu.configure(bg="gray")
    dropdown_menu.grid(row=3, column=0, sticky="e", padx=(20, 0), pady=(0, 0))

    # Create a StringVar to store the selected value from the group_by_dropdown menu
    global group_by_variable
    group_by_variable = StringVar()

    # Create a group_by_dropdown menu with "Random, Stratified, Systematic, Cluster" options
    group_by_dropdown = OptionMenu(window, group_by_variable, "Random", "Stratified", "Systematic", "Cluster")
    group_by_dropdown.configure(bg="gray")
    group_by_dropdown.grid(row=3, column=1, sticky="w", padx=(20, 0), pady=(0, 0))

    # Create the "GO!" button
    go_button = Button(window,
                       text="GO!",
                       command=on_go,
                       bg="gray",
                       fg="black",
                       highlightbackground="gray",
                       highlightcolor="gray")
    go_button.grid(row=4, column=1, sticky="w", padx=(12, 20), pady=(10, 0))

    button_exit.grid(row=5, column=1, sticky="w", padx=(12, 20), pady=(10, 0))

    window.mainloop()

def on_go():
    # Implement the functionality for the "GO!" button here
    pass
