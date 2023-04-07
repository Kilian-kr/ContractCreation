# First-party modules
import openpyxl
import pandas as pd
from docx2pdf import convert as docx2pdf_convert
from mailmerge import MailMerge

# Third-party modules
import datetime
import os
import re
import shutil
import tkinter as tk
import tkinter.filedialog as filedialog
import tkinter.messagebox as messagebox
import tkinter.ttk as ttk
from typing import Optional


GLOBAL_BG_COLOR = "White"
FILENAME_DEFAULT_TEXT = "Enter filename... You may use information from columns by using {column_name}. Only exact matches work."
TABLE_ROW_START = 8
ILLEGAL_CHARACTER_LIST = ["<", ">", ":", "\"", "\\", "/", "|", "?", "*"]
TEMP_DIR_NAME = ".temp"
LABEL_SETTINGS = {'bg': GLOBAL_BG_COLOR, 'fg': "Black", 'padx': 5, 'pady': 5}
BUTTON_SETTINGS = {"width": 10}
OPTION_MENU_SETTINGS = {"width": 50, 'padx': 5, 'pady': 5}
BORDER_SIZE = 3


class ContractCreationTool:
    """
    Initializes the Grant_Tool class, which creates a GUI window to automate the creation of contracts
    using a Word template and an Excel spreadsheet.
    """

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Contract Creation Tool")
        self.root.geometry("950x900+500+50")
        # Create a canvas widget with a vertical scrollbar
        self.canvas = tk.Canvas(self.root)
        self.scrollbar = tk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.canvas.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        # Create a frame to hold the content of the window
        self.frame = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.frame, anchor="nw")

        self.frame.configure(background=GLOBAL_BG_COLOR)
        self.root.configure(background=GLOBAL_BG_COLOR)
        self.canvas.configure(background=GLOBAL_BG_COLOR)

        self.word_file: str = str()
        self.excel_file: str = str()
        self.output_folder: str = str()
        self.table_labels: list[tk.Label] = []
        self.table_menu: list[tk.OptionMenu] = []
        self.ws_dict: pd.DataFrame = pd.DataFrame()
        self.id: int = 0

        self.create_widgets()
        # Update the scroll region of the canvas to include the frame
        self.frame.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))
        self.root.mainloop()

    def get_help(self) -> None:
        """
        Displays a messagebox that shows the available fields that can be used in the filenames entry field.

        Returns:
            None
        """
        help_msg: str = "Currently available Fields:\n "
        for col in list(self.ws_dict):
            if "Unnamed" not in col:
                help_msg += f"\n {{{col}}}"
                check_bool, check_list = contains_illegal_char(col)
                if check_bool:
                    help_msg += f" <= Cannot be used due to illegal characters: {' '.join(check_list)}"
        if self.ws_dict.empty:
            help_msg += "\n No Columns loaded/found!"

        help_msg += f"\n\n\n These characters cannot be used in the filename:\n {' '.join(ILLEGAL_CHARACTER_LIST)} "

        messagebox.showinfo("Help", help_msg)

    def extract_columns(self) -> list[str]:
        """
        Extracts all column names from the filenames entry field.

        Returns:
        List: column names.
        """
        try:
            columns_found_ck = [col.strip() for col in re.findall("{(.+?)}", self.filenames.get())]
        except AttributeError:
            columns_found_ck = []
        return columns_found_ck

    def generate_filename(self, row_id: int) -> str:
        """
        Generates a filename for the contract that includes the current date, an ID, and the provided filename (with variables if applicable).
        Args:
            row_id (int): The row number from the Excel spreadsheet that corresponds to the contract to be generated.
        Returns:
            str: Filename as a string.
        """
        id: str = f"{str(datetime.datetime.today().strftime('%H%M%S%f'))}{str(row_id)}"
        file: str = self.filenames.get()
        filetype: str = ".docx"
        if file == FILENAME_DEFAULT_TEXT or not file:
            return f"{get_date_file()} - Autocreation for {self.excel_file.split('/')[-1]}  - {id}{filetype}"

        columns_found_fn: list[str] = list[str]
        columns_found_fn = self.extract_columns()
        if columns_found_fn != "":
            for column in columns_found_fn:
                if column in list(self.ws_dict):
                    file: str = file.replace("{" + column + "}", self.ws_dict[column][row_id])
        return f"{get_date_file()} - {file} - {id}{filetype}"

    def check_columns(self) -> tuple[bool, Optional[list[str]]]:
        """
        Checks if all column names in the filenames entry field match the column names in the worksheet dictionary.
        Returns:
            Tuple: containing a boolean value indicating whether all columns are present, and a list of any missing column names.
        """
        bool_check: bool = True
        cols: list[str] = []
        if self.filenames.get() == FILENAME_DEFAULT_TEXT:
            return bool_check, None
        for col in self.extract_columns():
            if col not in list(self.ws_dict):
                bool_check = False
                if not cols:
                    cols = [col]
                else:
                    cols.append(col)
        return bool_check, cols

    def check_filename(self) -> str:
        """
        Checks if the filename contains any unknown columns or illegal characters.
        Returns:
          str: Error Message.
        """
        return_msg: str = ""
        col_check: bool
        col_found: Optional[list[str]]
        col_check, col_found = self.check_columns()
        if not col_check:
            for cols in col_found:
                return_msg += f"Unknown column: {{{cols}}}\n"

        try:
            filename_check_bool, filename_check_list = contains_illegal_char(self.filenames.get())
            if filename_check_bool:
                for char in filename_check_list:
                    return_msg += f"\nIllegal character found in filename: {char}"
        except TypeError as e:
            return f"Error: {str(e)}"
        if return_msg != "":
            return_msg += "\n\n See Help for more information."
        return return_msg

    def on_entry_click(self, event: tk.Event) -> None:
        """Function that gets called whenever the filenames entry is clicked.

        Args:
            event (tkinter.Event): The event object passed by tkinter.

        Returns:
            None
        """
        if self.filenames.get() == FILENAME_DEFAULT_TEXT:
            self.filenames.delete(0, "end")  # delete all the text in the entry
            self.filenames.insert(0, "")  # Insert blank for user input
            self.filenames.config(fg="black", border=BORDER_SIZE)

    def on_focusout(self, event: tk.Event) -> None:
        """Function that gets called whenever the filenames entry loses focus.

        Args:
            event (tkinter.Event): The event object passed by tkinter.
        """
        if self.filenames.get() == "":
            self.filenames.insert(0, FILENAME_DEFAULT_TEXT)
            self.filenames.config(fg="grey", border=BORDER_SIZE)

    def create_widgets(self) -> None:
        """
        Creates all the widgets for the GUI.

        Returns:
            None
        """

        # create labels
        tk.Label(self.frame, text="Select Word file:", **LABEL_SETTINGS).grid(row=0, column=0, )
        tk.Label(self.frame, text="Select Excel file:", **LABEL_SETTINGS).grid(row=1, column=0,)
        tk.Label(self.frame, text="Select Output folder:", **LABEL_SETTINGS).grid(row=2, column=0,)
        tk.Label(self.frame, text=f"Filenames:                 \"{get_date_file()} - ", **LABEL_SETTINGS).grid(row=6, column=0)
        tk.Label(self.frame, text="- ID.pdf\"", **LABEL_SETTINGS).grid(row=6, column=2)

        # create text fields
        self.word_file_entry: tk.Entry = tk.Entry(self.frame, width=100, state="readonly")
        self.word_file_entry.grid(row=0, column=1, padx=10, pady=10)

        self.excel_file_entry: tk.Entry = tk.Entry(self.frame, width=100, state="readonly")
        self.excel_file_entry.grid(row=1, column=1, padx=10, pady=10)

        self.output_folder_entry: tk.Entry = tk.Entry(self.frame, width=100, state="readonly")
        self.output_folder_entry.grid(row=2, column=1, padx=10, pady=10)

        self.filenames: tk.Entry = tk.Entry(self.frame, width=100)
        self.filenames.grid(row=6, column=1)
        self.filenames.insert(0, FILENAME_DEFAULT_TEXT)
        self.filenames.config(fg="grey", border=BORDER_SIZE)
        self.filenames.bind("<FocusIn>", self.on_entry_click)
        self.filenames.bind("<FocusOut>", self.on_focusout)

        # create label for progress bar value
        self.progress_bar_value_label_var = tk.StringVar()
        self.progress_bar_value_label = tk.Label(
            self.frame, textvariable=self.progress_bar_value_label_var, bg="#e8e4e4", font=("Arial", 7), pady=0, padx=0)
        self.progress_bar_value_label.grid(row=3, column=1)

        # create label for current task
        self.current_task_var = tk.StringVar()
        self.current_task_label = tk.Label(self.frame, textvariable=self.current_task_var, **LABEL_SETTINGS)
        self.current_task_label.grid(row=4, column=1)

        tk.Button(self.frame, text="Select", command=self.select_word_file, padx=10,
                  **BUTTON_SETTINGS).grid(row=0, column=2, padx=10, pady=10)
        tk.Button(self.frame, text="Select", command=self.select_excel_file, padx=10,
                  **BUTTON_SETTINGS).grid(row=1, column=2, padx=10, pady=10)
        tk.Button(self.frame, text="Select", command=self.select_output_folder, padx=10,
                  **BUTTON_SETTINGS).grid(row=2, column=2, padx=10, pady=10)
        tk.Button(self.frame, text="Load", command=self.load_data, padx=10,
                  **BUTTON_SETTINGS).grid(row=3, column=2, padx=10, pady=10)
        tk.Button(self.frame, text="Generate", command=self.generate_files, padx=10,
                  **BUTTON_SETTINGS).grid(row=5, column=1, padx=10, pady=10)
        tk.Button(self.frame, text="Help", command=self.get_help, padx=10,
                  **BUTTON_SETTINGS).grid(row=3, column=0, padx=10, pady=10)

        # create progress bar
        self.progress_bar: ttk.Progressbar = ttk.Progressbar(self.frame, orient="horizontal", length=400, mode="determinate")
        self.progress_bar.grid(row=3, column=1, pady=10)
        self.progress_bar["value"]: int = 0

        # create table
        tk.Label(self.frame, text="Template Fields", font=("Segoe UI", 9, "bold",
                                                           "roman", "underline"), **LABEL_SETTINGS).grid(row=7, column=0)
        tk.Label(self.frame, text="Data Fields", font=("Segoe UI", 9, "bold",
                                                       "roman", "underline"), **LABEL_SETTINGS).grid(row=7, column=1)

    def select_word_file(self) -> None:
        """
        Open a file dialog to select a Word file and update the corresponding entry field with the selected file path.

        Returns:
            None
        """
        self.word_file: filedialog = filedialog.askopenfilename(
            initialdir=os.getcwd(), title="Select Word file", filetypes=[("Word files", "*.docx")])
        self.word_file_entry.configure(state="normal")
        self.word_file_entry.delete(0, "end")
        self.word_file_entry.insert(0, self.word_file)
        self.word_file_entry.configure(state="readonly")

    def select_excel_file(self) -> None:
        """
        Open a file dialog to select an Excel file and update the corresponding entry field with the selected file path.

        Returns:
            None
        """
        self.excel_file: filedialog = filedialog.askopenfilename(
            initialdir=os.getcwd(), title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
        self.excel_file_entry.configure(state="normal")
        self.excel_file_entry.delete(0, "end")
        self.excel_file_entry.insert(0, self.excel_file)
        self.excel_file_entry.configure(state="readonly")

    def select_output_folder(self) -> None:
        """
        Open a file dialog to select an output folder and update the corresponding entry field with the selected folder path.

        Returns:
            None
        """
        self.output_folder: filedialog = filedialog.askdirectory(initialdir=os.getcwd(), title="Select Output folder")
        self.output_folder_entry.configure(state="normal")
        self.output_folder_entry.delete(0, "end")
        self.output_folder_entry.insert(0, self.output_folder)
        self.output_folder_entry.configure(state="readonly")

    def reset_tables_menus(self):
        # Reset table labels and menus
        self.current_row: int = TABLE_ROW_START
        if self.table_labels:
            for label in self.table_labels:
                label.destroy()
            for menu in self.table_menu:
                menu.destroy()
        self.table_labels = []
        self.table_menu = []

    def load_data(self) -> None:
        """Loads data from selected Word and Excel files, and creates dropdown menus for mail merge fields

        Returns:
            None
        """

        # Check if Word and Excel files have been selected
        if not self.word_file or not self.excel_file:
            messagebox.showerror("Error", "Please select Word and Excel files!")
            return
        self.reset_tables_menus()

        # Read the data from the Excel file
        self.wb: openpyxl.Workbook = openpyxl.load_workbook(self.excel_file)
        self.ws: openpyxl.Worksheet = self.wb.active
        self.ws_dict = pd.read_excel(self.excel_file)

        # Get the Excel file headings
        excel_headings: list[str] = [cell.value for cell in next(self.ws.rows)]
        excel_headings.append("Leave Empty")
        excel_headings.append(f"EXTRA - Add Current Date => {get_date_field()}")

        self.excel_headings = ["Empty Column Name (This Column cannot be mapped)" if x == None else x for x in excel_headings]

        # Read the merge fields from the Word file
        document: MailMerge = MailMerge(self.word_file)
        merge_fields: list[str] = sorted(document.get_merge_fields())

        # Create a dropdown menu for each merge field
        self.fields: dict[str, tk.StringVar] = {}
        retrieved_elements: list[str] = list()
        for field in merge_fields:
            # Create a variable to store the selected option
            self.fields[field] = tk.StringVar()

            # Set the default value of the variable
            retrieved_elements = list(filter(lambda x: field.upper() in x.upper(), self.excel_headings))
            self.fields[field].set(retrieved_elements[0] if retrieved_elements else self.excel_headings[0])

            # Create the drop-down menu
            self.table_labels.append(tk.Label(self.frame, text=field, **LABEL_SETTINGS))
            self.table_menu.append(tk.OptionMenu(self.frame, self.fields[field], *self.excel_headings))

        for menu in self.table_menu:
            menu.config(OPTION_MENU_SETTINGS)

        for label in self.table_labels:
            label.grid(row=self.current_row, column=0, padx=10, pady=10)
            self.table_menu[self.current_row-TABLE_ROW_START].grid(row=self.current_row, column=1, padx=10, pady=10)
            self.current_row += 1

        self.frame.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def generate_files(self) -> None:
        """Generates output files by iterating over the data in the worksheet
        and creating a docx file for each row of data and the converting it to pdf.

        Returns:
            None
        """
        # check if an output folder has been selected
        if not self.output_folder:
            messagebox.showerror("Error", "Please select an output directory!")
            return

        # get the number of rows in the worksheet and check the validity of the filename
        rows: int = len(self.ws_dict.index)
        r_i_check: str = self.check_filename()
        if r_i_check != "":
            messagebox.showerror("Error", r_i_check)
            return

        # initialize variables for progress tracking and time measurements
        start_time: datetime = datetime.datetime.now()
        end_time: datetime = datetime.datetime.min
        time_taken: datetime = datetime.datetime.min
        avg_time_per_file: float = float()
        self.progress_bar["value"]: int = 0  # reset progress bar
        # set the maximum value of the progress bar based on the number of rows
        self.progress_bar["maximum"] = (rows) * 2 + 1
        filename: str = str()

        # create a mapping of field names to their values for each row
        self.mappings: dict[str, str] = {}
        for field, value in self.fields.items():
            self.mappings[field] = value.get()

        # iterate over each row in the worksheet and create a docx file for it
        for row_id in range(rows):
            temp: dict[str, str] = {}
            for values in self.mappings:
                # handle cases where the column is empty or the mapping is invalid
                if self.mappings[values] == "Leave Empty" or self.mappings[values] == "Empty Column Name (This Column cannot be mapped)":
                    temp[values] = str()
                    continue
                if self.mappings[values] in list(self.ws_dict) == "nan":
                    if str(self.ws_dict[self.mappings[values]][row_id]) == "nan":
                        temp[values] = str()
                        continue

                # handle cases where the field is a special value (e.g. a date field)
                if self.mappings[values] == f"EXTRA - Add Current Date => {get_date_field()}":
                    temp[values] = get_date_field()
                    continue

                temp[values] = str(self.ws_dict[self.mappings[values]][row_id])

            # generate the filename for the current row and create a docx file for it
            filename = self.generate_filename(row_id)
            self.create_docx_file(temp, clean_filename(filename))

        # convert all docx files to pdf and remove the docx files
        docx2pdf_convert(self.docx_path, self.output_folder)
        shutil.rmtree(self.docx_path)

        # update the progress bar and display a success message with time measurements
        self.update_progress_bar("PDF Files created Succesfully")
        end_time = datetime.datetime.now()
        time_taken = end_time - start_time
        avg_time_per_file = round(time_taken.total_seconds() / rows, 2)
        tk.messagebox.showinfo("Success", f"{rows} PDF files successfully generated after {str(time_taken).split('.')[0]}."
                               + f"\n (~{avg_time_per_file:2f} seconds per file) ")

        # reset the GUI elements and progress bar
        self.filenames.delete(0, "end")
        self.canvas.focus_set()
        self.update_progress_bar()

    def create_docx_file(self, individ_mappings: dict[str, str], filename: str) -> None:
        """
        Create a Docx file with the given mappings and save it as a PDF file.

        Args:
            individ_mappings (Dict): A dictionary containing individual mappings for the given Docx file.
            filename (str): Name of the output file.

        Returns:
            None
        """

        # Increment ID and update progress bar
        self.id += 1
        self.update_progress_bar(f"Starting to create Word File for {filename}...")

        # Create temp directory if it doesn't exist
        self.docx_path: Optional[str] = os.path.join(self.output_folder, TEMP_DIR_NAME)
        os.makedirs(self.docx_path, exist_ok=True)

        # Merge templates and save as a Docx file
        docx_file_path = os.path.join(self.docx_path, filename)

        with MailMerge(self.word_file) as new_document:
            new_document.merge_templates([individ_mappings], separator="page_break")
            new_document.write(docx_file_path)

        # Update progress bar
        self.update_progress_bar(f"Created Word File for {filename}...")

    def update_progress_bar(self, label: str = None) -> None:
        """
        Updates the progress bar and sets the current task label.

        Args:
            label (str): Optional label to set for the current task.

        Returns:
            None
        """
        self.progress_bar["value"] += 1
        if label is not None:
            self.current_task_var.set(label)

        percentage = round((self.progress_bar['value'] / self.progress_bar['max']) * 100, 2)
        self.progress_bar_value_label_var.set(f"{percentage}%")
        if percentage >= 50.00:
            self.progress_bar_value_label.config(background="#08b424", fg="White")
        else:
            self.progress_bar_value_label.config(background="#e8e4e4", fg="Black")

        self.frame.update_idletasks()
        print(str(label))


def get_date_file() -> str:
    """
    Returns the current date in yymmdd format as a string.
    """
    return str(datetime.datetime.today().strftime("%y%m%d"))


def get_date_field() -> str:
    """
    Returns the current date in dd/mm/yyyy format as a string.
    """
    return str(datetime.datetime.today().strftime("%d/%m/%Y"))


def clean_filename(input_text: str) -> str:
    """
    Cleans the an input string, removing ILLEGAL_CHARACTERS(["<", ">", ":", "\"", "\\", "/", "|", "?", "*"])


    Args:
        input_text (str): String that needs to be cleaned.

    Returns:
        str: cleaned input
     """
    for character in ILLEGAL_CHARACTER_LIST:
        input_text = input_text.replace(character, "")
    return input_text


def contains_illegal_char(input_text: str) -> tuple[bool, Optional[list[str]]]:
    check_bool: bool = False
    illegal_chars: list[str] = []
    for character in ILLEGAL_CHARACTER_LIST:
        if character in input_text:
            check_bool = True
            illegal_chars.append(character)

    return check_bool, illegal_chars


if __name__ == '__main__':
    app = ContractCreationTool()

