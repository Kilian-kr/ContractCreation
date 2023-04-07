**Documentation for Grant\_Tool**

The Grant\_Tool class represents a tool that creates contracts using a Word template and data from an Excel file. The tool allows the user to select a Word template, an Excel file containing data for the contracts, and an output folder to save the resulting contracts. The tool also allows the user to specify the filenames for the output contracts, either by entering a static name or by using data from the Excel file.

**Libraries and Dependencies**

- **tkinter** : Standard Python library for creating graphical user interfaces (GUI).
- **filedialog** : A module of tkinter used for opening and saving files.
- **ttk** : A module of tkinter that provides enhanced widgets.
- **os** : A module that provides a way of interacting with the operating system.
- **pandas** : A library used for data manipulation and analysis.
- **openpyxl** : A library for reading and writing Excel (.xlsx, .xlsm, .xltx, and .xltm) files.
- **mailmerge** : A library for generating emails and documents.
- **shutil** : A module for file operations that provides an interface to several low-level Unix system calls.
- **datetime** : A module for working with dates and times.
- **pathlib** : A module that provides a platform-independent way of working with the filesystem.
- **docx2pdf** : A module for converting Word (.docx) files to PDF files.
- **re** : A module that provides support for regular expressions.
- **typing** : A module that provides support for type hints and annotations.

**Global Variables**

- FILENAME\_TEXT: A string that represents the default text displayed in the "Filename" field.
- GEN\_ROW\_START: An integer that represents the row in the Excel file where the data starts.
- ILLEGAL\_CHARS: A list of strings that represent illegal characters in a filename.

**Class Definition**

**Grant\_Tool**

This class is responsible for creating a GUI that allows the user to select a Word file, an Excel file, and an output folder. The user can also enter a filename template for the generated documents. The class provides several methods that are used to generate the documents using the data from the Excel file and the Word file as a template.

### Methods

#### \_\_init\_\_(self) -\> None

This method initializes an instance of the Grant\_Tool class. It sets up the main window of the tool, including a canvas widget with a vertical scrollbar and a frame to hold the contents of the window. It also initializes several attributes, including word\_file, excel\_file, output\_folder, table\_labels, table\_menu, and docx\_thread. Finally, it calls the create\_widgets method to create the widgets that make up the content of the main window.

#### create\_widgets(self) -\> None

This method creates the widgets that make up the content of the main window of the tool. Specifically, it creates several Label widgets to display instructions to the user, and several Button widgets to allow the user to select the Word template, Excel file, and output folder. It also creates several Label widgets and OptionMenu widgets to display the data from the Excel file in a table format.

#### get\_help(self) -\> None

This method displays a message box to the user that lists the available fields that can be used to create filenames for the output contracts. The available fields are determined by the columns in the selected Excel file.

#### get\_columns(self) -\> List[str]

This method returns a list of strings representing the column names in the selected Excel file that are used to create filenames for the output contracts. The column names are extracted from the filename entered by the user.

#### check\_columns(self) -\> Tuple[bool, Optional[List[str]]]

A method to check if all the columns included in the filename string are present in the Excel file. Returns a tuple with a Boolean indicating whether all columns are present, and an optional list of missing columns.

#### check\_filename(self) -\> str

A method to check if the entered filename string is valid. The method checks for illegal characters and missing columns. Returns an error message string if any issues are found.

#### on\_entry\_click(self, event: tk.Event) -\> None

A method to clear the default text in the filename entry widget when the user clicks on it.

#### on\_focusout(self, event: tk.Event) -\> None

A method to restore the default text in the filename entry widget when the user clicks off of it.

#### create\_widgets(self) -\> None

A method to create the GUI widgets, including labels, buttons, and menus, and grid them in the window. Binds the on\_entry\_click and on\_focusout methods to the filename entry widget.

**Usage**

To use the Grant\_Tool class, create an instance of the class. The instance will display the GUI. The user can then select a Word file, an Excel file, and an output folder. They can also enter a filename template for the generated documents. The user can then click on the "Generate" button to generate the documents. The generated documents will be saved in the selected output folder with filenames based on the filename template and the data from the Excel file.
