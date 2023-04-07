# Grant Tool Readme

This script allows you to generate multiple personalized contract files from a given Word document template and an Excel file containing the contract information.

## Prerequisites

Microsoft Word needs to be installed.

To run this script, you need to have Python installed on your system. You also need to install the following libraries:

    - 'tkinter'
    - 'pandas'
    - 'openpyxl'
    - 'mailmerge'
    - 'docx2pdf'

You can install these libraries using pip, by running the following command:

pip install tkinter pandas openpyxl mailmerge docx2pdf

## How to use

    1. Run the script.
    2. Select a Word file to use as the template for the contracts.
    3. Select an Excel file that contains the contract information. The file should be in '.xlsx' format.
    4. Select a folder to save the output files.
    5. Fill in the 'Enter filename...' field with the desired filename format for the generated contracts. You can use column names from the Excel file by enclosing them in curly braces, e.g. '{Name}'. Note that only exact matches will work.
    6. Click the 'Create contracts' button to generate the contracts.

## User Interface

The script will launch a user interface with the following components:

    - A label and button to select a Word file.
    - A label and button to select an Excel file.
    - A label and button to select an output folder.
    - An entry field to enter the desired filename format for the generated contracts.
    - A 'Help' button to display the available fields that can be used in the filename format.
    - A 'Create contracts' button to generate the contracts.
    - A scrolling text box that displays the status of the contract generation process.

Note that the script can take some time to generate the contracts, especially if the Excel file is large.
Important Notes

    The Excel file should have a header row that contains the column names.
    The script only supports '.xlsx' files.
    The 'docx2pdf' library is used to convert the generated Word documents to PDF files. If you do not have this library installed, the PDF conversion step will be skipped.
    
    
See more: [In depth code documentation](docs/code/DOCUMENTATION.md)
