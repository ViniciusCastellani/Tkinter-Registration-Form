# Tkinter Data Entry Form

This Python script uses Tkinter to create a user registration form that collects user data and stores it in an Excel file (.xlsx). Additionally, it allows users to view the entered data in a table within the graphical interface.

## Description

The code comprises the following key components:

### Registration Interface

- Captures information such as first name, last name, email, password, cellphone number, and age.
- Utilizes a `Spinbox` to capture the age input.
- Includes a checkbox to accept terms and conditions.

### Functionalities

- Shows/hides the password when clicking the "show password" option.
- Performs validations on form fields.
- Saves user data to an Excel file (`data.xlsx`).
- Loads data from the Excel file into the interface for visualization.

### Data Visualization

- Utilizes the `Treeview` component to display data in a table format.
- Features a scrollbar to view data beyond the widget's size.

### Load Data Functionality

- Checks if the Excel file already exists; if not, creates a new file and adds column titles.
- Loads data from the Excel file into the visualization interface.

### "Register" Button

- Upon being pressed, executes the `enterData` function that checks filled fields and saves data to the Excel file.
- Updates the data view in the table.

### Other Details

- Organizes interface elements using `Frame` to separate sections of the form.

## Execution

To run this code:
1. Ensure the `openpyxl` library is installed. If not, install it using `pip install openpyxl`.
2. Execute the code in a Python environment to interact with the interface.

## Usage

- Fill in the registration form fields.
- Check the "I accept terms and conditions" checkbox.
- Click the "Register" button to save the data.
- View the entered data in the table below the form.

