# DataInputForExcel

## Purpose
This script is specifically designed for the LOG1810 course

## Issue
When using Pandas to save data to an Excel file, there can be formatting issues that affect the appearance of the data.

## Instructions

1. **Install Dependencies:**
    To ensure the required dependencies are installed, run the following command in your terminal or command prompt:

    ```bash
    pip install -r requirements.txt
    ```

    The `requirements.txt` file includes project-specific dependencies such as Pandas, PySimpleGUI, and NumPy.

2. **Convert to Executable:**
    Use PyInstaller to convert the script to an executable. Run the following command:

    ```bash
    pyinstaller --onefile --noconsole window.py
    ```

    This will create a standalone executable file without opening a console window.

3. **Run the Executable:**
    Locate the generated executable file (in the 'dist' directory if using the `--onefile` option) and run it to execute the script.

## Notes
- This script is customized for the LOG1810 course and may not be suitable for other use cases without modification.
- Ensure that you have Python and Pip installed on your system before running the script.
- If issues persist, consider checking for updates to Pandas or exploring alternative solutions.

Feel free to modify the script to suit your specific needs or the requirements of your course.
