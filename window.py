"""os used to get current working directory to simplify finding xlsx.
    sg used for the GUI.
    table used to handle the table logic."""
import os
import PySimpleGUI as sg
from table import Table

class ExcelEditorWindow:
    """Class creates a window to edit excel files."""
    def __init__(self):
        self.df = None
        main_layout = [
            [sg.Text("Choose an Excel file:")],
            [sg.InputText(key="FILE_PATH"),
            sg.FileBrowse(initial_folder=os.getcwd(), file_types=[("Excel Files", "*.xlsx")])],
            [sg.Text("Which lab:"), sg.InputText(key="LAB", size=(10, 10))],
            [sg.Exit(), sg.Push(), sg.Button("Submit")]
        ]
        self.table = None
        self.window = sg.Window("Excel Input", main_layout)

    def run(self):
        """Main function to run the window."""
        while True:
            event, values = self.window.read()

            if event in (sg.WIN_CLOSED, 'Exit', 'Save'):
                break
            elif event == "Submit":
                self.load_file(values["FILE_PATH"], values["LAB"])
            elif event == 'Search_key_update_table':
                self.update_table(values["Search_key"])
            elif event == "-TABLE-_get_selected_row":
                self.get_selected_row(values['-TABLE-'][0], values["Search_key"])

        self.window.close()

    def load_file(self, path, lab_number):
        """Verifie if the file path and lab number are valid 
        if yes create table. else popup error message."""
        if path and lab_number.isdigit():
            self.create_table(path, lab_number)
        elif not lab_number.isdigit() and not path:
            sg.popup_error("Please enter a valid lab number and file path.")
        elif not path:
            sg.popup_error("Please enter a valid file path.")
        elif not lab_number.isdigit():
            sg.popup_error("Please enter a valid lab number.")
        else:
            sg.popup_error("Unknown error.")

    def create_table(self, path, lab_number):
        """Create table and update window."""
        self.table = Table(path, "MATRICULE", lab_number)
        table_window = sg.Table(
            values=self.table.data(),
            headings=self.table.headings(),
            max_col_width=35,
            auto_size_columns=True,
            display_row_numbers=False,
            justification='centre',
            num_rows=10,
            key='-TABLE-',
            row_height=35,
            tooltip="Grades Table",
            enable_events=True,
            expand_x=True,
            expand_y=True,
            enable_click_events=True
        )
        layout = [
            [table_window],
            [sg.Text("Matricule"), sg.InputText(key="Search_key", enable_events=True)],
            [sg.Push(), sg.Button("Save")]
        ]
        self.window.close()
        self.window = sg.Window("Grades Table", layout, finalize=True)
        self.window['Search_key'].bind('<Key>', '_update_table')
        self.window['-TABLE-'].bind('<Double-Button-1>', '_get_selected_row')

    def update_table(self, search_key=''):
        """Update the table with the new search key."""
        if self.table.df_exists():
            data = self.table.filter_df(search_key)
            self.window['-TABLE-'].update(values=data)

    def get_selected_row(self, selected_row_index, search_key=''):
        """Get the selected row and open a popup to enter a grade."""
        if self.table.df_exists():
            selected_row_data = self.table.filtered_df.iloc[selected_row_index]
            list_data = [f"{col}: {value}" for col, value in selected_row_data.items()]
            formatted_data = "\n".join(list_data)
            layout = [
                [sg.Multiline(formatted_data, size=(40, 5), key='selected_row', disabled=True)],
                [sg.Text("Enter Grade:"), sg.InputText(key="grade")],
                [sg.Button("OK"), sg.Button("Cancel")]
            ]

            popup_window = sg.Window("Enter Grade", layout)

            while True:
                popup_event, popup_values = popup_window.read()

                if popup_event in (sg.WIN_CLOSED, 'Cancel'):
                    break
                elif popup_event == 'OK':
                    grade = popup_values['grade']
                    self.add_grade_to_row(selected_row_data[0], grade, search_key)
                    sg.popup_ok(f"Entered Grade: {grade}")
                    break

            popup_window.close()

    def add_grade_to_row(self, row_key, grade, search_key):
        """Add a grade to the selected row."""
        if self.table.df_exists():
            self.table.update_df(row_key, grade)
            self.update_table(search_key)

    def save_changes(self):
        """Save changes to the excel file."""
        if self.table.df_exists():
            self.table.save_changes()

if __name__ == "__main__":
    editor = ExcelEditorWindow()
    editor.run()
    editor.save_changes()
