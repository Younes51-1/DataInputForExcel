import PySimpleGUI as sg
import pandas as pd
import os
from table import Table
class ExcelEditorWindow:
    def __init__(self):
        self.file_path = None
        self.df = None
        self.selected_row = None
        self.lab_number = None
        # Create layout
        main_layout = [
            [sg.Text("Choose an Excel file:")],
            [sg.InputText(key="FILE_PATH"), 
            sg.FileBrowse(initial_folder=os.getcwd(), file_types=[("Excel Files", "*.xlsx")])],
            [sg.Text("Which lab:"), sg.InputText(key="LAB", size=(10, 10))],
            [sg.Exit(), sg.Push(), sg.Button("Submit")]
        ]
        
        self.table = None
        # sg.Table(
        #     values=[],
        #     headings=[],
        #     display_row_numbers=True,
        #     max_col_width=35,
        #     auto_size_columns=True,
        #     display_row_numbers=False,
        #     justification='centre',
        #     num_rows=10,
        #     key='-TABLE-',
        #     row_height=35,
        #     tooltip="Grades Table",
        #     enable_events=True,
        #     expand_x=True,
        #     expand_y=True,
        #     enable_click_events=True
        # )
        # Create window
        self.window = sg.Window("Excel Input", main_layout)

    def run(self):
        while True:
            event, values = self.window.read()

            if event in (sg.WIN_CLOSED, 'Exit'):
                break
            elif event == "Submit":
                self.load_file(values["FILE_PATH"], values["LAB"])
            elif event == 'Search_key_update_table':
                self.update_table(values["Search_key"])
            elif event == "-TABLE-_get_selected_row":
                self.get_selected_row(values["Search_key"])
                
        self.window.close()

    def load_file(self, path, lab_number):
        if path and lab_number.isdigit():
            self.file_path = path
            self.lab_number = int(lab_number)
            self.create_table(self.lab_number)
        elif not lab_number.isdigit() and not path:
            sg.popup_error("Please enter a valid lab number and file path.")
        elif not path:
            sg.popup_error("Please enter a valid file path.")
        elif not lab_number.isdigit():
            sg.popup_error("Please enter a valid lab number.")
        else:
            sg.popup_error("Unknown error.")
            

    def create_table(self, lab_number):
        self.table = Table(self.file_path, "MATRICULE", lab_number)
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
    # sg.Text("Grade"), sg.InputText(key="Grade_key")
        layout = [
            [table_window],
            [sg.Text("Matricule"), sg.InputText(key="Search_key", enable_events=True)],
            [sg.Push(), sg.Button("Save")]
        ]
        self.window.close()
        self.window = sg.Window("Grades Table", layout, finalize=True)
        self.window['Save'].bind("<Return>", "_Enter")
        self.window['Search_key'].bind('<Key>', '_update_table')
        self.window['-TABLE-'].bind('<Double-Button-1>', '_get_selected_row')
    
    def update_table(self, search_key=''):
        if self.table.df_exists():
            data = self.table.filter_df(search_key)
            self.window['-TABLE-'].update(values=data)
    
    def get_selected_row(self, search_key=''):
        if self.table.df_exists():
            selected_row_index = self.window['-TABLE-'].SelectedRows[0]
            selected_row_data = self.table.filtered_df.iloc[selected_row_index]
            formatted_data = "\n".join([f"{col}: {value}" for col, value in selected_row_data.items()])
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

    def add_grade_to_row(self, row_index, grade, search_key):
        if self.table.df_exists():
            if not self.table.filtered_df.empty:
                self.table.filtered_df = self.table.df
            # Modify the code to add the grade to the selected row in the DataFrame
            # For example, you can update a specific cell in the DataFrame
            last_column_index = self.table.df.shape[1] - 1
            self.table.df.iloc[self.table.df.MATRICULE == row_index, last_column_index] = grade
            # Update the table to reflect the changes
            self.update_table(search_key)
            

if __name__ == "__main__":
    editor = ExcelEditorWindow()
    editor.run()