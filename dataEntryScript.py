import pandas as pd

# TODO Replace 'your_excel_file.xlsx' with the path to your Excel file
excel_file_path = 'file.xlsx'
# TODO Replace 'Sheet1' with the name of the sheet you want to go to
sheet_name = 'Sheet1'
# TODO Replace with the actual column name
column_name = 'xxxxx' 

def openExcelFile(file_path, sheet_name):
    return pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')

def lookupStudent(df, search_value):
    return df[df.MATRICULE == search_value]

def userInput(userPrompt):
    while True:
        try:
            user_input = int(userPrompt)
            return user_input
        except ValueError:
            print("Invalid input. Please enter an integer.")

def confirmInput():
    while True:
        confirm = input("Is this the correct? (y/n) ")
        if confirm == "y":
            return True
        elif confirm == "n":
            return False
        else:
            print("Invalid input. Please try again.")
            
def confirmStudent(df_student):
    print(f"Found '{df_student.MATRICULE}' in the following cells:")
    print(df_student)
    return confirmInput()
    
def inputGrade(df, search_value):
    while True:
        grade = userInput("Grade: ")
        print(f"Inputed grade: {grade}")
        if confirmInput():
            df.loc[df.MATRICULE == search_value, column_name] = grade
            return
    
def savingExcelFile(df, file_path, sheet_name):
    df.to_excel(file_path, sheet_name=sheet_name, index=False)


def main_loop():
    try:
        # Read the Excel file into a Pandas DataFrame
        df = openExcelFile(excel_file_path, sheet_name)
        while True:
            search_value = int(MatriculeInput("Matricule: "));
            if search_value == -1:
                break
            # Search for the student in the DataFrame
            result = lookupStudent(df, search_value)

            if not result.empty:
                #confirming students then adding grade
                if confirmStudent(result):
                    inputGrade(df, search_value)
            else:
                print(f"'{search_value}' not found in '{sheet_name}'")
            
        # Write the DataFrame back to the Excel file
        savingExcelFile(df, excel_file_path, sheet_name)
        print(f"Successfully saved File {excel_file_path}")
        
    except FileNotFoundError:
        print(f"File '{excel_file_path}' not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
    
    
if __name__ == '__main__':
    main_loop()