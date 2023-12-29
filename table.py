"""pd used to handle dataframes and excel files."""
import pandas as pd

class Table:
    """Class create a df and handle logic related to df manipulation."""
    def __init__(self, path, search_column, lab_number):
        self.path = path
        self.search_column = search_column
        self.df = pd.read_excel(self.path, engine='openpyxl')
        self.lab_name = f'TP{int(lab_number)}'
        self.columns_to_keep = ['MATRICULE', 'Nom de famille', 'Prénom',  self.lab_name]
        if  self.lab_name not in self.df.columns:
            self.df[ self.lab_name] = 0
        self.df = self.df[self.columns_to_keep]
        self.filtered_df = self.df


    def df_exists(self):
        """Returns True if df exists, False otherwise."""
        return not self.df.empty

    def headings(self):
        """Returns the df header as a list."""
        return self.df.columns.to_numpy().tolist()

    def data(self):
        """Returns the df data as a list."""
        return self.df.to_numpy().tolist()

    def filter_df(self, search_key):
        """Returns a filtered df based on the search key."""
        if search_key == '':
            self.filtered_df = self.df
        search_results = self.df.apply(self.search_func(search_key), axis=1)
        self.filtered_df = self.df[search_results]
        return self.filtered_df.to_numpy().tolist()

    def search_func(self, search_key):
        """Returns a function that can be used to search a df."""
        return lambda row: any(str(cell).find(str(search_key)) != -1 for cell in row)

    def update_df(self, row_key, grade):
        """Updates the df with the new grade."""
        if not self.filtered_df.empty:
            self.filtered_df = self.df
        last_column_index = self.df.shape[1] - 1
        self.df.iloc[self.df[self.search_column] == row_key, last_column_index] = grade

    def save_changes(self):
        """Save changes into the excel file."""
        old_df = pd.read_excel(self.path, engine='openpyxl')
        self.df = pd.merge(old_df, self.df, on=self.columns_to_keep.remove(self.lab_name))
        self.df.to_excel(self.path, index=False)
