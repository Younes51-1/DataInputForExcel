import pandas as pd
import PySimpleGUI as sg
import numpy as np

class Table:
    def __init__(self, path, search_column, lab_number):
        self.path = path
        self.search_column = search_column
        self.lab_number = int(lab_number)
        self.df =  pd.read_excel(self.path, engine='openpyxl')   
        columns_to_keep = list(set([0, 1, 2, self.lab_number]))
        self.df = self.df.iloc[:, columns_to_keep]
        self.filtered_df = self.df
    
    
    def df_exists(self):
        return not self.df.empty
    
    def headings(self):
        return self.df.columns.to_numpy().tolist()
    
    def data(self):
        return self.df.to_numpy().tolist()
    
    def filter_df(self, search_key):
        if search_key == '':
            self.filtered_df = self.df
        self.filtered_df = self.df[self.df.apply(lambda row: any(str(cell).find(str(search_key)) != -1 for cell in row), axis=1)]
        return self.filtered_df.to_numpy().tolist()