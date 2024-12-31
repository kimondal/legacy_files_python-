import os
import pandas as pd

class NumberIncrementer:
    def get_inv_number(self, filename):
        excel_file_path = r'C:\Users\Z19661\PycharmProjects\LegacyData_\CODES\res\Legacy_Dummy_INV.xlsx'
        df = pd.read_excel(excel_file_path, engine='openpyxl')
        filename_with_extension = os.path.basename(filename)
        matches = df[df['File Name'].str.contains(filename_with_extension, case=False, na=False)]
        if not matches.empty:
            inv_number = str(matches.iloc[0]['INV Number'])
            return inv_number
        else:
            return " "

    def get_product_code_suffix(self, filename):
        excel_file_path = r'C:\Users\Z19661\PycharmProjects\LegacyData_\CODES\res\Legacy_Dummy_INV.xlsx'
        df = pd.read_excel(excel_file_path, engine='openpyxl')
        filename_with_extension = os.path.basename(filename)
        matches = df[df['File Name'].str.contains(filename_with_extension, case=False, na=False)]
        if not matches.empty:
            inv_number = str(matches.iloc[0]['Product_Code'])
            return inv_number
        else:
            return "900x"


