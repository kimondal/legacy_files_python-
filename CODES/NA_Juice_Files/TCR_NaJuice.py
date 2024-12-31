import os
from datetime import datetime, timedelta

import re
from fileinput import filename

from CODES.database.INV_Generator import NumberIncrementer
from CODES.NA_Juice_Files.MappingData_extraction import MappingDataProcessor
import pandas as pd
import numpy as np


class NAJuiceETL:
    def __init__(self, filepath):

        # Set the file path for the Excel file
        self.product_data_all = None
        self.filepath = filepath
        product_data_all = None
        self.filename=os.path.splitext(os.path.basename(filepath))[0]

        # Initialize attributes to store the data
        self.raw_data = None
        self.key_decoder = None
        self.consumer_data = None
        self.project_data = None
        self.product_data = None

        # Call the methods to fetch and process data during initialization
        self.RawData(sheet_name='Raw Data')
        self.Consumer(sheet_name='Consumer Info')
        self.KeyDecoder(sheet_name='Key_Decoder')
        self.Project(sheet_name='Additional data for NDB')
        self.Product()


    def RawData(self, sheet_name):
        """Fetch and process the raw data."""
        raw = pd.read_excel(self.filepath, sheet_name=sheet_name)

        def excel_to_datetime(excel_serial):
            base_date = datetime(1900, 1, 1)  # Excel's base date (1900-01-01)
            # Excel counts days starting from 1, so adjust for leap year bug and 1-based indexing
            return base_date + timedelta(days=excel_serial - 2)
        raw['Date'] = raw['Date'].apply(lambda x: excel_to_datetime(x) if isinstance(x, (int, float)) else x)
        def format_date(date_obj):
            # Check if it's a datetime object and then format
            if isinstance(date_obj, datetime):
                return date_obj.strftime('%m/%d/%Y')  # Format as mm/dd/yyyy
            return date_obj  # If it's not a datetime object, return as is

        # Apply formatting to the 'Date' column
        # raw['Date'] = raw['Date'].apply(format_date)

        file_path_na_juice = r"C:\Users\Z19661\PycharmProjects\LegacyData_\CODES\res\Legacy_Mapping.xlsx"

        processor = MappingDataProcessor(file_path_na_juice)
        raw_map = processor.na_juice().get('TCR_Raw', None)
        columns_df = pd.DataFrame(raw.columns, columns=['columns'])
        matching_cols = pd.merge(columns_df, raw_map[['Attribute Name.1']], left_on='columns', right_on='Attribute Name.1', how='inner')
        cols = matching_cols['columns'].tolist()
        raw = raw[cols]
        rename_dict = dict(zip(raw_map['Attribute Name.1'], raw_map['Attribute Name']))
        raw.rename(columns=rename_dict, inplace=True)

        self.raw_data = raw


    def KeyDecoder(self, sheet_name):
        # Step 1: Load the data from Excel
        df = pd.read_excel(self.filepath, sheet_name=sheet_name, engine='pyxlsb')

        # Drop rows where all elements are NaN and forward fill 'Worksheet' column
        df = df.dropna(how='all')
        df['Worksheet'] = df['Worksheet'].ffill()

        file_path_na_juice = r"C:\Users\Z19661\PycharmProjects\LegacyData_\CODES\res\Legacy_Mapping.xlsx"

        processor = MappingDataProcessor(file_path_na_juice).na_juice()

        rawtab = processor.get("TCR_Raw", None)
        consumertab = processor.get("TCR_Consumer", None)

        # Combine raw and consumer data into a single DataFrame
        combined_df = pd.concat([rawtab, consumertab], ignore_index=True)

        # Step 3: Merge the original DataFrame with the combined mapping data
        df = pd.merge(df, combined_df[['Attribute Name.1', 'Attribute Name']],
                      left_on='Column', right_on='Attribute Name.1', how='left')

        # Rename 'Attribute Name' to 'New Column'
        df['New Column'] = df['Attribute Name']

        # Drop unnecessary columns
        df.drop(columns=['Attribute Name.1', 'Attribute Name'], inplace=True)

        # Step 4: Get column names from raw and consumer data
        raw_data_columns = self.raw_data.columns.to_list()
        consumer_data_columns = self.consumer_data.columns.to_list()

        # Step 5: Initialize 'Worksheet' column as NaN, clear if needed
        df['Worksheet'] = np.nan

        # Step 6: Apply the logic for updating the 'Worksheet' column based on 'New Column' and 'Product Code'
        def update_worksheet(row):
            if row['Column'] == 'Product Code':
                return ['TCR_Raw']  # Set Worksheet to 'TCR_Product' for 'Product Code'

            worksheets = []
            if row['New Column'] in raw_data_columns:
                worksheets.append('TCR_Raw')
            if row['New Column'] in consumer_data_columns:
                worksheets.append('TCR_Consumer')

            return worksheets if worksheets else [np.nan]  # If no match, return NaN

        # Efficiently apply the function to the DataFrame to update the 'Worksheet' column
        df['Worksheet'] = df.apply(update_worksheet, axis=1)

        # Step 7: Flatten 'Worksheet' lists and remove duplicates based on 'Worksheet' + 'New Column'
        df = df.explode('Worksheet')  # Expand lists in 'Worksheet' to individual rows
        df = df.drop_duplicates(subset=['Worksheet', 'New Column'], keep='first')

        # Step 8: Drop rows where 'New Column' is NaN
        df.dropna(subset=['New Column'], inplace=True)
        # print(tabulate(df, headers='keys', tablefmt='fancy_grid', showindex=True))

        # Step 9: Custom ordering for the 'Worksheet' column
        worksheet_order = ['TCR_Raw', 'TCR_Consumer']
        worksheet_sort_map = {val: idx for idx, val in enumerate(worksheet_order)}

        # Create a new column to map 'Worksheet' to its custom order
        df['Worksheet_Order'] = df['Worksheet'].map(worksheet_sort_map).fillna(len(worksheet_sort_map))
        # Drop the 'Column' column
        df = df.drop(columns=['Column'])

        # Rename 'New Column' to 'Column'
        df = df.rename(columns={'New Column': 'Column'})

        # Re-arrange the columns to 'Worksheet', 'Column', 'Description'
        df = df[['Worksheet', 'Column', 'Description']]
        df.loc[len(df)] = ['TCR_Consumer', 'Country', 'Country of Test']
        self.key_decoder = df

    def Consumer(self, sheet_name):
        """Fetch and process the consumer data."""
        df = pd.read_excel(self.filepath, sheet_name=sheet_name)
        df_abn = pd.read_excel(self.filepath, sheet_name='Additional data for NDB')
        df.dropna(how='all', inplace=True)

        # Create a mapping from the additional data
        data_map = df_abn.set_index(df_abn.columns[0]).to_dict()[df_abn.columns[1]]
        country_of_test = data_map.get('Country of Test')

        # Rename and add columns in one step using assign
        new_df = (df.rename(columns={'Site': 'City'})
                    .assign(Country=country_of_test)
                    .drop(columns=['Income'], errors='ignore'))

        # Arrange the DataFrame in the specified order
        column_order = ['Cell', 'Respondent', 'City', 'Country', 'Gender', 'Age', 'Ethnicity', 'User Group']
        new_df = new_df[column_order].loc[:, column_order]
        self.consumer_data = new_df

    def Project(self, sheet_name):
        incrementer = NumberIncrementer()
        df = pd.read_excel(self.filepath, sheet_name=sheet_name, header=1)
        df = df[df['Field'] != 'Testing Cost, including currency name']
        df.columns = ['Field', 'Value']
        # df['Value'] = df['Value'].apply(lambda x: pd.to_datetime(x, unit='D', origin='julian') if isinstance(x, (int, float)) else x)
        date_fields = [
            'Start Date of Study Field (mm/dd/yyyy)',
            'End Date of Study Field (mm/dd/yyyy)'
        ]

        def excel_to_datetime(excel_serial):
            base_date = datetime(1900, 1, 1)  # Excel's base date (1900-01-01)
            # Excel counts days starting from 1, so adjust for leap year bug and 1-based indexing
            return base_date + timedelta(days=excel_serial - 2)

        # Convert 'Value' to datetime if 'Field' is in the date-related fields
        df['Value'] = df.apply(lambda row: excel_to_datetime(row['Value'])
        if row['Field'] in date_fields and isinstance(row['Value'], (int, float))
        else row['Value'], axis=1)
        #
        # # Now, format the datetime in 'mm/dd/yy' format
        # def format_date(date_obj):
        #     # Check if it's a datetime object and then format
        #     if isinstance(date_obj, datetime):
        #         return date_obj.strftime('%m/%d/%Y')  # Format as mm/dd/yy
        #     return date_obj  # If not a datetime, return as is
        #
        # # Apply the formatting to the 'Value' column
        # df['Value'] = df['Value'].apply(format_date)


        df.columns = ['Field', 'Value']
        new_rows = pd.DataFrame({
            'Field': ['Report Name', 'WorkFront Project Number', 'File Version Number','TCR Research Project Manager','Language'],
            'Value': ['', incrementer.get_inv_number(self.filename), '1','MON','EN']
        })
        new_rows['Field'] = new_rows['Field'].str.strip()
        new_rows['Value'] = new_rows['Value'].str.strip()
        self.project_data = pd.concat([df, new_rows], ignore_index=True)

    def Product(self, source=None):
        """Fetch and process the product data."""
        source = source if source is not None else self.key_decoder
        data = source[source['Column'] == 'Product Code'].dropna(subset=['Description'])
        self.product_data_all=data

        description_str = data['Description'].str.cat(sep=' ')
        trimmed_string = re.sub(r'^[^\d]*', '', description_str)
        pairs = trimmed_string.split(';')
        data = {key.strip(): value.strip() for pair in pairs if '=' in pair for key, value in [pair.split('=', 1)]}
        df = pd.DataFrame(list(data.items()), columns=['Product Code', 'Sample Description'])

        words_to_check = ['Control', 'current' ,'Prototype', 'PROTO','test',
                          "Coca-Cola",
                          "Diet Coke / Coca-Cola Light",
                          "Coca-Cola Zero Sugar",
                          "Sprite",
                          "Fanta",
                          "Dasani",
                          "Smartwater",
                          "Ciel",
                          "Bonaqua / BonAqua",
                          "Minute Maid",
                          "Simply",
                          "Del Valle",
                          "Innocent",
                          "AdeS",
                          "Gold Peak",
                          "Honest Tea",
                          "Peace Tea",
                          "Costa Coffee",
                          "Georgia Coffee",
                          "Monster Beverage Corporation (partnership)",
                          "Coca-Cola Energy",
                          "Powerade",
                          "Aquarius",
                          "Schweppes",
                          "Fresca",
                          "Barq's (Root Beer)",
                          "Mello Yello",
                          "Thums Up",
                          "Coca-Cola Citra",
                          "Inca Kola",
                          "Kuva",
                          "Royal Tru",
                          "Beverly",
                          "Fanta Melon Frosty",
                          "Mezzo Mix",
                          "Kuat",
                          "I LOHAS",
                          "Crystal",
                          "Kinley",
                          "ViO",
                          "Maaza",
                          "Del Valle Frut",
                          "Rani Float",
                          "Pulpy",
                          "Matte Leão",
                          "Cappy",
                          "Tropico",
                          "Ayataka",
                          "Fuze Tea",
                          "Leão Fuze",
                          "Chaywa",
                          "Marqués de Cáceres",
                          "Burn",
                          "Relentless",
                          "Glacau Vitamin Energy",
                          "I Power",
                          "Powerade ION4",
                          "Aquarius Water+",
                          "Appletiser",
                          "Bibo",
                          "Breeze",
                          "Ciel",
                          "Krest",
                          "Limca",
                          "Quatro"
                          ]
        pattern = '|'.join([re.escape(word) for word in words_to_check])  # Create a case-insensitive regex pattern

        # Set 'Type of Product' to 1 if any of the words exist in Sample Description, otherwise set it to 2
        df['Type of Product'] = df['Sample Description'].str.contains(pattern, case=False, na=False).astype(
            int)  # True becomes 1, False becomes 0
        df['Type of Product'] = df['Type of Product'].replace(0, 2)
        df.loc[df['Sample Description'].isnull(), 'Product Code'] = None


        df['Formula Spec Number'] = ''
        df['Sample Origin'] = None
        df['Product Category Tested'] = None
        df['Brand Tested'] = None
        df['First Flavor Tested'] = None
        df['Carbonated/Non-Carbonated'] = None
        df['Bottle/Can or Fountain Formula'] = None
        df['Serving Size'] = None
        df['Calorie'] = None
        df['pH'] = None
        df['BRIX'] = None
        df['Sugar'] = None
        df['Preservatives'] = None
        df['AVB'] = None
        df['CO2'] = None

        column_order = [
            'Type of Product', 'Formula Spec Number', 'Product Code', 'Sample Description', 'Sample Origin',
            'Product Category Tested', 'Brand Tested', 'First Flavor Tested', 'Carbonated/Non-Carbonated',
            'Bottle/Can or Fountain Formula', 'Serving Size', 'Calorie', 'pH', 'BRIX', 'Sugar', 'Preservatives',
            'AVB', 'CO2'
        ]

        # Reorder the DataFrame columns
        df = df[column_order]

        fn = os.path.splitext(os.path.basename(self.filepath))[0]
        df['Formula Spec Number'] = df['Product Code'].apply(
            lambda x: str(NumberIncrementer.get_product_code_suffix(self, fn)) +
                      (str(x).zfill(3)) + '-001' if pd.notnull(x) else None)
        # print(tabulate(df, headers='keys', tablefmt='fancy_grid', showindex=True))
        df.loc[df['Type of Product'] != 1, 'Formula Spec Number'] = None

        self.product_data = df

#     def convert_to_excel(self, output_filepath):
#         """Convert all processed data to an Excel file."""
#         with pd.ExcelWriter(output_filepath, engine='openpyxl') as writer:
#             if self.raw_data is not None:
#                 self.raw_data.to_excel(writer, sheet_name='TCR_Raw', index=False)
#             if self.consumer_data is not None:
#                 self.consumer_data.to_excel(writer, sheet_name='TCR_Consumer', index=False)
#             if self.key_decoder is not None:
#                 self.key_decoder.to_excel(writer, sheet_name='Key_Decoder', index=False)
#             if self.project_data is not None:
#                 self.project_data.to_excel(writer, sheet_name='TCR_Project', index=False)
#             if self.product_data is not None:
#                 self.product_data.to_excel(writer, sheet_name='TCR_Product', index=False)
#
# # # Example Usage:
#
# Define the file path to the Excel file
na_juice_filepath = r"C:\Legacy Files\all_files\For TCS\NA Juice\Minute Maid Refreshment CLT - White Lemonade Phase 2_Full Data File with Key.xlsb"

# Create an instance of NAJuiceETL, this will automatically load all data
NAJuiceETL(filepath=na_juice_filepath)
#
# # Convert the processed data to an output Excel file
# output_filepath = "../res/data.xlsx"
# na_juice_etl.convert_to_excel(output_filepath)
#
# print("ETL process completed successfully and saved to Excel.")
