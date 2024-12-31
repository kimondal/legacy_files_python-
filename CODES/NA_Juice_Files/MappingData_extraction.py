import pandas as pd
import os

class MappingDataProcessor:
    def __init__(self, file_path=r'C:\Users\Z19661\PycharmProjects\LegacyData_\CODES\res\Legacy_Mapping.xlsx'):
        """Initialize the MappingDataProcessor with a file path."""
        self.file_path = file_path
        self.df = None
        if file_path:
            self.load_data(file_path)

    def load_data(self, file_path):
        """Load Excel data from the given file path."""
        try:
            # Load the Excel file
            self.df = pd.read_excel(file_path, header=1)  # Default header row is 1
            absolute_path = os.path.abspath(file_path)
        except Exception as e:
            self.df = None

    def na_juice(self):
        """Process the data for NA Juice."""
        if self.df is not None:
            try:
                # Filter for relevant columns and drop rows where 'Attribute Name.1' is missing
                nnewdf = self.df[['Tab Name', 'Attribute Name', 'Tab Name.1', 'Attribute Name.1']]
                nnewdf = nnewdf.dropna(subset=['Attribute Name.1'])

                # Group by 'Tab Name' to get separate DataFrames for each tab
                grouped = nnewdf.groupby('Tab Name')
                grouped_dfs = {tab_name: group_df for tab_name, group_df in grouped}

                return grouped_dfs
            except Exception as e:
                return None
        else:
            return None

    def EMEA(self):
        """Process the data for EMEA."""
        if self.df is not None:
            try:
                # Filter for relevant columns and drop rows where 'Attribute Name.2' is missing
                nnewdf = self.df[['Tab Name', 'Attribute Name', 'Tab Name.2', 'Attribute Name.2']]
                nnewdf = nnewdf.dropna(subset=['Attribute Name.2'])

                # Group by 'Tab Name' to get separate DataFrames for each tab
                grouped = nnewdf.groupby('Tab Name')
                grouped_dfs = {tab_name: group_df for tab_name, group_df in grouped}

                return grouped_dfs
            except Exception as e:
                return None
        else:
            return None

    def EMEA_withNDB(self, file_path):
        """Process the data for EMEA with NDB mapping."""
        try:
            # Load a different Excel file with a different header row (header=2)
            df = pd.read_excel(file_path, header=2)

            # Filter for relevant columns and drop rows where 'Attribute Name.2' is missing
            nnewdf = df[['Tab Name', 'Attribute Name', 'Tab Name.2', 'Attribute Name.2']]
            nnewdf = nnewdf.dropna(subset=['Attribute Name.2'])

            # Group by 'Tab Name' to get separate DataFrames for each tab
            grouped = nnewdf.groupby('Tab Name')
            grouped_dfs = {tab_name: group_df for tab_name, group_df in grouped}

            return grouped_dfs
        except Exception as e:
            return None

