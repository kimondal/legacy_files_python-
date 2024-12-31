import pandas as pd
import os
from CODES.NA_Juice_Files.MappingData_extraction import MappingDataProcessor


class EMEA_NBD_ETL:
    file_path_na_juice = r"../res/EMEA_Mapping.xlsx"

    mapping_data = MappingDataProcessor(file_path_na_juice).EMEA_withNDB(file_path_na_juice)

    def __init__(self, filepath):
        """Initialize the EMEA_NBD_ETL class with the provided file path."""
        self.filepath = filepath


        # Initialize attributes to store the data
        self.raw_data = None
        self.key_decoder = None
        self.consumer_data = None
        self.project_data = None
        self.product_data = None

        # Process the data when the class is instantiated
        self.process_data()

    def process_data(self):
        """Process all data required for the ETL."""
        self.RawData(sheet_name="Raw Data")
        self.Consumer(sheet_name="Consumer Info", product_tab="Additional data for NDB")
        self.Project(sheet_name="Additional data for NDB")
        self.KeyDecoder(sheet_name="Key_Decoder")
        self.Product(source=self.key_decoder)

    def RawData(self, sheet_name):
        """Fetch and process the raw data."""
        raw = pd.read_excel(self.filepath, sheet_name=sheet_name)

        raw_map = self.mapping_data.get('TCR_Raw', None)

        columns_df = pd.DataFrame(raw.columns, columns=['columns'])
        matching_cols = pd.merge(columns_df, raw_map[['Attribute Name.2']], left_on='columns',
                                 right_on='Attribute Name.2', how='inner')

        cols = matching_cols['columns'].tolist()

        raw = raw[cols]
        rename_dict = dict(zip(raw_map['Attribute Name.2'], raw_map['Attribute Name']))
        raw.rename(columns=rename_dict, inplace=True)
        self.raw_data = raw


    def Consumer(self, sheet_name, product_tab):
        """Fetch and process the consumer data."""
        consumer = pd.read_excel(self.filepath, sheet_name=sheet_name)
        consumer_mapping = self.mapping_data.get('TCR_Consumer', None)
        df_abn = pd.read_excel(self.filepath, sheet_name=product_tab)
        df_abn = df_abn.iloc[:, :2].dropna()
        data_map = df_abn.set_index(df_abn.columns[0]).to_dict()[df_abn.columns[1]]
        country_of_test = data_map.get('Country of Test', 'Unknown')
        consumer['Country of Test'] = country_of_test
        columns_df = pd.DataFrame(consumer.columns, columns=['columns'])
        matching_cols = pd.merge(columns_df, consumer_mapping[['Attribute Name.2']], left_on='columns',
                                 right_on='Attribute Name.2', how='inner')
        cols = matching_cols['columns']
        consumer = consumer[cols]
        rename_dict = dict(zip(consumer_mapping['Attribute Name.2'], consumer_mapping['Attribute Name']))
        consumer.rename(columns=rename_dict, inplace=True)
        self.consumer_data = consumer

    def Product(self, source=None):
        """Fetch and process the product data."""
        source = source if source is not None else self.key_decoder
        data = source[source['Column'] == 'Product Code'].dropna(subset=['Description'])
        self.product_data = data

    def KeyDecoder(self, sheet_name):
        """Fetch and process the key decoder data."""
        df = pd.read_excel(self.filepath, sheet_name=sheet_name, header=0)
        self.key_decoder = df

    def Project(self, sheet_name):
        """Fetch and process the project data."""
        df = pd.read_excel(self.filepath, sheet_name=sheet_name).iloc[:, :2]
        self.project_data = df

    def convert_to_excel(self, output_filepath):
        """Convert all processed data to an Excel file."""
        with pd.ExcelWriter(output_filepath, engine='openpyxl') as writer:
            if self.raw_data is not None:
                self.raw_data.to_excel(writer, sheet_name='TCR_Raw', index=False)
            if self.consumer_data is not None:
                self.consumer_data.to_excel(writer, sheet_name='TCR_Consumer', index=False)
            if self.key_decoder is not None:
                self.key_decoder.to_excel(writer, sheet_name='TCR_Key_Decoder', index=False)
            if self.project_data is not None:
                self.project_data.to_excel(writer, sheet_name='TCR_Project', index=False)
            if self.product_data is not None:
                self.product_data.to_excel(writer, sheet_name='TCR_Product', index=False)

    @staticmethod
    def find_xlsx_files_to_dataframe(directory):
        """Find all Excel files in a given directory and return their paths in a DataFrame."""
        file_data = []
        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.lower().endswith('.xlsx'):
                    file_path = os.path.join(root, file)
                    file_data.append((file, file_path))
        df = pd.DataFrame(file_data, columns=['File Name', 'File Path'])
        return df



