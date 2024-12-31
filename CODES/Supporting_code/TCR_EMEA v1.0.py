import pandas as pd
import re
from CODES.NA_Juice_Files.MappingData_extraction import MappingDataProcessor

# Load the mapping data


class EMEA_ETL:
    file_path_na_juice = "../res/Legacy_Mapping.xlsx"
    mapping_data = MappingDataProcessor(file_path_na_juice).EMEA()

    def __init__(self, filepath):
        """Initialize the EMEA_ETL class with the provided file path."""
        self.filepath = filepath

        # Initialize attributes to store the data
        self.raw_data = None
        self.key_decoder = None
        self.consumer_data = None
        self.project_data = None
        self.product_data = None

        # Call methods to process the data
        self.process_data()

    def process_data(self):
        """Process all data required for the ETL."""
        self.RawData(sheet_name="Raw Data")
        self.Consumer(sheet_name="Raw Data")
        self.Project(sheet_name="Additional MSE data")
        self.KeyDecoder(sheet_name="Decoder")
        self.Product(source=self.key_decoder)

    def RawData(self, sheet_name):
        """Fetch and process the raw data."""
        raw = pd.read_excel(self.filepath, sheet_name=sheet_name)
        raw_map = self.mapping_data.get('TCR_Raw', None)
        columns_df = pd.DataFrame(raw.columns, columns=['columns'])
        matching_cols = pd.merge(columns_df, raw_map[['Attribute Name.2']], left_on='columns', right_on='Attribute Name.2', how='inner')
        cols = matching_cols['columns'].tolist()
        raw = raw[cols]
        rename_dict = dict(zip(raw_map['Attribute Name.2'], raw_map['Attribute Name']))
        raw.rename(columns=rename_dict, inplace=True)
        self.raw_data = raw

    def Consumer(self, sheet_name):
        """Fetch and process the consumer data."""
        consumer = pd.read_excel(self.filepath, sheet_name=sheet_name)
        consumer_mapping = self.mapping_data.get('TCR_Consumer', None)
        df_abn = pd.read_excel(self.filepath, sheet_name='Additional MSE data')
        df_abn = df_abn.iloc[:, :2].dropna()
        data_map = df_abn.set_index(df_abn.columns[0]).to_dict()[df_abn.columns[1]]
        country_of_test = data_map.get('Country of Test', 'Unknown')
        consumer['Country of Test'] = country_of_test
        columns_df = pd.DataFrame(consumer.columns, columns=['columns'])
        matching_cols = pd.merge(columns_df, consumer_mapping[['Attribute Name.2']], left_on='columns', right_on='Attribute Name.2', how='inner')
        cols = matching_cols['columns']
        consumer = consumer[cols]
        rename_dict = dict(zip(consumer_mapping['Attribute Name.2'], consumer_mapping['Attribute Name']))
        consumer.rename(columns=rename_dict, inplace=True)
        self.consumer_data = consumer

    def Product(self, source=None):
        """Fetch and process the product data."""
        source = source if source is not None else self.key_decoder
        data = source[source['Variable for Automation'] == 'Product ID'].dropna(subset=['Scale'])
        description_str = data['Scale'].str.cat(sep=' ')
        trimmed_string = re.sub(r'^[^\d]*', '', description_str)
        pairs = trimmed_string.split(';')
        data_dict = {key.strip(): value.strip() for pair in pairs if '=' in pair for key, value in [pair.split('=', 1)]}
        df = pd.DataFrame(list(data_dict.items()), columns=['Product Code', 'Sample Description'])
        self.product_data = df

    def KeyDecoder(self, sheet_name):
        """Fetch and process the key decoder data."""
        df = pd.read_excel(self.filepath, sheet_name=sheet_name, header=0)
        self.key_decoder = df

    def Project(self, sheet_name):
        """Fetch and process the project data."""
        df = pd.read_excel(self.filepath, sheet_name=sheet_name).iloc[:, :2]
        self.project_data = df


#     def convert_to_excel(self, output_filepath):
#         """Convert all processed data to an Excel file."""
#         with pd.ExcelWriter(output_filepath, engine='openpyxl') as writer:
#             if self.raw_data is not None:
#                 self.raw_data.to_excel(writer, sheet_name='TCR_Raw', index=False)
#             if self.consumer_data is not None:
#                 self.consumer_data.to_excel(writer, sheet_name='TCR_Consumer', index=False)
#             if self.key_decoder is not None:
#                 self.key_decoder.to_excel(writer, sheet_name='TCR_Key_Decoder', index=False)
#             if self.project_data is not None:
#                 self.project_data.to_excel(writer, sheet_name='TCR_Project', index=False)
#             if self.product_data is not None:
#                 self.product_data.to_excel(writer, sheet_name='TCR_Product', index=False)
#
#     @staticmethod
#     def find_xlsx_files_to_dataframe(directory):
#         """Find all .xlsx files in a directory and return them as a DataFrame."""
#         file_data = []  # List to store tuples of (file_name, file_path)
#         for root, dirs, files in os.walk(directory):
#             for file in files:
#                 file_path = os.path.join(root, file)
#                 file_data.append((file, file_path))  # Append as a tuple (file_name, file_path)
#
#         # Create a DataFrame from the list of tuples
#
#         df = pd.DataFrame(file_data, columns=['File Name', 'File Path'])
#         return df
#
#     @classmethod
#     def process_directory(cls, directory):
#         """Process all .xlsx files in the directory."""
#         file_data_frame = cls.find_xlsx_files_to_dataframe(directory)
#         emea_filepaths = file_data_frame['File Path'].tolist()
#
#         print(emea_filepaths)
#         results = []
#         for filepath in emea_filepaths:
#             try:
#                 # Process each file
#                 emea_etl = cls(filepath)
#                 output_filepath = f"Output/TCR_emea/{os.path.basename(filepath)}.xlsx"
#                 emea_etl.convert_to_excel(output_filepath)
#                 results.append({
#                     "filename": filepath,
#                     "status": "OK"
#                 })
#             except Exception as e:
#
#                 results.append({
#                     "filename": filepath,
#                     "status": f"Not OK - {str(e)}"
#                 })
#
#         status_df = pd.DataFrame(results)
#         status_df.to_csv('res/output_emea_status.csv', index=False)
#         print("Processing completed. Status saved to output_emea_status.csv")
#
# # Example Usage:
#
# if __name__ == "__main__":
#     directory_path = r"C:\Legacy Files\all_files\For TCS\EMEA\Brussels TCR (Swetlana)"
#     EMEA_ETL.process_directory(directory_path)
