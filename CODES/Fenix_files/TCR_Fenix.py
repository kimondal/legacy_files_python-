import os
import re

import pandas as pd
from tabulate import tabulate

from CODES.Fenix_files.keydecoder_deduceFenix import AttributeDataProcessorFenix
from CODES.Fenix_files.keydecodermaking import ExcelProcessor
from CODES.Fenix_files.project_tab_determination import DataFrameEnhancerFenix
from CODES.database.INV_Generator import NumberIncrementer
from CODES.database.city_determination import get_most_frequent_country_from_cities
from CODES.Fenix_files.kd import ExcelProcessor_v2


class Fenix_ETL:

    def __init__(self, filepath):
        self.filepath = filepath

        # Initialize attributes to store the data
        self.raw_data = None
        self.key_decoder = None
        self.consumer_data = None
        self.project_data = None
        self.product_data = None
        self.null_columns_list = []
        self.lookup_data = None
        self.country_of_test = None
        self.start_date = None
        self.end_date = None

        # Call methods to process the data
        self.process_data()

    def check_null_columns(self, df):
        null_columns = df.columns[df.isnull().any()].tolist()
        self.null_columns_list.extend(null_columns)

    def process_data(self):
        sheet_names = pd.ExcelFile(self.filepath).sheet_names
        filtered_sheets = [sheet for sheet in sheet_names if sheet.lower() not in ['datamap', 'labels']]

        if filtered_sheets:
            self.KeyDecoder()
            self.RawData(sheet_name=filtered_sheets[0])
            self.Consumer(sheet_name=filtered_sheets[0])
            self.Project()
            self.Product(self.filepath)
            self.reversal()
        else:
            print("No valid sheets found.")

    def RawData(self, sheet_name):
        raw = pd.read_excel(self.filepath, sheet_name=sheet_name, dtype=str)
        # print(tabulate(raw.head(10), headers='keys', tablefmt='fancy_grid', showindex=True))
        processor = ExcelProcessor_v2(self.filepath)
        df = processor.process()
        deducer = AttributeDataProcessorFenix()
        new_df = deducer.generate_new_dataframe_from_excel(df)
        key_decoder_data = new_df[new_df['Worksheet'] == 'TCR_Raw']
        columns_to_select = key_decoder_data['variable'].tolist()
        filtered_raw = raw[columns_to_select]
        self.check_null_columns(filtered_raw)
        rename_dict = dict(zip(key_decoder_data['variable'], key_decoder_data['Column']))
        raw = filtered_raw.rename(columns=rename_dict)
        respondent_columns = [col for col in raw.columns if col == 'Respondent']

        if len(respondent_columns) > 1:
            raw['Respondent'] = raw['Respondent'].apply(lambda row: '_'.join(row.dropna().astype(str)), axis=1)
            raw = raw.loc[:, ~raw.columns.duplicated()]
        self.get_date_range(raw)
        self.raw_data = raw

    def Consumer(self, sheet_name):
        consumer = pd.read_excel(self.filepath, sheet_name=sheet_name, dtype=str)
        processor = ExcelProcessor_v2(self.filepath)
        df = processor.process()
        deducer = AttributeDataProcessorFenix()
        new_df = deducer.generate_new_dataframe_from_excel(df)

        key_decoder_data = new_df[new_df['Worksheet'] == 'TCR_Consumer']

        # Check for necessary columns in key_decoder_data
        if 'variable' not in key_decoder_data.columns:
            print("Error: 'variable' column missing in key_decoder_data.")
            return

        columns_to_select = key_decoder_data['variable'].tolist()

        # Ensure selected columns exist in consumer
        missing_columns = [col for col in columns_to_select if col not in consumer.columns]
        if missing_columns:
            print(f"Warning: The following columns are missing in consumer: {missing_columns}")

        filtered_Consumer = consumer[columns_to_select]

        self.check_null_columns(filtered_Consumer)

        rename_dict = dict(zip(key_decoder_data['variable'], key_decoder_data['Column']))
        consumer = filtered_Consumer.rename(columns=rename_dict)

        if 'City' in consumer.columns:
            if 'Key' in self.lookup_data.columns and 'City' in self.lookup_data.columns:
                city_dict = self.lookup_data.set_index('Key')['City'].to_dict()
                consumer = consumer.dropna(subset=['City'])
                list_city = consumer['City'].replace(city_dict)
                # print(tabulate(consumer.head(10), headers='keys', tablefmt='fancy_grid', showindex=True))

                city_list = list_city.tolist()
                self.country_of_test = get_most_frequent_country_from_cities(city_list)
                consumer['Country'] = self.country_of_test
            else:
                print("Error: 'Key' or 'City' column missing in lookup_data.")

        respondent_columns = [col for col in consumer.columns if 'Respondent' in col]
        if len(respondent_columns) > 1:
            consumer['Respondent'] = consumer['Respondent'].apply(lambda row: '_'.join(row.dropna().astype(str)),
                                                                  axis=1)
            consumer = consumer.loc[:, ~consumer.columns.duplicated()]

        # print(tabulate(consumer.head(5), headers='keys', tablefmt='fancy_grid', showindex=True))

        self.consumer_data = consumer

    def Product(self,filepath):
        source = self.key_decoder
        data = source[source['Column'] == 'Product Code'].dropna(subset=['Description'])
        description_str = data['Description'].str.cat(sep=' ')
        trimmed_string = re.sub(r'^[^\d]*', '', description_str)
        pairs = trimmed_string.split(';')
        data_dict = {key.strip(): value.strip() for pair in pairs if '=' in pair for key, value in [pair.split('=', 1)]}
        df = pd.DataFrame(list(data_dict.items()), columns=['Product Code', 'Sample Description'])
        df = df[~df['Sample Description'].str.contains('cell', case=False, na=False)]
        result_str = ';'.join(df.apply(lambda row: f"{row['Product Code']}={row['Sample Description']}", axis=1))
        source.loc[source['Column'] == 'Product Code', 'Description'] = result_str
        self.key_decoder = source

        words_to_check = ['Control', 'current' 'Prototype', 'PROTO',
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

        df['Formula Spec Number'] = "000000"
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
        df = df[column_order]

        fn=os.path.splitext(os.path.basename(filepath))[0]
        df['Formula Spec Number'] = df['Product Code'].apply(
            lambda x: str(NumberIncrementer.get_product_code_suffix(self,fn)) +
                      (str(x).zfill(3)) + '-001' if pd.notnull(x) else None)
        # print(tabulate(df, headers='keys', tablefmt='fancy_grid', showindex=True))
        df.loc[df['Type of Product'] != 1, 'Formula Spec Number'] = None
        self.product_data = df

    def KeyDecoder(self):
        processor = ExcelProcessor_v2(self.filepath)
        df = processor.process()
        deducer = AttributeDataProcessorFenix()
        new_df = deducer.generate_new_dataframe_from_excel(df)
        filtered_df = new_df[new_df['Column'].str.contains('city', case=False, na=False)]
        if not filtered_df.empty:
            city_string = str(filtered_df['Description'].iloc[0])
            # print(city_string)
            city_df = self.create_city_dataframe(city_string)
            self.lookup_data = city_df
            # print(tabulate(city_df, headers='keys', tablefmt='fancy_grid', showindex=True))
        else:
            print("No descriptions found.")
        new_df = new_df[['Worksheet', 'Column', 'Description']]
        new_df.loc[len(new_df)] = ['TCR_Consumer', 'Country', 'Country of Test']
        new_df = new_df.drop_duplicates(subset=['Worksheet', 'Column'], keep='first', inplace=False)

        self.key_decoder = new_df

    def Project(self):
        self.project_data = DataFrameEnhancerFenix(None, self.filepath, self.country_of_test, self.start_date,
                                                   self.end_date).enhance_dataframe()

    @staticmethod
    def create_city_dataframe(city_string):
        city_pairs = city_string.split(' ; ')
        key_value_pairs = [pair.split('=') for pair in city_pairs]
        city_df = pd.DataFrame(key_value_pairs, columns=['Key', 'City'])
        city_df['Key'] = city_df['Key']
        return city_df

    def get_date_range(self, raw):
        # Check if the 'Date' column exists in the dataframe
        if 'Date' not in raw.columns:
            print("Error: 'Date' column not found.")
            return None, None  # Return None if the column is missing

        # Ensure the 'Date' column is in datetime format
        raw['Date'] = pd.to_datetime(raw['Date'], errors='coerce')

        # Check if there is any data in the 'Date' column
        if raw['Date'].notna().any():  # Checks if there is any non-null date
            # Get the min and max date from the 'Date' column
            self.start_date = raw['Date'].min()
            self.end_date = raw['Date'].max()
            return self.start_date, self.end_date  # Return the computed dates
        else:
            print("No valid data in the 'Date' column.")
            return None, None  # Return None if no valid dates exist

    def reversal(self):
        processor = DescriptionProcessor()
        processed_df, map = processor.process_dataframe(self.key_decoder)
        raw_df = self.raw_data
        consumer_df = self.consumer_data
        raw_df = processor.apply_column_mapping(raw_df, map)
        null_columns1 = raw_df.columns[raw_df.isna().all()].tolist()
        null_columns2 = consumer_df.columns[consumer_df.isna().all()].tolist()
        merged_null_columns = null_columns1 + null_columns2
        raw_df = raw_df.drop(columns=null_columns1)
        consumer_df = consumer_df.drop(columns=null_columns2)
        processed_df = processed_df[~processed_df['Column'].isin(merged_null_columns)]
        self.key_decoder = processed_df
        self.raw_data = raw_df
        self.consumer_data = consumer_df





class DescriptionProcessor:
    def transform_description(self, description):
        pairs = description.split('; ')
        min_key = int(pairs[0].split('=')[0])
        max_key = int(pairs[-1].split('=')[0])
        scale_factor = min_key + max_key
        transformed_pairs = []

        for pair in pairs:
            key, value = pair.split('=')
            try:
                new_key = scale_factor - int(key)
            except ValueError:
                new_key = key
            transformed_pairs.append((new_key, value))

        # Sort based on the key, handling both integers and strings
        transformed_pairs.sort(key=lambda x: (
            int(x[0]) if isinstance(x[0], str) and x[0].isdigit() else float('inf') if isinstance(x[0], str) else x[0]))

        # Rebuild the string from the sorted pairs
        transformed_pairs = [f"{key}={value}" for key, value in transformed_pairs]

        return "; ".join(transformed_pairs), scale_factor

    def process_dataframe(self, df):
        pd.options.mode.chained_assignment = None
        keywords = [
             'purchase intent',
             'carbonation',
             'balance of sweetness to tartness/sourness',
             'sweetness',
             'color',
             'flavor',
             'overall flavor',
             'expectations',
             'aroma',
             'tartness/sourness',
             'overall',
             ]


        filtered_df = df[df['Column'].str.contains('|'.join(keywords), case=False, na=False)]
        filtered_df['Description'], filtered_df['Scale_Factor'] = zip(*filtered_df['Description'].apply(
            lambda x: self.transform_description(x)
        ))
        df.loc[filtered_df.index, ['Description', 'Scale_Factor']] = filtered_df[['Description', 'Scale_Factor']]
        column_scale_dict = {row['Column']: row['Scale_Factor'] for _, row in filtered_df.iterrows()}
        df = df.drop(columns=['Scale_Factor'], errors='ignore')
        return df, column_scale_dict

    def apply_column_mapping(self, raw_df, column_scale_dict):
        for column in raw_df.columns:
            if column in column_scale_dict:

                # Apply the operation (value from dict - x) to the values in the column
                raw_df[column] = raw_df[column].apply(
                    lambda x: column_scale_dict[column] - int(pd.to_numeric(x, errors='coerce')) if pd.notnull(x) else x
                )




        return raw_df


if __name__ == "__main__":
    directory_path = r"C:\Users\Z19661\PycharmProjects\LegacyData_\InputFiles\fenix\Monadic\08_DATA_Cleaned_AG_Fanta_Regular V02.xlsx"
    Fenix_ETL(directory_path)
