import os
import re
from _pydatetime import datetime, timedelta

from CODES.EMEA.keydecoder_deduce_EMEA import  AttributeDataProcessor
import pandas as pd

from CODES.EMEA.project_tab_determination import DataFrameEnhancer
from CODES.database.INV_Generator import NumberIncrementer


# Load the mapping data


class EMEA_ETL:


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
        self.Product(source=AttributeDataProcessor().generate_new_dataframe_from_excel(file_path=self.filepath,sheet_name='Decoder'))
        self.reversal()

    def RawData(self, sheet_name):
        """Fetch and process the raw data."""
        raw = pd.read_excel(self.filepath, sheet_name=sheet_name)
        new_df = AttributeDataProcessor().generate_new_dataframe_from_excel(file_path=self.filepath,sheet_name='Decoder')
        key_decoder_data = new_df[new_df['Worksheet'] == 'TCR_Raw']
        columns_to_select = key_decoder_data['variable_for_automation'].tolist()
        filtered_raw = raw[columns_to_select]

        rename_dict = dict(zip(key_decoder_data['variable_for_automation'], key_decoder_data['Column']))
        raw = filtered_raw.rename(columns=rename_dict)

        def excel_to_datetime(excel_serial):
            base_date = datetime(1900, 1, 1)  # Excel's base date (1900-01-01)
            # Excel counts days starting from 1, so adjust for leap year bug and 1-based indexing
            return base_date + timedelta(days=excel_serial - 2)

        raw['Date'] = raw['Date'].apply(lambda x: excel_to_datetime(x) if isinstance(x, (int, float)) else x)

        self.raw_data = raw

    def Consumer(self, sheet_name):
        """Fetch and process the consumer data."""
        consumer = pd.read_excel(self.filepath, sheet_name=sheet_name)
        new_df = AttributeDataProcessor().generate_new_dataframe_from_excel(file_path=self.filepath,
                                                                            sheet_name='Decoder')
        key_decoder_data = new_df[new_df['Worksheet'] == 'TCR_Consumer']
        columns_to_select = key_decoder_data['variable_for_automation'].tolist()
        filtered_raw = consumer[columns_to_select]

        rename_dict = dict(zip(key_decoder_data['variable_for_automation'], key_decoder_data['Column']))
        consumer = filtered_raw.rename(columns=rename_dict)

        df_abn = pd.read_excel(self.filepath, sheet_name='Additional MSE data')
        df_abn = df_abn.iloc[:, :2].dropna()
        data_map = df_abn.set_index(df_abn.columns[0]).to_dict()[df_abn.columns[1]]
        country_of_test = data_map.get('Country of Test', 'Unknown')
        consumer['Country'] = country_of_test
        # print(tabulate(consumer, headers='keys', tablefmt='fancy_grid', showindex=True))
        self.consumer_data = consumer

    def Product(self, source=None):
        source = source if source is not None else self.key_decoder
        data = source[source['variable_for_automation'] == 'Product ID'].dropna(subset=['Description'])
        description_str = data['Description'].str.cat(sep=' ')
        trimmed_string = re.sub(r'^[^\d]*', '', description_str)
        pairs = trimmed_string.split(';')
        data_dict = {key.strip(): value.strip() for pair in pairs if '=' in pair for key, value in [pair.split('=', 1)]}
        df = pd.DataFrame(list(data_dict.items()), columns=['Product Code', 'Sample Description'])
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

        df['Formula Spec Number'] = '1234678-001'
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

        fn = os.path.splitext(os.path.basename(self.filepath))[0]
        df['Formula Spec Number'] = df['Product Code'].apply(
            lambda x: str(NumberIncrementer.get_product_code_suffix(self, fn)) +
                      (str(x).zfill(3)) + '-001' if pd.notnull(x) else None)
        # print(tabulate(df, headers='keys', tablefmt='fancy_grid', showindex=True))
        df.loc[df['Type of Product'] != 1, 'Formula Spec Number'] = None


        self.product_data = df

    def KeyDecoder(self, sheet_name):
        keydecoder = AttributeDataProcessor().generate_new_dataframe_from_excel(file_path=self.filepath,sheet_name=sheet_name)
        keydecoder = keydecoder[['Worksheet', 'Column', 'Description']]
        keydecoder.loc[len(keydecoder)] = ['TCR_Consumer', 'Country', 'Country of Test']
        keydecoder.loc[keydecoder['Column'] == 'Age', 'Description'] = 'Exact Age'
        # print(tabulate(keydecoder, headers='keys', tablefmt='fancy_grid', showindex=True))
        self.key_decoder=keydecoder

    def Project(self, sheet_name):
        df = pd.read_excel(self.filepath, sheet_name=sheet_name,header=1)

        if 'field' in df.columns and 'value' in df.columns:
            print("OK:Columns 'field' and 'value' already exist, no changes made.")
        else:
            df.columns = ['Field', 'Value'] + df.columns[2:].tolist()

        df = df[['Field', 'Value']]
        # print(tabulate(df, headers='keys', tablefmt='fancy_grid', showindex=True))
        enhancer = DataFrameEnhancer(df,self.filepath)
        df = enhancer.enhance_dataframe()
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

        # print(tabulate(df, headers='keys', tablefmt='fancy_grid', showindex=True))
        self.project_data = df

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
             'Gender'
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


