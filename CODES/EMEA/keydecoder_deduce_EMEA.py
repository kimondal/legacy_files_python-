import pandas as pd
from tabulate import tabulate

import pandas as pd
from tabulate import tabulate


class AttributeDataProcessor:
    def __init__(self):
        self.attribute_dict = self.get_attribute_dict()

    def get_attribute_dict(self):
        attribute_dict = {
            'Respondent': {
                'column_positive': ["res+id"],
                'description_positive': [],
                'column_negative': [],
                'description_negative': []
            },
            'Date': {
                'column_positive': ['Date'],
                'description_positive': [],
                'column_negative': [],
                'description_negative': []
            },
            'Cell no': {
                'column_positive': ['cell'],
                'description_positive': ['cell'],
                'column_negative': ['group'],
                'description_negative': []
            },
            'Cell': {
                'column_positive': ['cell'],
                'description_positive': [],
                'column_negative': ['Group'],
                'description_negative': []
            },
            'Product Code': {
                'column_positive': ['Product','code of test product'],
                'description_positive': ['prototype','type','current'],
                'column_negative': [],
                'description_negative': []
            },
            'Overall_LIK': {
                'column_positive': ['lik+bot'],
                'description_positive': ['like'],
                'column_negative': ['sip'],
                'description_negative': []
            },
            'Overall Flavor_LIK': {
                'column_positive': ['Overall Flavor'],
                'description_positive': ['Just about right'],
                'column_negative': [],
                'description_negative': []
            },
            'Appearance_LIK': {
                'column_positive': ['Appearance'],
                'description_positive': ['Just about right'],
                'column_negative': [],
                'description_negative': []
            },
            'Likes_OE': {
                'column_positive': ['Likes'],
                'description_positive': [],
                'column_negative': [],
                'description_negative': []
            },
            'Dislikes_OE': {
                'column_positive': ['Dislikes'],
                'description_positive': [],
                'column_negative': [],
                'description_negative': []
            },
            'Sweetness_INT': {
                'column_positive': ['Sweetness'],
                'description_positive': ['Just about right'],
                'column_negative': [],
                'description_negative': ['too sweet', 'not sweet enough']
            },
            'Tartness/Sourness_INT': {
                'column_positive': ['Tartness/Sourness'],
                'description_positive': ['Just about right'],
                'column_negative': [],
                'description_negative': ['too tart', 'too sour']
            },
            'Overall Flavor_JAR': {
                'column_positive': ['Flavour strength'],
                'description_positive': ['Just about right'],
                'column_negative': [],
                'description_negative': []
            },
            'Sweetness_JAR': {
                'column_positive': ['Sweetness'],
                'description_positive': ['Just about right'],
                'column_negative': [],
                'description_negative': ['sour']
            },
            'Tartness/Sourness_JAR': {
                'column_positive': ['Sourness / Tartness'],
                'description_positive': ['Just about right'],
                'column_negative': [],
                'description_negative': []
            },
            'Color_JAR': {
                'column_positive': ['Color','COLOUR'],
                'description_positive': ['Just about right'],
                'column_negative': [],
                'description_negative': []
            },
            'Carbonation_JAR': {
                'column_positive': ['Fizziness', 'Carbonation'],
                'description_positive': ['Just about right'],
                'column_negative': [],
                'description_negative': []
            },
            'Aroma_JAR': {
                'column_positive': ['Aroma'],
                'description_positive': ['Just about right'],
                'column_negative': [],
                'description_negative': []
            },
            'Mouthfeel_JAR': {
                'column_positive': ['Mouthfeel'],
                'description_positive': ['Just about right'],
                'column_negative': ['dry'],
                'description_negative': []
            },
            'Cola Flavor_JAR': {
                'column_positive': ['Cola flavor'],
                'description_positive': ['Just about right'],
                'column_negative': [],
                'description_negative': []
            },
            'Orange Flavor_JAR': {
                'column_positive': ['Orange flavor'],
                'description_positive': ['Just about right'],
                'column_negative': [],
                'description_negative': []
            },
            'Lemon Flavor_JAR': {
                'column_positive': ['Lemon flavor'],
                'description_positive': ['Just about right'],
                'column_negative': [],
                'description_negative': []
            },
            'Fruit Flavor_JAR': {
                'column_positive': ['Fruit flavor'],
                'description_positive': ['Just about right'],
                'column_negative': [],
                'description_negative': []
            },
            'Balance of Sweetness to Tartness/Sourness_BAL': {
                'column_positive': ['Balance'],
                'description_positive': ['Just about right'],
                'column_negative': [],
                'description_negative': []
            },
            'Aftertaste Detection_AT': {
                'column_positive': ['Aftertaste Detection'],
                'description_positive': [],
                'column_negative': [],
                'description_negative': []
            },
            'Aftertaste Pleasantness_AT': {
                'column_positive': ['Aftertaste Pleasantness'],
                'description_positive': [],
                'column_negative': [],
                'description_negative': []
            },
            'Aftertaste_AT': {
                'column_positive': ['Aftertaste'],
                'description_positive': [],
                'column_negative': [],
                'description_negative': []
            },
            'Expectations_EXP': {
                'column_positive': ['expec'],
                'description_positive': ['expec'],
                'column_negative': [],
                'description_negative': []
            },
            'Purchase Intent_PI': {
                'column_positive': ['Purchase Intent', 'Purchase', 'buy'],
                'description_positive': [],
                'column_negative': [],
                'description_negative': []
            },
            'Satisfaction_SAT': {
                'column_positive': ['Satisfaction', 'satisfied'],
                'description_positive': ['Satisfaction', 'satisfied'],
                'column_negative': [],
                'description_negative': []
            },

            # /////////////CONSUMER///////////////

            'City': {
                'column_positive': ['Location', 'HALL'],
                'description_positive': [],
                'column_negative': [],
                'description_negative': []
            },

            'Gender': {
                'column_positive': ['Gender', 'sGender'],
                'description_positive': [],
                'column_negative': [],
                'description_negative': []
            },

            'Age': {
                'column_positive': ['Exact Age', 'Age'],
                'description_positive': [],
                'column_negative': [],
                'description_negative': ["-"]
            },

            'Age Group': {
                'column_positive': ['Exact Age', 'Age'],
                'description_positive': ["-"],
                'column_negative': [],
                'description_negative': []
            },

            'User Group': {
                'column_positive': ['Target', 'SAMPLEUSER'],
                'description_positive': [],
                'column_negative': [],
                'description_negative': []
            }
        }
        return attribute_dict

    def check_word_in_string(self, word_list, text):
        # If the list contains '+', it's an "AND" check (all words must be present)
        if '+' in word_list[0]:
            for part in word_list[0].split('+'):
                if part.strip().lower() not in text.lower():
                    return False
            return True
        # Otherwise, it's an "OR" check (any one word must be present)
        else:
            return any(word.lower() in text.lower() for word in word_list)

    def check_no_word_in_string(self, word_list, text):
        return all(word.lower() not in text.lower() for word in word_list)

    def raw_consumer(self):
        TCR_Raw = [
            "Respondent", "Date", "Cell no", "Product Code", "Overall_LIK", "Overall Flavor_LIK",
            "Appearance_LIK", "Likes_OE", "Dislikes_OE", "Sweetness_INT", "Tartness/Sourness_INT",
            "Overall Flavor_JAR", "Sweetness_JAR", "Tartness/Sourness_JAR", "Color_JAR", "Carbonation_JAR",
            "Aroma_JAR", "Mouthfeel_JAR", "Cola Flavor_JAR", "Orange Flavor_JAR", "Lemon Flavor_JAR",
            "Fruit Flavor_JAR", "Balance of Sweetness to Tartness/Sourness_BAL", "Aftertaste Detection_AT",
            "Aftertaste Pleasantness_AT", "Aftertaste_AT", "Expectations_EXP", "Purchase Intent_PI",
            "Satisfaction_SAT", "Ageement Statement1_AGR", "Ageement Statement2_AGR", "Ageement Statement3_AGR"
        ]

        TCR_Consumer = [
            "Cell",'Age Group', "Respondent", "City", "Country", "Gender", "Age", "Ethnicity", "User Group"
        ]

        # Prepare the data
        data = []
        for col in TCR_Raw:
            data.append(["TCR_Raw", col])

        for col in TCR_Consumer:
            data.append(["TCR_Consumer", col])

        # Create DataFrame
        df = pd.DataFrame(data, columns=["Worksheet", "Column"])

        return df





    def generate_new_dataframe(self, df):
        new_data = []

        for _, row in df.iterrows():
            reporting_name = row['Name for reporting']
            description = row['Scale']
            variable_for_automation = row['Variable for Automation']

            if not isinstance(reporting_name, str) or not isinstance(description, str):
                continue

            for key, conditions in self.attribute_dict.items():
                column_check_list = conditions['column_positive']
                description_check_list = conditions['description_positive']
                column_negative_check_list = conditions['column_negative']
                description_negative_check_list = conditions['description_negative']

                # Check conditions only if the corresponding list is non-empty
                column_check_match = (not column_check_list or self.check_word_in_string(column_check_list, reporting_name))
                description_check_match = (not description_check_list or self.check_word_in_string(description_check_list, description))
                column_negative_check_match = (not column_negative_check_list or self.check_no_word_in_string(column_negative_check_list, reporting_name))
                description_negative_check_match = (not description_negative_check_list or self.check_no_word_in_string(description_negative_check_list, description))

                if (
                    column_check_match and
                    column_negative_check_match and
                    description_check_match and
                    description_negative_check_match
                ):
                    new_data.append({
                        "reporting_name": reporting_name,
                        "Column": key,
                        "Description": description,
                        "variable_for_automation": variable_for_automation,
                    })

        new_df = pd.DataFrame(new_data)
        return new_df








    def generate_new_dataframe_from_excel(self, file_path, sheet_name):
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        # print(tabulate(df, headers='keys', tablefmt='fancy_grid', showindex=True))

        new_dataframe = self.generate_new_dataframe(df)

        # Space for alignment
        determiner_df = self.raw_consumer()

        merged_df = pd.merge(new_dataframe, determiner_df, on='Column', how='left')
        new_data = {
            "reporting_name": ["QHIDUSERID", "QHIDUSERID"],
            "Column": ["Respondent", "Respondent"],
            "Description": ["ID of Respondent", "ID of Respondent"],
            "variable_for_automation": ["QHIDUSERID", "QHIDUSERID"],
            "Worksheet": ["TCR_Raw", "TCR_Consumer"]
        }
        df_new = pd.DataFrame(new_data)
        if 'Respondent' not in merged_df['Column'].values:
            merged_df = pd.concat([df_new, merged_df], ignore_index=True)

        # Return the new DataFrame
        return merged_df



# processor = AttributeDataProcessor()
# new_df = processor.generate_new_dataframe_from_excel(
#     file_path=r'C:\Users\Z19661\PycharmProjects\LegacyData_\InputFiles\EMEA_ETL\P320121 CCC RTD Tea prototype testing MSE Data FINAL V2.xlsb',
#     sheet_name='Decoder'
# )
# print(tabulate(new_df, headers='keys', tablefmt='fancy_grid', showindex=True))





