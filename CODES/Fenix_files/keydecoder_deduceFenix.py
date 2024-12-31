import pandas as pd
from tabulate import tabulate


class AttributeDataProcessorFenix:
    def __init__(self):
        self.attribute_dict = self.get_attribute_dict()

    def get_attribute_dict(self):
        attribute_dict = {
            'Respondent': {
                'column_positive': ['Anonymised_ID','participantid','serial'],
                'description_positive': [],
                'column_negative': [],
                'description_negative': []
            },
            'Date': {
                'column_positive': ['DataCollection_StartTime','date_bookingdatetime'],
                'description_positive': ['time','start'],
                'column_negative': [],
                'description_negative': []
            },
            'Cell no': {
                'column_positive': ['sampleallocation_final'],
                'description_positive': [],
                'column_negative': [],
                'description_negative': []
            },
            'Cell': {
                'column_positive': ['sampleallocation_final'],
                'description_positive': [],
                'column_negative': [],
                'description_negative': []
            },
            'Product Code': {
                'column_positive': ['CellAllocation_Final_Recode'],
                'description_positive': [],
                'column_negative': ['ProductNameReport_Recode - DO NOT USE','use'],
                'description_negative': []
            },
            'Overall_LIK': {
                'column_positive': ['Overall_Recode'],
                'description_positive': [],
                'column_negative': [],
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
                'column_positive': [],
                'description_positive': ['Just about right+sweet'],
                'column_negative': [],
                'description_negative': ['too sweet', 'not sweet enough']
            },
            'Tartness/Sourness_INT': {
                'column_positive': [],
                'description_positive': [],
                'column_negative': [],
                'description_negative': []
            },
            'Overall Flavor_JAR': {
                'column_positive': ['adj_3'],
                'description_positive': ['Just about right'],
                'column_negative': [],
                'description_negative': []
            },
            'Sweetness_JAR': {
                'column_positive': [],
                'description_positive': ['Just about right+sweet'],
                'column_negative': [],
                'description_negative': ['sour']
            },
            'Tartness/Sourness_JAR': {
                'column_positive': [],
                'description_positive': ['Just about right + sour/tart'],
                'column_negative': [],
                'description_negative': []
            },
            'Color_JAR': {
                'column_positive': ['adjective_1'],
                'description_positive': ['Just about right'],
                'column_negative': [],
                'description_negative': []
            },
            'Carbonation_JAR': {
                'column_positive': [],
                'description_positive': ['Just about right + fizzy'],
                'column_negative': [],
                'description_negative': []
            },
            'Aroma_JAR': {
                'column_positive': ['adj_2'],
                'description_positive': ['Just about right'],
                'column_negative': [],
                'description_negative': []
            },
            'Mouthfeel_JAR': {
                'column_positive': ['adj_8'],
                'description_positive': ['Just about right'],
                'column_negative': [],
                'description_negative': []
            },
            'Cola Flavor_JAR': {
                'column_positive': ['Cola flavor'],
                'description_positive': ['Just about right'],
                'column_negative': [],
                'description_negative': []
            },
            'Orange Flavor_JAR': {
                'column_positive': [],
                'description_positive': [],
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
                'column_positive': [],
                'description_positive': ['Just about right+sour+sweet'],
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
                'column_positive': ['expect'],
                'description_positive': [],
                'column_negative': [],
                'description_negative': []
            },
            'Purchase Intent_PI': {
                'column_positive': ['Purchase Intent', 'Purchase', 'buy','pi'],
                'description_positive': [],
                'column_negative': [],
                'description_negative': []
            },
            'Satisfaction_SAT': {
                'column_positive': [],
                'description_positive': ['Satisfaction', 'satisfied'],
                'column_negative': [],
                'description_negative': []
            },
            # Consumer Fields
            'City': {
                'column_positive': ['Location', 'HALL'],
                'description_positive': [],
                'column_negative': ['allocation'],
                'description_negative': []
            },
            'Gender': {
                'column_positive': ['Gender', 'sGender'],
                'description_positive': [],
                'column_negative': [],
                'description_negative': []
            },
            'Age': {
                'column_positive': ['Exact Age','srvage'],
                'description_positive': [],
                'column_negative': [],
                'description_negative': ['under', 'birthday']
            },
            'Age Group': {
                'column_positive': ['Agerange'],
                'description_positive': ['under', ],
                'column_negative': [],
                'description_negative': []
            },
            # Uncomment the following if needed
            # 'User Group': {
            #     'column_positive': ['Target', 'SAMPLEUSER'],
            #     'description_positive': [],
            #     'column_negative': [],
            #     'description_negative': []
            # }
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
        # Check that no words in word_list are present in the text
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
            "Cell", "Respondent", "City", "Country", "Gender", "Age", "Ethnicity", "User Group","Age Group","Country"
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

            reporting_name = row['Key']
            description = row['Description']
            if not isinstance(reporting_name, str) or not isinstance(description, str):
                continue

            for key, conditions in self.attribute_dict.items():
                column_check_list = conditions['column_positive']
                description_check_list = conditions['description_positive']
                column_negative_check_list = conditions['column_negative']
                description_negative_check_list = conditions['description_negative']
                if not column_check_list and not description_check_list and not column_negative_check_list and not description_negative_check_list:
                    continue
                # Check conditions only if the corresponding list is non-empty
                column_check_match = (
                            not column_check_list or self.check_word_in_string(column_check_list, reporting_name))
                description_check_match = (
                            not description_check_list or self.check_word_in_string(description_check_list,
                                                                                    description))
                column_negative_check_match = (
                            not column_negative_check_list or self.check_no_word_in_string(column_negative_check_list,
                                                                                           reporting_name))
                description_negative_check_match = (not description_negative_check_list or self.check_no_word_in_string(
                    description_negative_check_list, description))

                if (
                        column_check_match and column_negative_check_match and description_check_match and description_negative_check_match):
                    new_data.append({
                        "variable" :reporting_name,
                        "Column": key,
                        "Description": description,
                    })

        # Create a DataFrame with the filtered data
        new_df = pd.DataFrame(new_data)
        return new_df

    def generate_new_dataframe_from_excel(self,df):
        # Read the Excel file


        # Generate the new dataframe based on the attribute_dict
        new_dataframe = self.generate_new_dataframe(df)

        # Optionally merge with other data if needed (e.g., raw_consumer data)
        determiner_df = self.raw_consumer()
        merged_df = pd.merge(new_dataframe, determiner_df, on='Column', how='left')

        # Print the new dataframe with the merged data
        # print("New DataFrame based on attribute conditions:")
        # print(tabulate(merged_df, headers='keys', tablefmt='fancy_grid', showindex=True))

        return merged_df
