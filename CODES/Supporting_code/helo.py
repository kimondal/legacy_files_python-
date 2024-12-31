import pandas as pd
from nltk.sentiment.vader import SentimentIntensityAnalyzer
from tabulate import tabulate

# Sample data (full dataframe for testing)
data = [
    {"Variable": "Aftertaste_Pleasant_Binary_Total_recode",
     "Column": "Aftertaste_AT",
     "Description": "1=Unpleasant ; 2=No aftertaste ; 3=Pleasant",
     "Worksheet": "TCR_Raw"},
    {"Variable": "CellAllocation_Final_Recode",
     "Column": "Product Code",
     "Description": "1=Cell 1 ; 2=Cell 2 ; 3=Cell 3 ; 4=Cell 4 ; 5=Cell 5",
     "Worksheet": "TCR_Raw"},
    {"Variable": "JAR_Std_5_Bi_ScaleSpec_Adj_2_Recode",
     "Column": "Aroma_JAR",
     "Description": "1=Much too weak ; 2=A little too weak ; 3=Just about right ; 4=A little too strong ; 5=Much too strong",
     "Worksheet": "TCR_Raw"},
    {"Variable": "JAR_Std_5_Bi_ScaleSpec_Adj_3_Recode",
     "Column": "Overall Flavor_JAR",
     "Description": "1=Much too weak ; 2=A little too weak ; 3=Just about right ; 4=A little too strong ; 5=Much too strong",
     "Worksheet": "TCR_Raw"},
    {"Variable": "JAR_Std_5_Bi_ScaleSpec_Adj_6_Recode",
     "Column": "Balance of Sweetness to Tartness/Sourness_BAL",
     "Description": "1=A little too sour / Not quite sweet enough ; 2=Much too sour / Not nearly sweet enough ; 3=Just about right ; 4=Not quite sour enough / A little too sweet ; 5=Not nearly sour enough / Much too sweet",
     "Worksheet": "TCR_Raw"},
    {"Variable": "JAR_Std_5_Bi_ScaleSpec_Adj_8_Recode",
     "Column": "Mouthfeel_JAR",
     "Description": "1=Much too thin ; 2=A little too thin ; 3=Just about right ; 4=A little too thick ; 5=Much too thick",
     "Worksheet": "TCR_Raw"},
    {"Variable": "JAR_Std_5_Uni_ScaleSpec_Adjective_1_Recode",
     "Column": "Color_JAR",
     "Description": "1=Much too orange ; 2=A little too orange ; 3=Just about right ; 4=Not orange enough ; 5=Not at all orange enough",
     "Worksheet": "TCR_Raw"},
    {"Variable": "JAR_Std_5_Uni_ScaleSpec_Adjective_4_Recode",
     "Column": "Sweetness_JAR",
     "Description": "1=Not quite sweet enough ; 2=Not nearly sweet enough ; 3=Just about right ; 4=A little too sweet ; 5=Much too sweet",
     "Worksheet": "TCR_Raw"},
    {"Variable": "JAR_Std_5_Uni_ScaleSpec_Adjective_5_Recode",
     "Column": "Tartness/Sourness_JAR",
     "Description": "1=Much too sour/tart ; 2=A little too sour/tart ; 3=Just about right ; 4=Not quite sour/tart enough ; 5=Not nearly sour/tart enough",
     "Worksheet": "TCR_Raw"},
    {"Variable": "JAR_Std_5_Uni_ScaleSpec_Adjective_7_Recode",
     "Column": "Carbonation_JAR",
     "Description": "1=Much too fizzy ; 2=A little too fizzy ; 3=Just about right ; 4=Not quite fizzy enough ; 5=Not nearly fizzy enough",
     "Worksheet": "TCR_Raw"},
]


# Create a DataFrame
df = pd.DataFrame(data)



def transform_description(description):
    # Split the description into individual key-value pairs
    pairs = description.split('; ')

    # Get the min_key (first key) and max_key (last key)
    min_key = int(pairs[0].split('=')[0])  # First key is the min_key
    max_key = int(pairs[-1].split('=')[0])  # Last key is the max_key
    scale_factor = min_key + max_key

    transformed_pairs = []

    # For each pair, calculate the new key and add it to the transformed list
    for pair in pairs:
        key, value = pair.split('=')
        new_key = scale_factor - int(key)
        transformed_pairs.append(f"{new_key}={value}")

    return "; ".join(transformed_pairs), scale_factor  # Return both transformed description and scale_factor
def process_dataframe(df):
    # List of keywords to filter columns by
    keywords = ['Aroma', 'Overall Flavor', 'Mouthfeel', 'Color']
    # Step 1: Filter rows where the 'Column' contains any word in the list
    filtered_df = df[df['Column'].str.contains('|'.join(keywords), case=False, na=False)]

    # Step 2: Apply the transformation to the filtered rows
    filtered_df[['Transformed_Description', 'Scale_Factor']] = filtered_df['Description'].apply(
        lambda x: pd.Series(transform_description(x)))

    # Step 3: Return the dataframe with the required transformations
    return filtered_df


# Create a DataFrame
df = pd.DataFrame(data)
processed_df = process_dataframe(df)
print(tabulate(processed_df, headers='keys', tablefmt='grid'))
