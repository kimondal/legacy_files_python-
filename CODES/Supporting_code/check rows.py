import pandas as pd


def validate_columns(file_path):
    # Step 1: Read the Excel file and load the specific sheets into DataFrames
    tcr_raw_df = pd.read_excel(file_path, sheet_name='TCR_Raw')
    tcr_key_decoder_df = pd.read_excel(file_path, sheet_name='TCR_Key_Decoder')

    # Step 2: Extract columns from TCR_Raw and TCR_Key_Decoder
    tcr_raw_columns = tcr_raw_df.columns.tolist()  # Get the list of column names from TCR_Raw

    # Assuming that the column name with the list of columns in the TCR_Key_Decoder is called 'Column'
    key_decoder_columns = tcr_key_decoder_df[
        'Column'].dropna().tolist()  # Get the list of valid columns from TCR_Key_Decoder

    # Step 3: Validate columns in TCR_Raw against TCR_Key_Decoder
    raw_columns_valid = all(col in key_decoder_columns for col in tcr_raw_columns)

    # Step 4: Find the columns in TCR_Raw that are not present in TCR_Key_Decoder's 'columns' list
    columns_not_in_key_decoder = [col for col in tcr_raw_columns if col not in key_decoder_columns]

    # Step 5: Output validation result
    if not raw_columns_valid:
        print("Some columns in TCR_Raw are not listed in TCR_Key_Decoder's 'Column' list:")
        for col in columns_not_in_key_decoder:
            print(col)
    else:
        print("All columns in TCR_Raw are present in TCR_Key_Decoder's 'Column' list.")

    # Optional: Return whether all columns are valid and the missing columns
    return raw_columns_valid, columns_not_in_key_decoder


# Example usage
file_path = r'C:\Users\Z19661\PycharmProjects\LegacyData_\OutputFiles\tcr_Minute Maid Refreshment CLT - White Lemonade Phase 2_Full Data File with Key.xlsx'
valid, missing_columns = validate_columns(file_path)

if not valid:
    print("Missing columns:", missing_columns)
else:
    print("Validation successful. All columns are valid.")
