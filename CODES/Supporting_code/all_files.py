import os
import pandas as pd
from pyxlsb import open_workbook


def find_xlsx_files_to_dataframe(directory):
    file_data = []  # List to store tuples of (file_name, file_path)
    for root, dirs, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            file_data.append((file, file_path))  # Append as a tuple (file_name, file_path)

    # Create a DataFrame from the list of tuples
    df = pd.DataFrame(file_data, columns=['File Name', 'File Path'])
    return df

def get_sheets_in_xlsb(file_path):
    try:
        with open_workbook(file_path) as wb:
            return wb.sheets
    except Exception as e:
        return str(e)

if __name__ == "__main__":
    your_directory_path = r"C:\Legacy Files\all_files\For TCS\EMEA"
    df = find_xlsx_files_to_dataframe(your_directory_path)

    # Save DataFrame to a CSV file (optional)
    output_file = '../res/output.xlsx_files.csv'
    df.to_csv(output_file, index=False)

    print(f"DataFrame created with {len(df)} .xlsx files.")
    print(df)
