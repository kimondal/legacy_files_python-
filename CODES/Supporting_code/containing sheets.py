import os
import pandas as pd
from pyxlsb import open_workbook


def find_xlsx_files_to_dataframe(directory):
    file_data = []  # List to store tuples of (file_name, file_path, sheets)

    for root, dirs, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            if file.endswith(".xlsb"):  # Check for .xlsb files
                try:
                    sheet_names = get_sheets_in_xlsb(file_path)
                    file_data.append((file, file_path, sheet_names))  # Add sheet names for .xlsb
                except Exception as e:
                    file_data.append((file, file_path, f"Error: {str(e)}"))
            elif file.endswith(".xlsx"):  # Check for .xlsx files
                file_data.append((file, file_path, "N/A"))  # Placeholder for .xlsx files

    # Create a DataFrame from the list of tuples
    df = pd.DataFrame(file_data, columns=['File Name', 'File Path', 'Sheets'])
    return df


def get_sheets_in_xlsb(file_path):
    try:
        with open_workbook(file_path) as wb:
            return wb.sheets  # Return sheet names as a list of strings
    except Exception as e:
        return str(e)


if __name__ == "__main__":
    your_directory_path = r"C:\Legacy Files\all_files\For TCS"
    df = find_xlsx_files_to_dataframe(your_directory_path)

    # Save DataFrame to a CSV file (optional)
    output_file = '../res/outputfiles.csv'
    df.to_csv(output_file, index=False)

    print(f"DataFrame created with {len(df)} files.")
    print(df)
