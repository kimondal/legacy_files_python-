import os
import pandas as pd
import shutil


def check_and_copy_files(input_folder):
    # Ensure x and y folders exist, create if they don't
    x_folder = os.path.join(input_folder, 'x')
    y_folder = os.path.join(input_folder, 'y')

    if not os.path.exists(x_folder):
        os.makedirs(x_folder)
    if not os.path.exists(y_folder):
        os.makedirs(y_folder)

    # List all files in the folder
    for filename in os.listdir(input_folder):
        file_path = os.path.join(input_folder, filename)

        # Skip directories
        if os.path.isdir(file_path):
            continue

        # Only process Excel files
        if filename.endswith(('.xlsx', '.xls')):
            try:
                # Read the TCR_Project sheet from the Excel file
                df = pd.read_excel(file_path, sheet_name='TCR_Project')

                # Convert the dataframe into a dictionary with keys and values
                key_value_dict = dict(zip(df.iloc[:, 0], df.iloc[:, 1]))

                # Define the required keys
                required_keys = [
                    'Start Date of Study Field (mm/dd/yyyy)',
                    'End Date of Study Field (mm/dd/yyyy)',
                    'Country of Test',
                    'WorkFront Project Number'
                ]

                # Check if any required key is missing or has a null value
                missing_values = False
                for key in required_keys:
                    if key not in key_value_dict or pd.isnull(key_value_dict[key]):
                        missing_values = True
                        break

                # Copy the file to the appropriate folder
                if missing_values:
                    shutil.copy(file_path, os.path.join(x_folder, filename))
                else:
                    shutil.copy(file_path, os.path.join(y_folder, filename))

            except Exception as e:
                print(f"Error processing {filename}: {e}")


if __name__ == '__main__':
    # Input folder containing the Excel files
    folder = r"C:\Users\Z19661\PycharmProjects\LegacyData_\OutputFiles"
    check_and_copy_files(folder)
