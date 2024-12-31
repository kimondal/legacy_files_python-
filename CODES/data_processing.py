import os
import pandas as pd
from datetime import datetime

from CODES.Fenix_files.TCR_Fenix import Fenix_ETL
from CODES.MOT_files.TCR_MOT import MOT_ETL
from CODES.NA_Juice_Files.TCR_NaJuice import NAJuiceETL
from CODES.ETL import ETL
from CODES.Supporting_code.TCR_EMEA_NDB import EMEA_NBD_ETL
from CODES.EMEA.TCR_EMEA import EMEA_ETL
from pathlib import Path

# Register the ETL classes
ETL.register_class("NAJuiceETL", NAJuiceETL)
ETL.register_class("EMEA_ETL", EMEA_ETL)
ETL.register_class("EMEA_NBD_ETL", EMEA_NBD_ETL)
ETL.register_class("Fenix", Fenix_ETL)
ETL.register_class("MOT", MOT_ETL)



class ETLProcessor:

    @staticmethod
    def find_files_in_directory(directory, file_extension='.xlsb'):
        """Find all files with a specific extension in the given directory."""
        file_data = []  # List to store tuples of (file_name, file_path)
        for root, dirs, files in os.walk(directory):
            for file in files:
                file_path = os.path.join(root, file)
                file_data.append((file, file_path))  # Append as a tuple (file_name, file_path)
        return file_data

    @staticmethod
    def process_file(filepath, etl_class_name, output_folder):
        """Process a single file using the given ETL class and save the result to Excel."""
        try:
            # Extract filename
            filename = os.path.basename(filepath)
            # filename = Path(filepath).name
            filename_without_extension = Path(filepath).stem


            print(f"Processing file: {filename_without_extension} using {etl_class_name}...")

            etl = ETL()
            etl.initializer(etl_class_name, filepath=filepath)
            output_filename = f"tcr_{filename_without_extension}.xlsx"
            output_filepath = os.path.join(output_folder, output_filename)
            etl.ConvertToExcel(output_filepath)

            return {
                "filename": filename,
                "status": "OK"
            }
        except KeyError as e:
            return {
                "filename": os.path.basename(filepath),
                "status": f"Not OK - Missing column: {str(e)}"
            }
        except AttributeError as e:
            return {
                "filename": os.path.basename(filepath),
                "status": f"Not OK - Data object not initialized: {str(e)}"
            }
        except Exception as e:
            return {
                "filename": os.path.basename(filepath),
                "status": f"Not OK - {str(e)}"
            }

    @classmethod
    def process_directory(cls, input_directory, output_directory, etl_class_name, all_results):
        """Process all files in the directory using the given ETL class and append results to all_results."""
        # Find all .xlsb files in the directory
        file_data = cls.find_files_in_directory(input_directory)

        # Process each file
        for filename, filepath in file_data:
            result = cls.process_file(filepath, etl_class_name, output_directory)
            all_results.append(result)

    @classmethod
    def process_all_directories(cls, directories, output_directory):
        """Process files from all given directories and append results to a single CSV."""
        all_results = []

        # Ensure the output directory exists
        if not os.path.exists(output_directory):
            os.makedirs(output_directory)
            print(f"Output directory created: {output_directory}")

        # Process each directory
        for directory, etl_class_name in directories:
            print(f"Processing files from {directory} using {etl_class_name}...")
            cls.process_directory(directory, output_directory, etl_class_name, all_results)

        # Convert the results to a DataFrame
        status_df = pd.DataFrame(all_results)

        # Prepare the data in the required format (Date only)
        status_df['date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # Use only the date, column name is 'date'
        status_df['file_name'] = status_df['filename']  # File name without path already extracted
        status_df = status_df[['date', 'file_name', 'status']]  # Reorder columns to match format

        # Ensure the CSV file exists and has the correct headers
        status_file = os.path.join(output_directory, 'logger.csv')
        if os.path.exists(status_file):
            # If the file exists, append new status data
            status_df.to_csv(status_file, mode='a', header=False, index=False)
            print("Appended new status data to existing logger.csv.")
        else:
            # If the file doesn't exist, create a new file with headers
            status_df.to_csv(status_file, index=False)


        print(f"All processing completed. Status saved to logger.csv")


# Main script for processing
if __name__ == "__main__":
    # Define directories (input and output folders)
    na_juice_directory = r"C:\Users\Z19661\PycharmProjects\LegacyData_\InputFiles\NaJuice"
    emea_directory = r"C:\Users\Z19661\PycharmProjects\LegacyData_\InputFiles\EMEA_ETL"
    emea_nbd_directory = r"C:\Users\Z19661\PycharmProjects\LegacyData_\InputFiles\EMEA_NDB_ETL"
    fenix_directory = r"C:\Users\Z19661\PycharmProjects\LegacyData_\InputFiles\fenix"
    MOT_directory = r"C:\Users\Z19661\PycharmProjects\LegacyData_\InputFiles\MOT"
    output_directory = r"C:\Users\Z19661\PycharmProjects\LegacyData_\OutputFiles"

    # List of directories to process along with their respective ETL classes
    directories_to_process = [
        (na_juice_directory, "NAJuiceETL"),
        (emea_directory, "EMEA_ETL"),
        (emea_nbd_directory, "EMEA_NBD_ETL"),
        (fenix_directory, "Fenix"),
        (MOT_directory, "MOT")
    ]

    # Process files from all directories and save the status to a single CSV
    ETLProcessor.process_all_directories(directories_to_process, output_directory)

    print("--------------------------------------------------------------------------------------")
    print("[STATUS] All files processed successfully. Status has been saved to process_data.csv.")
    print("--------------------------------------------------------------------------------------")
