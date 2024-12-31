import os

import pandas as pd
from CODES.database.INV_Generator import NumberIncrementer

class DataFrameEnhancer:
    def __init__(self, df_to_enhance,filename):
        self.incrementer = NumberIncrementer()
        self.df_to_enhance = df_to_enhance
        self.filename=os.path.splitext(os.path.basename(filename))[0]

        # Predefined fields and values to be added
        self.fields = [
            "Study Name",
            "Vendor's Name",
            "Start Date of Study Field (mm/dd/yyyy)",
            "End Date of Study Field (mm/dd/yyyy)",
            "Type Of Users",
            "Country of Test",
            "WorkFront Project Number",
            "File Version Number",
            "TCR Research Project Manager",
            "Language"
        ]
        self.values = [
            filename,
            "Spencer Research, Inc.",
            "06/10/2024",
            "06/12/2024",
            "",
            "USA",
            self.incrementer.get_inv_number(filename),
            "1",
            "Laura Valenta da Silva",
            "EN"
        ]

    def enhance_dataframe(self):
        # Create a DataFrame from predefined fields and values
        additional_data = pd.DataFrame({"Field": self.fields, "Value": self.values})

        # Check for missing fields in the original DataFrame
        missing_data = additional_data[~additional_data['Field'].isin(self.df_to_enhance['Field'])]

        # Append only the missing fields to the original DataFrame
        enhanced_df = pd.concat([self.df_to_enhance, missing_data], ignore_index=True)

        return enhanced_df

# Example of using the enhanced DataFrame: