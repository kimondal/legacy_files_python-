import os

import pandas as pd
from CODES.database.INV_Generator import NumberIncrementer



class DataFrameEnhancerMOT:
    def __init__(self, key_decoder,filename,country_of_test,startdate,enddate):
        self.incrementer = NumberIncrementer()
        self.df_to_enhance = key_decoder
        self.filename=os.path.splitext(os.path.basename(filename))[0]
        if self.df_to_enhance is None:
            self.df_to_enhance = pd.DataFrame(columns=["Field", "Value"])

        # Predefined fields and values to be added
        self.fields = [
            "Study Name",
            "Vendor's Name",
            "Start Date of Study Field (mm/dd/yyyy)",
            "End Date of Study Field (mm/dd/yyyy)",
            "Country of Test",
            "WorkFront Project Number",
            "File Version Number",
            "TCR Research Project Manager",
            "Language"
        ]
        self.values = [
            self.filename,
            "Ipsos",
            startdate,
            enddate,
            country_of_test,
            self.incrementer.get_inv_number(filename),
            "1",
            "Rui Xiong",
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