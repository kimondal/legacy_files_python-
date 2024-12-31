import pandas as pd
from CODES.NA_Juice_Files.TCR_NaJuice import NAJuiceETL
import logging


class ETL:
    dataObject = None
    class_registry = {}

    # Set up logger (Only once)
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.DEBUG)
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s - %(message)s')
    ch.setFormatter(formatter)
    logger.addHandler(ch)

    @classmethod
    def register_class(cls, class_name, class_type):
        """Registers a class to the class registry"""
        try:
            if class_name not in cls.class_registry:
                cls.class_registry[class_name] = class_type
                cls.logger.info(f"Class {class_name} registered successfully.")
            else:
                cls.logger.warning(f"Class {class_name} is already registered.")
        except Exception as e:
            cls.logger.error(f"Error while registering class {class_name}: {str(e)}")
            raise  # Reraise the exception

    def initializer(self, class_name, *args, **kwargs):
        """Initialize the class object based on the class name"""
        try:
            if class_name in self.class_registry:
                self.dataObject = self.class_registry[class_name](*args, **kwargs)

                if isinstance(self.dataObject, NAJuiceETL):
                    self.dataObject.RawData(sheet_name='Raw Data')
                    self.dataObject.KeyDecoder(sheet_name='Key_Decoder')
                    self.dataObject.Consumer(sheet_name='Consumer Info')
                    self.dataObject.Project(sheet_name='Additional data for NDB')
                    self.dataObject.Product()
            else:
                self.logger.warning(f"Class {class_name} not found in registry!")
                raise ValueError(f"Class {class_name} not found in registry!")
        except Exception as e:
            self.logger.error(f"Error during initialization: {str(e)}")
            self.dataObject = None
            raise  # Reraise the exception

    def _get_data(self, data_type):
        """Helper function to fetch data from dataObject based on the type"""
        try:
            if self.dataObject:
                return getattr(self.dataObject, data_type)  # Dynamically access the data attribute
            else:
                raise ValueError("dataObject is not initialized.")
        except Exception as e:
            self.logger.error(f"Error fetching {data_type}: {str(e)}")
            raise  # Reraise the exception

    def getRaw(self):
        """Fetch raw data from the dataObject"""
        return self._get_data('raw_data')

    def getConsumer(self):
        """Fetch consumer data from the dataObject"""
        return self._get_data('consumer_data')

    def getKeyDecoder(self):
        """Fetch key decoder data from the dataObject"""
        return self._get_data('key_decoder')

    def getProduct(self):
        """Fetch product data from the dataObject"""
        return self._get_data('product_data')

    def getProject(self):
        """Fetch project data from the dataObject"""
        return self._get_data('project_data')

    def ConvertToExcel(self, output_filepath):
        """Convert all processed data to an Excel file"""
        try:
            # Retrieve your DataFrames
            raw = self.getRaw()
            consumer = self.getConsumer()
            keydecoder = self.getKeyDecoder()
            product = self.getProduct()
            project = self.getProject()

            # Use pd.ExcelWriter to save the DataFrames to different sheets
            with pd.ExcelWriter(output_filepath, engine='xlsxwriter') as writer:
                if raw is not None:
                    raw.to_excel(writer, sheet_name='TCR_Raw', index=False)
                if consumer is not None:
                    consumer.to_excel(writer, sheet_name='TCR_Consumer', index=False)
                if keydecoder is not None:
                    keydecoder.to_excel(writer, sheet_name='TCR_Key_Decoder', index=False)
                if project is not None:
                    project.to_excel(writer, sheet_name='TCR_Project', index=False)
                if product is not None:
                    product.to_excel(writer, sheet_name='TCR_Product', index=False)

            self.logger.info("Data successfully saved to Excel.")

        except Exception as e:
            self.logger.error(f"Error converting to Excel: {str(e)}")
            raise  # Reraise the exception
