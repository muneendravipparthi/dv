import json
import logging
import os

import pandas as pd
from jproperties import Properties
from tqdm import tqdm

from readAPI.ReadAPI import ReadAPIExecution

configs = Properties()
home_folder = os.getenv('HOME')
ROOT_DIR1 = os.path.abspath(os.curdir)
ROOT_DIR = ROOT_DIR1.replace('readAPI', 'configuration.properties')
jsonDir = ROOT_DIR1 + '/jsonfiles'
excelDir = ROOT_DIR1 + '/ds3files'
logDir = ROOT_DIR1 + '/logfiles'
with open(ROOT_DIR, 'rb') as config_file:
    configs.load(config_file)

customerextenction = configs.get('customerextenction').data
clienttimezone = configs.get("clienttimezone").data
clientSite = configs.get('clientSite').data
user = configs.get('apikey').data
executionmode = configs.get('executionmode').data
addonesecond = configs.get('addonesecond').data
datetimeformat = configs.get('datetimeformat').data
cents = configs.get('cents').data
# now we will Create and configure logger
logging.basicConfig(filename=logDir + "/customer.log",
                    format='[%(asctime)s] %(lineno)d %(levelname)s - %(message)s',
                    filemode='w')
# Let us Create an object
logger = logging.getLogger()

# Now we are going to Set the threshold of logger to DEBUG
logger.setLevel(logging.DEBUG)
outputFile = "_DS3_AllCustomers.xlsx"
RequireAPIExecution = True


class CustomerExecution:

    def getAllCustomers(self):
        if RequireAPIExecution:
            logger.info("Executiong Extraction of Customers")
            url = clientSite + customerextenction
            newjson = ''
            TotalCustomerResponse = ReadAPIExecution.getDataFromAPI(self, url, user, logger)
            Customerdictionary = {
                "list": TotalCustomerResponse
            }
            with open(jsonDir + '/' + configs.get("clientName").data + "_AllCustomers.json", "w") as outfile:
                json.dump(Customerdictionary, outfile)
            logger.info("Final Json File" + str(Customerdictionary))
        try:
            with open(jsonDir + '/' + configs.get("clientName").data + "_AllCustomers.json", 'r') as f:
                data = json.loads(f.read())
            # Flatten data
            df_nested_list = pd.json_normalize(data, record_path=['list'])
            df_nested_list_temp = pd.json_normalize(data, record_path=['list'], max_level=1)
            if 'customer.meta_data' in list(df_nested_list_temp.head()):
                df_nested_list_temp = df_nested_list_temp[['customer.id', 'customer.meta_data']]
                df_nested_list = pd.merge(df_nested_list, df_nested_list_temp, how='inner', left_on=['customer.id'],
                                          right_on=['customer.id'])
            headers = list(df_nested_list.head())
            newheaders = {}
            for ch in headers:
                newheaders[ch] = ch.replace(".", "_")
            df_nested_list.rename(columns=newheaders, inplace=True)
            df_nested_list.to_excel(excelDir + '/' + configs.get("clientName").data + outputFile, index=False)

            tdf = pd.read_excel(excelDir + '/' + configs.get("clientName").data + outputFile)
            emaillist = ['customer_email', 'customer_billing_address_email']
            for col in emaillist:
                if col in list(tdf.head()):
                    tdf[col] = tdf[col].replace('_AT_', '@', regex=True)
                    tdf[col] = tdf[col].replace('@example.com', '', regex=True)
            dateconvertioncollist = ["customer_created_at", "customer_updated_at", "card_created_at", "card_updated_at"]
            for col in tqdm(dateconvertioncollist, desc='dateconvertioncollist'):
                if col in list(tdf.head()):
                    tdf[col] = tdf[col].apply(
                        lambda x: ReadAPIExecution.epoch_To_Datetime_Convert(self, x, clienttimezone) if pd.isna(
                            x) != True else None)
            tdf.to_excel(excelDir + '/' + configs.get("clientName").data + outputFile, index=False)
            logger.info("Completed data conversion from Json to Excel")
        except Exception as e:
            logger.error("exception in customers:" + str(e))
            logger.info("Something failed during data conversion from Json to Excel")
            logger.exception(e)


customerobj = CustomerExecution()
customerobj.getAllCustomers()
