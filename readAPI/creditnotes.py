import datetime
import json
import logging
import os
import sys
from datetime import datetime

import jsonpath
import pandas as pd
import pytz
import requests
from jproperties import Properties
from requests.auth import HTTPBasicAuth
from datetime import datetime, timedelta

from readAPI.ReadAPI import ReadAPIExecution

configs = Properties()
home_folder = os.getenv('HOME')
ROOT_DIR1 = os.path.abspath(os.curdir)
ROOT_DIR = ROOT_DIR1.replace('readAPI', 'configuration.properties')
jsonDir = ROOT_DIR1 + '/jsonfiles'
with open(ROOT_DIR, 'rb') as config_file:
    configs.load(config_file)


creditsexecution = configs.get('creditsexecution').data

clienttimezone = configs.get("clienttimezone").data
clientSite = configs.get('clientSite').data
user = configs.get('apikey').data

addonesecond = configs.get('addonesecond').data
datetimeformat = configs.get('datetimeformat').data
cents = configs.get('cents').data
# now we will Create and configure logger
logging.basicConfig(filename=os.getcwd() + "/creditnote.log",
                    format='[%(asctime)s] %(lineno)d %(levelname)s - %(message)s',
                    filemode='w')
# Let us Create an object
logger = logging.getLogger()

# Now we are going to Set the threshold of logger to DEBUG
logger.setLevel(logging.DEBUG)


class CreditnoteExecution:
    def getAllCreditNotes(self):
        url = clientSite + creditsexecution
        TotalCreditNotesResponse = ReadAPIExecution.getDataFromAPI(self, url, user)
        CreditNotesdictionary = {
            "list": TotalCreditNotesResponse
        }

        with open(jsonDir + '/' + configs.get("clientName").data + "_AllCreditNotes.json", "w") as outfile:
            json.dump(CreditNotesdictionary, outfile)
        logger.info("Final Json File" + str(CreditNotesdictionary))
        try:
            with open(jsonDir + '/' + configs.get("clientName").data + "_AllCreditNotes.json", 'r') as f:
                data = json.loads(f.read())
            # Flatten data
            df_nested_list = pd.json_normalize(data, record_path=['list'])
            headers = list(df_nested_list.head())
            newheaders = {}
            for ch in headers:
                newheaders[ch] = ch.replace(".", "_")
            df_nested_list.rename(columns=newheaders, inplace=True)
            df_nested_list.to_excel(configs.get("clientName").data + "_AllCreditNotes.xlsx", index=False)
            tdf = pd.read_excel(configs.get("clientName").data + "_AllCreditNotes.xlsx")
            dateconvertioncollist = ["credit_note_date", "credit_note_generated_at", "credit_note_updated_at", "credit_note_refunded_at"]
            for col in dateconvertioncollist:
                if col in list(tdf.head()):
                    tdf[col] = tdf[col].apply(
                        lambda x: ReadAPIExecution.epoch_To_Datetime_Convert(x, clienttimezone) if pd.isna(x) != True else None)
            tdf.to_excel(configs.get("clientName").data + "_AllCreditNotes.xlsx", index=False)
        except Exception as e:
            logger.error("exception in credit_notes:" + str(e))

creditnoteobj = CreditnoteExecution()
creditnoteobj.getAllCreditNotes()
