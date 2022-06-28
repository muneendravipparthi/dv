import datetime
import json
import logging
import os
import sys
from datetime import datetime
from timeit import default_timer as timer
import jsonpath
import pandas as pd
import pytz
import requests
from jproperties import Properties
from requests.auth import HTTPBasicAuth
from datetime import datetime, timedelta

configs = Properties()
home_folder = os.getenv('HOME')
ROOT_DIR1 = os.path.abspath(os.curdir)
ROOT_DIR = ROOT_DIR1.replace('readAPI', 'configuration.properties')
jsonDir = ROOT_DIR1 + '/jsonfiles'
with open(ROOT_DIR, 'rb') as config_file:
    configs.load(config_file)

customerextenction = configs.get('customerextenction').data
subscriptionextenction = configs.get('subscriptionextenction').data
invoiceextenction = configs.get('invoiceextenction').data
creditsexecution = configs.get('creditsexecution').data
transactionsexecution = configs.get('transactionsexecution').data
clienttimezone = configs.get("clienttimezone").data
clientSite = configs.get('clientSite').data
user = configs.get('apikey').data
executionmode = configs.get('executionmode').data
addonesecond = configs.get('addonesecond').data
datetimeformat = configs.get('datetimeformat').data
cents = configs.get('cents').data

class ReadAPIExecution:
    def epoch_To_Datetime_Convert(self, epochtimestamp, timezoneOfCustomer):
        my_datetime = datetime.fromtimestamp(epochtimestamp, tz=pytz.timezone(timezoneOfCustomer))
        # my_datetime = datetime.fromtimestamp(epochtimestamp, tz=pytz.timezone("UTC"))
        # my_datetime = datetime.fromtimestamp(epochtimestamp)
        modified = my_datetime.strftime(datetimeformat)
        if modified.endswith(':59') and addonesecond == "True":
            expected_date = datetime.strptime(modified, '%Y-%m-%d %H:%M:%S')
            expected_date += timedelta(seconds=1)
            modified = expected_date.strftime(datetimeformat)
        return modified

    def getDataFromAPI(self, url, user, logger):
        page = 0
        offsetData = ''
        # offsetData = '["1541245200000","229501294"]'

        def callAPI(url, user, offsetValue):
            print(page)
            if (offsetValue != ''):
                resp = requests.get(url + '&offset=' + offsetValue, auth=HTTPBasicAuth(user, user))
                jsonresp = json.loads(resp.text)
                # with open("/Users/cb-muneendra/cb_data_validation/readAPI/pixelluJson/pixellu_b2_"+ str(timer()) +"_Invoices.json", "w") as responseFile:
                #     json.dump(jsonresp, responseFile)
                return jsonresp
            else:
                resp = requests.get(url, auth=HTTPBasicAuth(user, user))
                jsonresp = json.loads(resp.text)
                # with open("/Users/cb-muneendra/cb_data_validation/readAPI/pixelluJson/pixellu_b2_"+ str(timer()) +"_Invoices.json", "w") as responseFile:
                #     json.dump(jsonresp, responseFile)
                return jsonresp

        while True:
            jsonpathresponse = callAPI(url, user, offsetData)
            isOffsetPresent = "next_offset" in jsonpathresponse
            if page == 0:
                TotalResponse = jsonpathresponse['list']
            if page != 0:
                jsonresp1ListElement = jsonpathresponse['list']
                l = jsonpath.jsonpath(jsonpathresponse, "list")
                totalRecordCountInResp = len(l[0])
                for i in range(0, totalRecordCountInResp):
                    try:
                        TotalResponse.append(jsonresp1ListElement[i])
                    except Exception as e:
                        logger.info("element missing in list" + str(e))
            page += 1
            if (isOffsetPresent):
                offsetValue = jsonpath.jsonpath(jsonpathresponse, "next_offset")
                offsetValue[0].replace('"', '\\"')
                offsetData = offsetValue[0]
                logger.info("offset value : {}".format(offsetData))
                if(executionmode == "DEBUG"):
                    logger.info("Executed in DEBUG Mode")
                    print("Executed in DEBUG Mode")
                    break
            else:
                break
        logger.info('Execution completed. and Total no of pages :' + str(page))
        return TotalResponse


# obj = ReadAPIExecution()
# if len(sys.argv) >= 2:
#     for arg in range(1, len(sys.argv)):
#         logger.info('=======Executing for :' + str(sys.argv[arg]) + '=======')
#         print('=======Executing for :', sys.argv[arg], '=======')
#         if sys.argv[arg] == 'c':
#             obj.getAllCustomers()
#         elif sys.argv[arg] == 's':
#             obj.getAllSubscriptions()
#         elif sys.argv[arg] == 'i':
#             obj.getAllInvoices()
#         elif sys.argv[arg] == 'cn':
#             obj.getAllCreditNotes()
#         elif sys.argv[arg] == 't':
#             obj.getAllTransactions()
# else:
#     # obj.getAllCustomers()
#     # obj.getAllSubscriptions()
#     # obj.getAllInvoices()
#     # obj.getAllCreditNotes()
#     # obj.getAllTransactions()
