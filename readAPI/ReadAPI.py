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
# now we will Create and configure logger
logging.basicConfig(filename=os.getcwd() + "/ReadAPI.log",
                    format='[%(asctime)s] %(lineno)d %(levelname)s - %(message)s',
                    filemode='w')
# Let us Create an object
logger = logging.getLogger()

# Now we are going to Set the threshold of logger to DEBUG
logger.setLevel(logging.DEBUG)
# customerid = ["PTA300", "16CGhWT6Vri6qMkhc", "PSA003", "PSA004", "PSA006", "PSA001", "PSA005"]
customerid = ["PTA300", "16CGhWT6Vri6qMkhc", "PSA003", "PSA004", "PSA006", "PSA001", "PSA005"]

class ReadAPIExecution:
    def epoch_To_Datetime_Convert(self, epochtimestamp, timezoneOfCustomer):
        my_datetime = datetime.fromtimestamp(epochtimestamp, tz=pytz.timezone(timezoneOfCustomer))
        modified = my_datetime.strftime(datetimeformat)
        if modified.endswith(':59') and addonesecond == "True":
            expected_date = datetime.strptime(modified, '%Y-%m-%d %H:%M:%S')
            expected_date += timedelta(seconds=1)
            modified = expected_date.strftime(datetimeformat)
        return modified

    def getDataFromAPI(self, url, user):
        page = 0

        def callAPI(url, user, cusid):
            resp = requests.get(url + '?customer_id[is]=' + cusid, auth=HTTPBasicAuth(user, user))
            jsonresp = json.loads(resp.text)
            return jsonresp

        for cusid in customerid:
            jsonpathresponse = callAPI(url, user, cusid)
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
        return TotalResponse



