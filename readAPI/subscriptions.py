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

subscriptionextenction = configs.get('subscriptionextenction').data
clienttimezone = configs.get("clienttimezone").data
clientSite = configs.get('clientSite').data
user = configs.get('apikey').data
addonesecond = configs.get('addonesecond').data
datetimeformat = configs.get('datetimeformat').data
cents = configs.get('cents').data
# now we will Create and configure logger
logging.basicConfig(filename=logDir + "/subscription.log",
                    format='[%(asctime)s] %(lineno)d %(levelname)s - %(message)s',
                    filemode='w')
# Let us Create an object
logger = logging.getLogger()

# Now we are going to Set the threshold of logger to DEBUG
logger.setLevel(logging.DEBUG)
outputFile = "_AllSubscriptions.xlsx"


class SubscriptionExecution:
    def getAllSubscriptions(self):
        url = clientSite + subscriptionextenction
        TotalSubscriptionResponse = ReadAPIExecution.getDataFromAPI(self, url, user, logger)
        Subscriptiondictionary = {
            "list": TotalSubscriptionResponse
        }

        with open(jsonDir + '/' + configs.get("clientName").data + "_AllSubscriptions.json", "w") as outfile:
            json.dump(Subscriptiondictionary, outfile)
        logger.info("Final Json File" + str(Subscriptiondictionary))
        try:
            with open(jsonDir + '/' + configs.get("clientName").data + "_AllSubscriptions.json", 'r') as f:
                data = json.loads(f.read())
            # Flatten data
            df_nested_list = pd.json_normalize(data, record_path=['list'])
            df_nested_list_temp = pd.json_normalize(data, record_path=['list'], max_level=1)
            if 'subscription.meta_data' in list(df_nested_list_temp.head()):
                df_nested_list_temp = df_nested_list_temp[['subscription.id', 'subscription.meta_data']]
                df_nested_list = pd.merge(df_nested_list, df_nested_list_temp, how='inner', left_on=['subscription.id'],
                                          right_on=['subscription.id'])
            headers = list(df_nested_list.head())
            newheaders = {}
            for ch in headers:
                newheaders[ch] = ch.replace(".", "_")
            df_nested_list.rename(columns=newheaders, inplace=True)
            df_nested_list.to_excel(excelDir + '/' + configs.get("clientName").data + outputFile, index=False)
            df_splitlineitems = pd.read_excel(excelDir + '/' + configs.get("clientName").data + outputFile)
            if "subscription_subscription_items" in list(df_nested_list.head()):
                dftemp = df_splitlineitems['subscription_subscription_items'].str.split('}, {', -1, expand=True)
                acol = len(dftemp.columns)

                lineitemslist = []
                for i in range(acol):
                    lineitemslist.append('lineitem' + str(i))

                df_splitlineitems[lineitemslist] = df_splitlineitems['subscription_subscription_items'].str.split(
                    '}, {', -1, expand=True)
                for item in lineitemslist:
                    df_splitlineitems[item] = df_splitlineitems[item].str.replace(r'(?s), \'free_quantity.*', '',
                                                                                  regex=True)
                    df_splitlineitems[item] = df_splitlineitems[item].str.replace(r'(?s), \'object.*', '', regex=True)
                    df_splitlineitems[item] = df_splitlineitems[item].str.replace('[{', '', regex=False)
                    df_splitlineitems[item] = df_splitlineitems[item].str.replace('\'', '', regex=False)
                count = 0
                for item in lineitemslist:
                    clist = ['item_price_id[' + str(count) + ']', 'items_type[' + str(count) + ']',
                             'items_quantity[' + str(count) + ']', 'items_unit_price[' + str(count) + ']',
                             'items_amount[' + str(count) + ']']
                    df_splitlineitems['item_price_id[' + str(count) + ']'] = df_splitlineitems[item].str.extract(
                        r"item_price_id: (.*?),")
                    df_splitlineitems['item_type[' + str(count) + ']'] = df_splitlineitems[item].str.extract(
                        r"item_type: (.*?),")
                    df_splitlineitems['item_quantity[' + str(count) + ']'] = df_splitlineitems[item].str.extract(
                        r"quantity: (.*?),")
                    # df_splitlineitems['item_unit_price['+ str(count) + ']'] = df_splitlineitems[item].str.extract(r"unit_price: (.*?),")
                    df_splitlineitems['item_unit_price[' + str(count) + ']'] = df_splitlineitems[item].str.extract(
                        r"unit_price: (\d+)")
                    df_splitlineitems['item_amount[' + str(count) + ']'] = df_splitlineitems[item].str.extract(
                        r"amount: (\d+)")

                    # df_splitlineitems[item] = df_splitlineitems[item].replace("trial_end:(.*?),", "", regex=True)
                    # df_splitlineitems[item] = df_splitlineitems[item].replace("object:(.*?),", "", regex=True)
                    # df_splitlineitems[item] = df_splitlineitems[item].replace("charge_on_event:(.*?),", "", regex=True)
                    # df_splitlineitems[item] = df_splitlineitems[item].replace("charge_once:(.*?)", "", regex=True)
                    # df_splitlineitems[clist] = df_splitlineitems[item].str.split(',', -1, expand=True)
                    # for col in clist:
                    #     df_splitlineitems[col] = df_splitlineitems[col].str.replace(r'.*?: ', '', regex=True)
                    count += 1
                df_splitlineitems.to_excel(excelDir + '/' + configs.get("clientName").data + outputFile, index=False)
            if "subscription_addons" in list(df_nested_list.head()):
                df_splitaddon = self.subscription_addons_split(df_splitlineitems)
                df_splitaddon.to_excel(excelDir + '/' + configs.get("clientName").data + outputFile, index=False)
            tdf = pd.read_excel(excelDir + '/' + configs.get("clientName").data + outputFile)

            dateconvertioncollist = ["subscription_start_date", "subscription_trial_start", "subscription_trial_end",
                                     "subscription_current_term_start", "subscription_current_term_end",
                                     "subscription_next_billing_at", "subscription_created_at",
                                     "subscription_started_at", "subscription_pause_date",
                                     "subscription_activated_at", "subscription_updated_at", "subscription_due_since",
                                     "subscription_cancelled_at",
                                     "customer_created_at", "customer_updated_at", "card_created_at", "card_updated_at",
                                     "subscription_contract_term_contract_start",
                                     "subscription_contract_term_contract_end"]
            for col in tqdm(dateconvertioncollist, desc='dateconvertioncollist'):
                if col in list(tdf.head()):
                    tdf[col] = tdf[col].apply(
                        lambda x: ReadAPIExecution.epoch_To_Datetime_Convert(self, x, clienttimezone) if pd.isna(
                            x) != True else None)
            tdf.to_excel(excelDir + '/' + configs.get("clientName").data + outputFile, index=False)
            # configure cents : False to execute below block
            if cents == 'False':
                logger.info("convertingtocents")
                centsToDoller = pd.read_excel(excelDir + '/' + configs.get("clientName").data + outputFile)
                centsToDollerlist = ["subscription_mrr", "items_unit_price[0]", "items_amount[0]",
                                     "items_unit_price[1]",
                                     "items_amount[1]", "items_unit_price[2]", "items_amount[2]"]
                for col in tqdm(centsToDollerlist, desc='centsToDollerlist'):
                    if col in list(centsToDoller.head()):
                        centsToDoller[col] = centsToDoller[col].div(100)
                centsToDoller.to_excel(excelDir + '/' + configs.get("clientName").data + outputFile, index=False)
        except Exception as e:
            logger.info(e)
            logger.error("exception in subscriptions:" + str(e))
            logger.exception(e)

    def subscription_addons_split(self, dfdata):
        df = dfdata[["subscription_id", "subscription_addons"]]
        dfs = pd.DataFrame
        dfl = pd.DataFrame
        df['subscription_addons'] = df['subscription_addons'].replace("'", '"', regex=True)
        df['subscription_addons'] = df['subscription_addons'].replace(": False,", ': "False",', regex=True)
        df['subscription_addons'] = df['subscription_addons'].replace(": True,", ': "True",', regex=True)
        for i, j in tqdm(zip(df['subscription_id'], df['subscription_addons']), total=len(df['subscription_id']),
                         desc='subscription_addons'):
            logger.info("splitting for '{}' subscription_id and the date is :{}".format(i, j))
            if not pd.isna(j) and j != '[]':
                data = json.loads(j)
                for k in range(len(data)):
                    prefix = "addon_"
                    dfli = pd.json_normalize(data[k])
                    sufix = "[" + str(k) + "]"
                    headers = list(dfli.head())
                    newheaders = {}
                    for ch in headers:
                        newheaders[ch] = prefix + ch + sufix
                    dfli.rename(columns=newheaders, inplace=True)
                    dfli['subscription_id'] = [i]
                    if k == 0:
                        dfl = dfli
                    else:
                        dfl = pd.merge(dfl, dfli, left_on="subscription_id", right_on="subscription_id", how='inner')
                        # dfl = dfl.append(dfli)
            else:
                wdata = {'subscription_id': [i]}
                dfl = pd.DataFrame(wdata)
            try:
                dfs = dfs.append(dfl)
            except:
                dfs = dfl
        dfaddon = pd.merge(dfdata, dfs, how='inner', left_on=['subscription_id'], right_on=['subscription_id'])
        return dfaddon


subscriptionobj = SubscriptionExecution()

subscriptionobj.getAllSubscriptions()
