import json
import logging
import os

import pandas as pd
from jproperties import Properties

from readAPI.ReadAPI import ReadAPIExecution

configs = Properties()
home_folder = os.getenv('HOME')
ROOT_DIR1 = os.path.abspath(os.curdir)
ROOT_DIR = ROOT_DIR1.replace('readAPI', 'configuration.properties')
jsonDir = ROOT_DIR1 + '/jsonfiles'
with open(ROOT_DIR, 'rb') as config_file:
    configs.load(config_file)

invoiceextenction = configs.get('invoiceextenction').data
clienttimezone = configs.get("clienttimezone").data
clientSite = configs.get('clientSite').data
user = configs.get('apikey').data
addonesecond = configs.get('addonesecond').data
datetimeformat = configs.get('datetimeformat').data
cents = configs.get('cents').data
# now we will Create and configure logger
logging.basicConfig(filename=os.getcwd() + "/invoice.log",
                    format='[%(asctime)s] %(lineno)d %(levelname)s - %(message)s',
                    filemode='w')
# Let us Create an object
logger = logging.getLogger()

# Now we are going to Set the threshold of logger to DEBUG
logger.setLevel(logging.DEBUG)

class InvoiceExecution:

    def getAllInvoices(self):
        logger.info("Executiong Extraction of Invoices")
        url = clientSite + invoiceextenction
        TotalInvoicesResponse = ReadAPIExecution.getDataFromAPI(self, url, user)
        Invoicedictionary = {
            "list": TotalInvoicesResponse
        }

        with open(jsonDir + '/' + configs.get("clientName").data + "_AllInvoices.json", "w") as outfile:
            json.dump(Invoicedictionary, outfile)
        logger.info("Final Json File" + str(Invoicedictionary))
        logger.info("Execution Completed for Extraction of Invoices")
        try:
            logger.info("Converting Json data to Excel data Initiated")
            with open(jsonDir + '/' + configs.get("clientName").data + "_AllInvoices.json", 'r') as f:
                data = json.loads(f.read())
            # Flatten data
            df_nested_list = pd.json_normalize(data, record_path=['list'])
            headers = list(df_nested_list.head())
            newheaders = {}
            for ch in headers:
                newheaders[ch] = ch.replace(".", "_")
            df_nested_list.rename(columns=newheaders, inplace=True)
            df_nested_list.to_excel(configs.get("clientName").data + "_AllInvoices.xlsx", index=False)
            if "invoice_line_items" in list(df_nested_list.head()):
                df_splitlineitems = pd.read_excel(configs.get("clientName").data + "_AllInvoices.xlsx")
                df_splitlineitems = self.invoice_lineitem_split(df_splitlineitems)
                df_splitlineitems.to_excel(configs.get("clientName").data + "_AllInvoices.xlsx", index=False)
            if "invoice_line_item_taxes" in list(df_nested_list.head()):
                df_splitlineitemtaxes = pd.read_excel(configs.get("clientName").data + "_AllInvoices.xlsx")
                df_splitlineitemtaxes = self.invoice_lineitemtaxes_split(df_splitlineitemtaxes)
                df_splitlineitemtaxes.to_excel(configs.get("clientName").data + "_AllInvoices.xlsx", index=False)
            if "invoice_linked_payments" in list(df_nested_list.head()):
                df_splitpayments = pd.read_excel(configs.get("clientName").data + "_AllInvoices.xlsx")
                df_splitpayments = self.invoice_payment_split(df_splitpayments)
                df_splitpayments.to_excel(configs.get("clientName").data + "_AllInvoices.xlsx", index=False)
            if "invoice_line_item_discounts" in list(df_nested_list.head()):
                df_splitdiscounts = pd.read_excel(configs.get("clientName").data + "_AllInvoices.xlsx")
                df_splitdiscounts = self.invoice_discount_split(df_splitdiscounts)
                df_splitdiscounts.to_excel(configs.get("clientName").data + "_AllInvoices.xlsx", index=False)
            tdf = pd.read_excel(configs.get("clientName").data + "_AllInvoices.xlsx")
            dateconvertioncollist = ["invoice_date", "invoice_due_date", "invoice_paid_at", "invoice_updated_at",
                                     "invoice_generated_at", "line_item_date_from[0]", "line_item_date_to[0]",
                                     "line_item_date_from[1]", "line_item_date_to[1]", "line_item_date_from[2]",
                                     "line_item_date_to[2]", "line_item_date_from[3]", "line_item_date_to[3]",
                                     "line_item_date_from[4]", "line_item_date_to[4]", "line_item_date_from[5]",
                                     "line_item_date_to[5]", "line_item_date_from[6]", "line_item_date_to[6]",
                                     "line_item_date_from[7]", "line_item_date_to[7]", "line_item_date_from[8]",
                                     "line_item_date_to[8]","line_item_date_from[9]",
                                     "line_item_date_to[9]","line_item_date_from[10]",
                                     "line_item_date_to[10]", "applied_at", "txn_date", "payments_txn_date[0]", "payments_txn_date[1]"]
            for col in dateconvertioncollist:
                if col in list(tdf.head()):
                    tdf[col] = tdf[col].apply(
                        lambda x: ReadAPIExecution.epoch_To_Datetime_Convert(self, x, clienttimezone) if pd.isna(x) != True else None)
            tdf.to_excel(configs.get("clientName").data + "_AllInvoices.xlsx", index=False)
            if cents == 'False':
                centsToDoller = pd.read_excel(configs.get("clientName").data + "_AllInvoices.xlsx")
                centsToDollerlist = ["invoice_total", "invoice_amount_paid", "invoice_amount_adjusted",
                                     "invoice_write_off_amount", "invoice_credits_applied", "invoice_amount_due",
                                     "invoice_amount_to_collect", "invoice_new_sales_amount", "line_item_unit_amount[0]",
                                     "line_item_amount[0]", "line_item_tax_amount[0]", "line_item_unit_amount[1]",
                                     "line_item_amount[1]", "line_item_tax_amount[1]", "line_item_unit_amount[2]",
                                     "line_item_amount[2]", "line_item_tax_amount[2]", "line_item_unit_amount[3]",
                                     "line_item_amount[3]", "line_item_tax_amount[3]", "line_item_unit_amount[4]",
                                     "line_item_amount[4]", "line_item_tax_amount[4]", "line_item_unit_amount[5]",
                                     "line_item_amount[5]", "line_item_tax_amount[5]", "line_item_unit_amount[6]",
                                     "line_item_amount[6]", "line_item_tax_amount[6]", "line_item_unit_amount[7]",
                                     "line_item_amount[7]", "line_item_tax_amount[7]", "applied_amount", "txn_amount", "payments_txn_amount[0]"]
                for col in centsToDollerlist:
                    if col in list(tdf.head()):
                        centsToDoller[col] = centsToDoller[col].div(100)
                centsToDoller.to_excel(configs.get("clientName").data + "_AllInvoices.xlsx", index=False)
                logger.info("Converting Json data to Excel data Completed")
        except Exception as e:
            logger.info("Something failed during data convertion from Json to Excel")
            logger.error("exception in invoices:" + str(e))
            logger.exception(e)

    def invoice_lineitem_split(self,dfdata):
        df = dfdata[["invoice_id", "invoice_line_items"]]
        dfs = pd.DataFrame
        dfl = pd.DataFrame
        df['invoice_line_items'] = df['invoice_line_items'].replace("Tiina's addon", "Tiina^s addon", regex=True)
        df['invoice_line_items'] = df['invoice_line_items'].replace("'", '"', regex=True)
        df['invoice_line_items'] = df['invoice_line_items'].replace(": False,", ': "False",', regex=True)
        df['invoice_line_items'] = df['invoice_line_items'].replace(": True,", ': "True",', regex=True)
        for i, j in zip(df['invoice_id'], df['invoice_line_items']):
            print("splitting for '{}' invoice id and the date is :{}".format(i,j))
            # if "INVFI026771" in i:
            #     print("getting isssue here after")
            if not pd.isna(j):
                data = json.loads(j)
                for k in range(len(data)):
                    prefix = "line_item_"
                    dfli = pd.json_normalize(data[k])
                    sufix = "[" + str(k) + "]"
                    headers = list(dfli.head())
                    newheaders = {}
                    for ch in headers:
                        newheaders[ch] = prefix + ch + sufix
                    dfli.rename(columns=newheaders, inplace=True)
                    dfli['invoice_id'] = [i]
                    if k == 0:
                        # dfli['invoice_id'] = [i]
                        dfl = dfli
                    else:
                        dfl = pd.merge(dfl, dfli, left_on= "invoice_id",right_on="invoice_id", how='inner')
                        # dfl = dfl.append(dfli)
            else:
                wdata = {'invoice_id': [i]}
                dfl = pd.DataFrame(wdata)
            try:
                dfs = dfs.append(dfl)
            except:
                dfs = dfl
        dflineitem = pd.merge(dfdata, dfs, how='inner', left_on=['invoice_id'],
                         right_on=['invoice_id'])
        return dflineitem

    def invoice_lineitemtaxes_split(self,dfdata):
        df = dfdata[["invoice_id", "invoice_line_item_taxes"]]
        dfs = pd.DataFrame
        dfl = pd.DataFrame
        df['invoice_line_item_taxes'] = df['invoice_line_item_taxes'].replace("'", '"', regex=True)
        df['invoice_line_item_taxes'] = df['invoice_line_item_taxes'].replace(": False,", ': "False",', regex=True)
        df['invoice_line_item_taxes'] = df['invoice_line_item_taxes'].replace(": False}", ': "False"}', regex=True)
        df['invoice_line_item_taxes'] = df['invoice_line_item_taxes'].replace(": True}", ': "True"}', regex=True)
        df['invoice_line_item_taxes'] = df['invoice_line_item_taxes'].replace(": True,", ': "True",', regex=True)
        for i, j in zip(df['invoice_id'], df['invoice_line_item_taxes']):
            print("splitting for '{}' invoice id and the date is :{}".format(i, j))
            if not pd.isna(j):
                data = json.loads(j)
                for k in range(len(data)):
                    prefix = "line_item_taxes_"
                    dfli = pd.json_normalize(data[k])
                    sufix = "[" + str(k) + "]"
                    headers = list(dfli.head())
                    newheaders = {}
                    for ch in headers:
                        newheaders[ch] = prefix + ch + sufix
                    dfli.rename(columns=newheaders, inplace=True)
                    dfli['invoice_id'] = [i]
                    if k == 0:
                        dfl = dfli
                    else:
                        dfl = pd.merge(dfl, dfli, left_on="invoice_id", right_on="invoice_id", how='inner')
                        # dfl = dfl.append(dfli)
            else:
                wdata = {'invoice_id': [i]}
                dfl = pd.DataFrame(wdata)
            try:
                dfs = dfs.append(dfl)
            except:
                dfs = dfl
        dflineitemtax = pd.merge(dfdata, dfs, how='inner', left_on=['invoice_id'], right_on=['invoice_id'])
        return dflineitemtax

    def invoice_payment_split(self, dfdata):
        df = dfdata[["invoice_id", "invoice_linked_payments"]]
        dfs = pd.DataFrame
        dfl = pd.DataFrame
        df['invoice_linked_payments'] = df['invoice_linked_payments'].replace("'", '"', regex=True)
        df['invoice_linked_payments'] = df['invoice_linked_payments'].replace(": False,", ': "False",', regex=True)
        df['invoice_linked_payments'] = df['invoice_linked_payments'].replace(": True,", ': "True",', regex=True)
        for i, j in zip(df['invoice_id'], df['invoice_linked_payments']):
            print("splitting for '{}' invoice id and the date is :{}".format(i, j))
            if not pd.isna(j) and j != '[]':
                data = json.loads(j)
                for k in range(len(data)):
                    prefix = "payments_"
                    dfli = pd.json_normalize(data[k])
                    sufix = "[" + str(k) + "]"
                    headers = list(dfli.head())
                    newheaders = {}
                    for ch in headers:
                        newheaders[ch] = prefix + ch + sufix
                    dfli.rename(columns=newheaders, inplace=True)
                    dfli['invoice_id'] = [i]
                    if k == 0:
                        dfl = dfli
                    else:
                        dfl = pd.merge(dfl, dfli, left_on="invoice_id", right_on="invoice_id", how='inner')
                        # dfl = dfl.append(dfli)
            else:
                wdata = {'invoice_id': [i]}
                dfl = pd.DataFrame(wdata)
            try:
                dfs = dfs.append(dfl)
            except:
                dfs = dfl
        dfpayment = pd.merge(dfdata, dfs, how='inner', left_on=['invoice_id'], right_on=['invoice_id'])
        return dfpayment

    def invoice_discount_split(self, dfdata):
        df = dfdata[["invoice_id", "invoice_line_item_discounts"]]
        dfs = pd.DataFrame
        dfl = pd.DataFrame
        df['invoice_line_item_discounts'] = df['invoice_line_item_discounts'].replace("'", '"', regex=True)
        df['invoice_line_item_discounts'] = df['invoice_line_item_discounts'].replace(": False,", ': "False",', regex=True)
        df['invoice_line_item_discounts'] = df['invoice_line_item_discounts'].replace(": True,", ': "True",', regex=True)
        for i, j in zip(df['invoice_id'], df['invoice_line_item_discounts']):
            print("splitting for '{}' invoice id and the date is :{}".format(i, j))
            if not pd.isna(j) and j != '[]':
                data = json.loads(j)
                for k in range(len(data)):
                    prefix = "discounts_"
                    dfli = pd.json_normalize(data[k])
                    sufix = "[" + str(k) + "]"
                    headers = list(dfli.head())
                    newheaders = {}
                    for ch in headers:
                        newheaders[ch] = prefix + ch + sufix
                    dfli.rename(columns=newheaders, inplace=True)
                    dfli['invoice_id'] = [i]
                    if k == 0:
                        dfl = dfli
                    else:
                        dfl = pd.merge(dfl, dfli, left_on="invoice_id", right_on="invoice_id", how='inner')
                        # dfl = dfl.append(dfli)
            else:
                wdata = {'invoice_id': [i]}
                dfl = pd.DataFrame(wdata)
            try:
                dfs = dfs.append(dfl)
            except:
                dfs = dfl
        dfpayment = pd.merge(dfdata, dfs, how='inner', left_on=['invoice_id'], right_on=['invoice_id'])
        return dfpayment


invoiceobj = InvoiceExecution()
invoiceobj.getAllInvoices()
