import json
import logging
import os
from datetime import datetime, timedelta
import pandas as pd
from jproperties import Properties
import pytz


configs = Properties()
home_folder = os.getenv('HOME')
ROOT_DIR1 = os.path.abspath(os.curdir)
ROOT_DIR = '/Users/cb-muneendra/cb_data_validation/configuration.properties'
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
inputfile = 'pixellu_invoice_lineitems14.xlsx'
outputfile = 'output_pixellu_invoice_lineitems14.xlsx'

class InvoiceExecution:

    def getAllInvoices(self):
        try:
            df_splitlineitems = pd.read_excel(inputfile)
            if "invoice_line_items" in list(df_splitlineitems.head()):
                df_splitlineitems = self.invoice_lineitem_split(df_splitlineitems)
                df_splitlineitems.to_excel(outputfile, index=False)
            # if "invoice_line_item_taxes" in list(df_nested_list.head()):
            #     df_splitlineitemtaxes = pd.read_excel(outputfile)
            #     df_splitlineitemtaxes = self.invoice_lineitemtaxes_split(df_splitlineitemtaxes)
            #     df_splitlineitemtaxes.to_excel(outputfile, index=False)
            # if "invoice_linked_payments" in list(df_nested_list.head()):
            #     df_splitpayments = pd.read_excel(outputfile)
            #     df_splitpayments = self.invoice_payment_split(df_splitpayments)
            #     df_splitpayments.to_excel(outputfile, index=False)
            tdf = pd.read_excel(outputfile)
            dateconvertioncollist = ["invoice_date", "invoice_due_date", "invoice_paid_at", "invoice_updated_at",
                                     "invoice_generated_at", "line_item_date_from[0]", "line_item_date_to[0]",
                                     "line_item_date_from[1]", "line_item_date_to[1]", "line_item_date_from[2]",
                                     "line_item_date_to[2]", "line_item_date_from[3]", "line_item_date_to[3]",
                                     "line_item_date_from[4]", "line_item_date_to[4]", "line_item_date_from[5]",
                                     "line_item_date_to[5]", "line_item_date_from[6]", "line_item_date_to[6]",
                                     "line_item_date_from[7]", "line_item_date_to[7]", "line_item_date_from[8]",
                                     "line_item_date_to[8]", "applied_at", "txn_date"]
            for col in dateconvertioncollist:
                if col in list(tdf.head()):
                    tdf[col] = tdf[col].apply(
                        lambda x: self.epoch_To_Datetime_Convert(x, clienttimezone) if pd.isna(x) != True else None)
            tdf.to_excel(outputfile, index=False)
            if cents == False:
                centsToDoller = pd.read_excel(outputfile)
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
                                     "line_item_amount[7]", "line_item_tax_amount[7]", "applied_amount", "txn_amount"]
                for col in centsToDollerlist:
                    if col in list(tdf.head()):
                        centsToDoller[col] = centsToDoller[col].div(100)
                centsToDoller.to_excel(outputfile, index=False)
                logger.info("Converting Json data to Excel data Completed")
        except Exception as e:
            logger.info("Something failed during data convertion from Json to Excel")
            logger.error("exception in invoices:" + str(e))
            logger.exception(e)

    def invoice_lineitem_split(self,dfdata):
        df = dfdata[["invoice_id", "invoice_line_items"]]
        dfs = pd.DataFrame
        dfl = pd.DataFrame
        df['invoice_line_items'] = df['invoice_line_items'].replace("'", '"', regex=True)
        df['invoice_line_items'] = df['invoice_line_items'].replace(": False,", ': "False",', regex=True)
        df['invoice_line_items'] = df['invoice_line_items'].replace(": True,", ': "True",', regex=True)
        for i, j in zip(df['invoice_id'], df['invoice_line_items']):
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
                        dfl = dfli
                    else:
                        dfl = dfl.append(dfli)
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
                        dfl = dfl.append(dfli)
            else:
                wdata = {'invoice_id': [i]}
                dfl = pd.DataFrame(wdata)
            try:
                dfs = dfs.append(dfl)
            except:
                dfs = dfl
        dflineitemtax = pd.merge(dfdata, dfs, how='inner', left_on=['invoice_id'],
                         right_on=['invoice_id'])
        return dflineitemtax

    def invoice_payment_split(self, dfdata):
        df = dfdata[["invoice_id", "invoice_linked_payments"]]
        dfs = pd.DataFrame
        dfl = pd.DataFrame
        df['invoice_linked_payments'] = df['invoice_linked_payments'].replace("'", '"', regex=True)
        df['invoice_linked_payments'] = df['invoice_linked_payments'].replace(": False,", ': "False",', regex=True)
        df['invoice_linked_payments'] = df['invoice_linked_payments'].replace(": True,", ': "True",', regex=True)
        for i, j in zip(df['invoice_id'], df['invoice_linked_payments']):
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
                        dfl = dfl.append(dfli)
            else:
                wdata = {'invoice_id': [i]}
                dfl = pd.DataFrame(wdata)
            try:
                dfs = dfs.append(dfl)
            except:
                dfs = dfl
        dfpayment = pd.merge(dfdata, dfs, how='inner', left_on=['invoice_id'], right_on=['invoice_id'])
        return dfpayment

    def epoch_To_Datetime_Convert(self, epochtimestamp, timezoneOfCustomer):
        my_datetime = datetime.fromtimestamp(epochtimestamp, tz=pytz.timezone(timezoneOfCustomer))
        modified = my_datetime.strftime(datetimeformat)
        if modified.endswith(':59') and addonesecond:
            expected_date = datetime.strptime(modified, '%Y-%m-%d %H:%M:%S')
            expected_date += timedelta(seconds=1)
            modified = expected_date.strftime(datetimeformat)
        return modified

invoiceobj = InvoiceExecution()
invoiceobj.getAllInvoices()
