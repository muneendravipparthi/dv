import json
import logging
import os

import pandas as pd
from jproperties import Properties
from tqdm import tqdm

from readAPI.ReadAPI import ReadAPIExecution
from readAPI.splitHelper import SplitHelper

configs = Properties()
home_folder = os.getenv('HOME')
ROOT_DIR1 = os.path.abspath(os.curdir)
ROOT_DIR = ROOT_DIR1.replace('readAPI', 'configuration.properties')
jsonDir = ROOT_DIR1 + '/jsonfiles'
excelDir = ROOT_DIR1 + '/ds3files'
logDir = ROOT_DIR1 + '/logfiles'
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
logging.basicConfig(filename=logDir + "/invoice.log",
                    format='[%(asctime)s] %(lineno)d %(levelname)s - %(message)s',
                    filemode='w')
# Let us Create an object
logger = logging.getLogger()

# Now we are going to Set the threshold of logger to DEBUG
logger.setLevel(logging.DEBUG)
outputFile = "_DS3_AllInvoices.xlsx"
RequireAPIExecution = False


class InvoiceExecution:

    def getAllInvoices(self):
        if RequireAPIExecution:
            logger.info("Executiong Extraction of Invoices")
            url = clientSite + invoiceextenction
            TotalInvoicesResponse = ReadAPIExecution.getDataFromAPI(self, url, user, logger)
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
            df_nested_list.to_excel(excelDir + '/' + configs.get("clientName").data + outputFile, index=False)
            try:
                if "invoice_line_items" in list(df_nested_list.head()):
                    df_splitlineitems = pd.read_excel(excelDir + '/' + configs.get("clientName").data + outputFile)
                    df_splitlineitems = SplitHelper.invoice_lineitem_split(self, df_splitlineitems)
                    df_splitlineitems.to_excel(excelDir + '/' + configs.get("clientName").data + outputFile,
                                               index=False)
            except Exception as e:
                logger.info(e)
                logger.error("exception at invoice_line_items:" + str(e))
                logger.exception(e)
            try:
                if "invoice_line_item_taxes" in list(df_nested_list.head()):
                    df_splitlineitemtaxes = pd.read_excel(excelDir + '/' + configs.get("clientName").data + outputFile)
                    df_splitlineitemtaxes = SplitHelper.invoice_lineitemtaxes_split(self, df_splitlineitemtaxes)
                    df_splitlineitemtaxes.to_excel(excelDir + '/' + configs.get("clientName").data + outputFile,
                                                   index=False)
            except Exception as e:
                logger.info(e)
                logger.error("exception at invoice_line_item_taxes:" + str(e))
                logger.exception(e)
            try:
                if "invoice_linked_payments" in list(df_nested_list.head()):
                    df_splitpayments = pd.read_excel(excelDir + '/' + configs.get("clientName").data + outputFile)
                    df_splitpayments = SplitHelper.invoice_payment_split(self, df_splitpayments)
                    df_splitpayments.to_excel(excelDir + '/' + configs.get("clientName").data + outputFile, index=False)
            except Exception as e:
                logger.info(e)
                logger.error("exception at invoice_linked_payments:" + str(e))
                logger.exception(e)
            try:
                if "invoice_line_item_discounts" in list(df_nested_list.head()):
                    df_splitdiscounts = pd.read_excel(excelDir + '/' + configs.get("clientName").data + outputFile)
                    df_splitdiscounts = SplitHelper.invoice_discount_split(self, df_splitdiscounts)
                    df_splitdiscounts.to_excel(excelDir + '/' + configs.get("clientName").data + outputFile,
                                               index=False)
            except Exception as e:
                logger.info(e)
                logger.error("exception at invoice_line_item_discounts:" + str(e))
                logger.exception(e)

            try:
                if "invoice_discounts" in list(df_nested_list.head()):
                    df_splitdiscounts = pd.read_excel(excelDir + '/' + configs.get("clientName").data + outputFile)
                    df_splitdiscounts = SplitHelper.invoice_discounts_split(self, df_splitdiscounts)
                    df_splitdiscounts.to_excel(excelDir + '/' + configs.get("clientName").data + outputFile,
                                               index=False)
            except Exception as e:
                logger.info(e)
                logger.error("exception at invoice_discounts_split:" + str(e))
                logger.exception(e)

            tdf = pd.read_excel(excelDir + '/' + configs.get("clientName").data + outputFile)
            dateconvertioncollist = ["invoice_date", "invoice_due_date", "invoice_paid_at", "invoice_updated_at",
                                     "invoice_generated_at", "line_items_date_from[0]", "line_items_date_to[0]",
                                     "line_items_date_from[1]", "line_items_date_to[1]", "line_items_date_from[2]",
                                     "line_items_date_to[2]", "line_items_date_from[3]", "line_items_date_to[3]",
                                     "line_items_date_from[4]", "line_items_date_to[4]", "line_items_date_from[5]",
                                     "line_items_date_to[5]", "line_items_date_from[6]", "line_items_date_to[6]",
                                     "line_items_date_from[7]", "line_items_date_to[7]", "line_items_date_from[8]",
                                     "line_items_date_to[8]", "line_items_date_from[9]",
                                     "line_items_date_to[9]", "line_items_date_from[10]",
                                     "line_items_date_to[10]", "applied_at", "txn_date", "payments_txn_date[0]",
                                     "payments_txn_date[1]", ]
            for col in tqdm(dateconvertioncollist, desc='dateconvertioncollist'):
                if col in list(tdf.head()):
                    tdf[col] = tdf[col].apply(
                        lambda x: ReadAPIExecution.epoch_To_Datetime_Convert(self, x, clienttimezone) if pd.isna(
                            x) != True else None)
            tdf.to_excel(excelDir + '/' + configs.get("clientName").data + outputFile, index=False)

            if cents == 'False':
                centsToDoller = pd.read_excel(excelDir + '/' + configs.get("clientName").data + outputFile)
                centsToDollerlist = ["invoice_total", "invoice_amount_paid", "invoice_amount_adjusted",
                                     "invoice_write_off_amount", "invoice_credits_applied", "invoice_amount_due",
                                     "invoice_amount_to_collect", "invoice_new_sales_amount",
                                     "line_items_unit_amount[0]",
                                     "line_items_amount[0]", "line_items_tax_amount[0]", "line_items_unit_amount[1]",
                                     "line_items_amount[1]", "line_items_tax_amount[1]", "line_items_unit_amount[2]",
                                     "line_items_amount[2]", "line_items_tax_amount[2]", "line_items_unit_amount[3]",
                                     "line_items_amount[3]", "line_items_tax_amount[3]", "line_items_unit_amount[4]",
                                     "line_items_amount[4]", "line_items_tax_amount[4]", "line_items_unit_amount[5]",
                                     "line_items_amount[5]", "line_items_tax_amount[5]", "line_items_unit_amount[6]",
                                     "line_items_amount[6]", "line_items_tax_amount[6]", "line_items_unit_amount[7]",
                                     "line_items_amount[7]", "line_items_tax_amount[7]", "applied_amount", "txn_amount",
                                     "payments_txn_amount[0]", "payments_txn_amount[1]", "discounts_discount_amount[0]",
                                     "discounts_discount_amount[1]", "discounts_discount_amount[2]"]

                for col in tqdm(centsToDollerlist, desc='centsToDollerlist'):
                    if col in list(tdf.head()):
                        centsToDoller[col] = centsToDoller[col].div(100)
                centsToDoller.to_excel(excelDir + '/' + configs.get("clientName").data + outputFile, index=False)
                logger.info("Completed data conversion from Json to Excel")
        except Exception as e:
            logger.info("Something failed during data conversion from Json to Excel")
            logger.error("exception in invoices:" + str(e))
            logger.exception(e)




invoiceobj = InvoiceExecution()
invoiceobj.getAllInvoices()
