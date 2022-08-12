import json
import logging
import pandas as pd
from tqdm import tqdm
logger = logging.getLogger()

class SplitHelper:
    def subscription_subscription_items_split(self, dfdata):
        df = dfdata[["subscription_id", "subscription_subscription_items"]]
        dfs = pd.DataFrame
        dfl = pd.DataFrame
        df['subscription_subscription_items'] = df['subscription_subscription_items'].replace("'", '"', regex=True)
        df['subscription_subscription_items'] = df['subscription_subscription_items'].replace(": False,", ': "False",', regex=True)
        df['subscription_subscription_items'] = df['subscription_subscription_items'].replace(": True,", ': "True",', regex=True)
        for i, j in tqdm(zip(df['subscription_id'], df['subscription_subscription_items']), total=len(df['subscription_id']),
                         desc='subscription_subscription_items'):
            logger.info("splitting for '{}' subscription_id and the date is :{}".format(i, j))
            if not pd.isna(j) and j != '[]':
                data = json.loads(j)
                for k in range(len(data)):
                    prefix = "item_"
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
        dflineitem = pd.merge(dfdata, dfs, how='inner', left_on=['subscription_id'], right_on=['subscription_id'])
        return dflineitem


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

    def subscription_item_tiers(self, dfdata):
        df = dfdata[["subscription_id", "subscription_item_tiers"]]
        dfs = pd.DataFrame
        dfl = pd.DataFrame
        df['subscription_item_tiers'] = df['subscription_item_tiers'].replace("'", '"', regex=True)
        df['subscription_item_tiers'] = df['subscription_item_tiers'].replace(": False,", ': "False",', regex=True)
        df['subscription_item_tiers'] = df['subscription_item_tiers'].replace(": True,", ': "True",', regex=True)
        for i, j in tqdm(zip(df['subscription_id'], df['subscription_item_tiers']), total=len(df['subscription_id']),
                         desc='subscription_item_tiers'):
            logger.info("splitting for '{}' subscription_id and the date is :{}".format(i, j))
            if not pd.isna(j) and j != '[]':
                data = json.loads(j)
                for k in range(len(data)):
                    prefix = "item_tiers_"
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
        dftiers = pd.merge(dfdata, dfs, how='inner', left_on=['subscription_id'], right_on=['subscription_id'])
        return dftiers

    def subscription_discounts(self, dfdata):
        df = dfdata[["subscription_id", "subscription_discounts"]]
        dfs = pd.DataFrame
        dfl = pd.DataFrame
        df['subscription_discounts'] = df['subscription_discounts'].replace("'", '"', regex=True)
        df['subscription_discounts'] = df['subscription_discounts'].replace(": False,", ': "False",', regex=True)
        df['subscription_discounts'] = df['subscription_discounts'].replace(": True,", ': "True",', regex=True)
        for i, j in tqdm(zip(df['subscription_id'], df['subscription_discounts']), total=len(df['subscription_id']),
                         desc='subscription_discounts'):
            logger.info("splitting for '{}' subscription_id and the date is :{}".format(i, j))
            if not pd.isna(j) and j != '[]':
                data = json.loads(j)
                for k in range(len(data)):
                    prefix = "discount_"
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
        dfdiscounts = pd.merge(dfdata, dfs, how='inner', left_on=['subscription_id'], right_on=['subscription_id'])
        return dfdiscounts

    def invoice_lineitem_split(self, dfdata):
        df = dfdata[["invoice_id", "invoice_line_items"]]
        dfs = pd.DataFrame
        dfl = pd.DataFrame
        df['invoice_line_items'] = df['invoice_line_items'].replace("Tiina's addon", "Tiina^s addon", regex=True)
        df['invoice_line_items'] = df['invoice_line_items'].replace("'", '"', regex=True)
        df['invoice_line_items'] = df['invoice_line_items'].replace(": False,", ': "False",', regex=True)
        df['invoice_line_items'] = df['invoice_line_items'].replace(": True,", ': "True",', regex=True)
        for i, j in tqdm(zip(df['invoice_id'], df['invoice_line_items']), total=len(df['invoice_id']),
                         desc='invoice_line_items'):
            logger.info("splitting for '{}' invoice id and the date is :{}".format(i, j))
            if not pd.isna(j):
                data = json.loads(j)
                for k in range(len(data)):
                    prefix = "line_items_"
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
                        dfl = pd.merge(dfl, dfli, left_on="invoice_id", right_on="invoice_id", how='inner')
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

    def invoice_lineitemtaxes_split(self, dfdata):
        df = dfdata[["invoice_id", "invoice_line_item_taxes"]]
        dfs = pd.DataFrame
        dfl = pd.DataFrame
        df['invoice_line_item_taxes'] = df['invoice_line_item_taxes'].replace("'", '"', regex=True)
        df['invoice_line_item_taxes'] = df['invoice_line_item_taxes'].replace(": False,", ': "False",', regex=True)
        df['invoice_line_item_taxes'] = df['invoice_line_item_taxes'].replace(": False}", ': "False"}', regex=True)
        df['invoice_line_item_taxes'] = df['invoice_line_item_taxes'].replace(": True}", ': "True"}', regex=True)
        df['invoice_line_item_taxes'] = df['invoice_line_item_taxes'].replace(": True,", ': "True",', regex=True)
        for i, j in tqdm(zip(df['invoice_id'], df['invoice_line_item_taxes']), total=len(df['invoice_id']),
                         desc='invoice_line_item_taxes'):
            logger.info("splitting for '{}' invoice id and the date is :{}".format(i, j))
            if not pd.isna(j):
                data = json.loads(j)
                for k in range(len(data)):
                    prefix = "line_items_taxes_"
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
        for i, j in tqdm(zip(df['invoice_id'], df['invoice_linked_payments']), total=len(df['invoice_id']),
                         desc='invoice_linked_payments'):
            logger.info("splitting for '{}' invoice id and the date is :{}".format(i, j))
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
        df = dfdata[["invoice_id", "invoice_line_items_discounts"]]
        dfs = pd.DataFrame
        dfl = pd.DataFrame
        df['invoice_line_item_discounts'] = df['invoice_line_item_discounts'].replace("'", '"', regex=True)
        df['invoice_line_item_discounts'] = df['invoice_line_item_discounts'].replace(": False,", ': "False",',
                                                                                      regex=True)
        df['invoice_line_item_discounts'] = df['invoice_line_item_discounts'].replace(": True,", ': "True",',
                                                                                      regex=True)
        for i, j in tqdm(zip(df['invoice_id'], df['invoice_line_item_discounts']), total=len(df['invoice_id']),
                         desc='invoice_line_item_discounts'):
            logger.info("splitting for '{}' invoice id and the date is :{}".format(i, j))
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

    def invoice_discounts_split(self, dfdata):
        df = dfdata[["invoice_id", "invoice_discounts"]]
        dfs = pd.DataFrame
        dfl = pd.DataFrame
        df['invoice_discounts'] = df['invoice_discounts'].replace("'", '"', regex=True)
        df['invoice_discounts'] = df['invoice_discounts'].replace(": False,", ': "False",',
                                                                                      regex=True)
        df['invoice_discounts'] = df['invoice_discounts'].replace(": True,", ': "True",',
                                                                                      regex=True)
        for i, j in tqdm(zip(df['invoice_id'], df['invoice_discounts']), total=len(df['invoice_id']),
                         desc='invoice_discounts'):
            logger.info("splitting for '{}' invoice id and the date is :{}".format(i, j))
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