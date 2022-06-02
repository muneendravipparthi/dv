from datetime import datetime

import pytz


class Ephoc2DateTime:

    def epoch_To_Datetime_Convert(epochtimestamp, timezoneOfCustomer):
        my_datetime = datetime.fromtimestamp(epochtimestamp, tz=pytz.timezone(timezoneOfCustomer))
        modified = my_datetime.strftime('%Y-%m-%d %H:%M:%S')
        return modified

    def epoch_To_ShortDatetime_Convert(epochtimestamp, timezoneOfCustomer):
        my_datetime = datetime.fromtimestamp(epochtimestamp, tz=pytz.timezone(timezoneOfCustomer))
        modified = my_datetime.strftime('%d/%m/%Y')
        return modified

    def centToDollar(amt):
        centValue = amt / 100
        return centValue
