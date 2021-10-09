import pandas as pd
import numpy as np
import re
from datetime import datetime

def zh_to_datetime(date):
    year, month = re.split('年|月', date)[:2]
    year = int(year)
    if len(month) == 1:
        month = int(month)
    elif (len(month) == 2) and (month[0] == '0'):
        month = int(month[1])
    elif (len(month) == 2) and (month[0] == '1'):
        month = int(month)
    date = datetime(year=year, month=month, day=1)
    return date


def datetime_to_zh(date):

    if date.month < 10:
        Month = "0" + str(date.month)
    else:
        Month = str(date.month)
    return (f"{date.year}年{Month}月")



def generate_time(year, month, k):
    """
    用于生成 y年m月 之后 k 期的月份名称
    :param k: 是k期，不包括本月
    :return: 返回的是一个list
    """
    date = []
    for i in range(k):
        if month < 12:
            month = month + 1
        elif month == 12:
            year = year + 1
            month = 1
        if month < 10:
            m = '0' + str(month)
            date.append('%s年%s月' % (year, m))
        else:
            date.append('%s年%s月' % (year, month))
    return date

