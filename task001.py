# -*- coding: utf-8 -*-

import xlrd
import datetime
import math
import sys

version = sys.version_info.major    # 大版本号
if version != 3:
    reload(sys)
    sys.setdefaultencoding('utf8')


# 日期
COL_INDEX_DATE = 0
# 复权单位净值
COL_INDEX_WORTH = 2

# 日期索引
date_arr = []
date_arr_desc = []
# 复权单位净值索引
worth_arr = []


def date_delta(start_date_str, end_date_str):
    """
    日期差值
    :param date1:
    :param date2:
    :return:
    """
    start_date = datetime.datetime.strptime(start_date_str, "%Y-%m-%d")
    end_date = datetime.datetime.strptime(end_date_str, "%Y-%m-%d")
    if start_date < end_date:
        return end_date - start_date
    else:
        return start_date - end_date


def find_gte_worth(start_date_index, end_date_index, worth):
    # 区间内最大复权单位净值
    max_range_worth = 0
    max_range_worth_index = 0
    tmp_start_date_index = start_date_index
    for worth_val in worth_arr[start_date_index:end_date_index]:
        if float(worth_val) > max_range_worth:
            max_range_worth = worth_val
            max_range_worth_index = tmp_start_date_index
        if float(worth_val) >= float(worth):
            return tmp_start_date_index, None
        tmp_start_date_index += 1
    return None, max_range_worth_index


def calculator_pain_index(start_date, end_date):
    """
    计算痛苦指数
    :param start_date:
    :param end_date:
    :param date_arr:
    :param worth_arr:
    :return:
    """
    print('时间区间: %s -> %s: ' % (start_date, end_date))
    start_date_index = date_arr.index(start_date)
    end_date_index = date_arr.index(end_date)
    # print(start_date_index, end_date_index)
    start_worth = worth_arr[start_date_index]

    print('起点复权单位净值: %s  ' % (start_worth,))
    result_index, max_index = find_gte_worth(start_date_index + 1, end_date_index, start_worth)
    if result_index is None:
        print('未找到最近回本日期，取区间内最高点')
        print('回本日期: %s 复权单位净值: %s' % (date_arr[max_index], worth_arr[max_index]))
        years = date_delta(start_date, date_arr[max_index])
    else:
        print('回本日期: %s 复权单位净值: %s' % (date_arr[result_index], worth_arr[result_index]))
        years = date_delta(start_date, date_arr[result_index])

    years = math.ceil(years.days / 36.5) / 10
    print('痛苦指数: %s年' % (years,))


f = xlrd.open_workbook(
    filename='519069.xlsx'
)
table = f.sheets()[0]
rows_len = table.nrows
for i in range(1, rows_len):
    row_values = table.row_values(rows_len - i)
    if not row_values[0] or str(row_values[0]).startswith('数'):
        continue
    # print(row_values[COL_INDEX_DATE], row_values[COL_INDEX_WORTH])
    date_arr.append(row_values[COL_INDEX_DATE])
    worth_arr.append(row_values[COL_INDEX_WORTH])

# 2015-06-12 最高 5.2421
# 2017-11-09 最高 5.259
calculator_pain_index('2015-06-12', '2017-12-01')
