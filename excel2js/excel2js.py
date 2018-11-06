#!/usr/bin/python
# -*- coding:utf-8 -*-

"""
Python version: V3.7

Author: Robin&HXN

File version: V1.0

File name: excel2js.py

Created on: 20181029

Resume: Transform buried_point Excel to Json

"""

import xlrd
from collections import OrderedDict
import json
import datetime as dt


def excel_to_json(file_path, save_path):
    """

    :param file_path: the path of excel file
    :param save_path: which path you want to save json file
    :return: json
    """

    wb = xlrd.open_workbook(file_path)
    convert_list = []
    sh = wb.sheet_by_index(0)
    title = sh.row_values(1)  # 表头，json文件的key
    print(title)
    for row_num in range(2, sh.nrows):
        row_value = sh.row_values(row_num)
        single = OrderedDict()  # 有序字典
        for col_num in range(0, len(row_value)):
            if isinstance(row_value[col_num], float):
                row_value[col_num] = int(row_value[col_num])
            # print("key:{0}, value:{1}".format(title[col_num], row_value[col_num]))
            single[title[col_num]] = row_value[col_num]
        convert_list.append(single)

    j = json.dumps(convert_list)

    with open(save_path, "w", encoding="utf8") as f:
        f.write(j)


if __name__ == '__main__':
    file_name = input("Please input Excel name:")
    f_path = u"/Users/huangxingnai/to_excel/%s.xlsx" % file_name
    s_path = u"/Users/huangxingnai/to_json/%s.json" % file_name
    excel_to_json(f_path, s_path)
    print(dt.datetime.today().strftime('%Y-%m-%d %H:%M:%S'))
    exit(0)
