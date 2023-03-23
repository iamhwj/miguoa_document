#!/usr/bin/env python
# -*- coding:utf-8 -*-
# author:HeWenjun
# datetime:2022/8/18 17:38
# software: PyCharm
import xlrd, pytest


def _read_doc_info():
    data_list = []
    file = xlrd.open_workbook(r'D:\Python Projects\miguoa\data_file\起草公文数据.xlsx')
    # sheet表名
    # print(file.sheet_names())
    # 使用sheet1的数据
    sheet = file.sheet_by_index(0)
    # 获取每行的数据
    for i in range(sheet.nrows - 1):
        # print(sheet.row_values(i + 1))
        title = sheet.row_values(i + 1)[0]
        article = sheet.row_values(i + 1)[1]
        # 构造数据格式=元组为参数的列表:[(a,b),(1,2)]
        data_list.append((str(title), str(article)))
    return data_list


@pytest.mark.parametrize('title, article', _read_doc_info())
def test(logged_in_driver, title, article):
    print(title,article)
