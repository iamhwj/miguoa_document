#!/usr/bin/env python
# -*- coding:utf-8 -*-
# author:HeWenjun
# datetime:2022/7/7 9:55
# software: PyCharm
import time

import pytest

# 参数化
import xlrd


def _read_userinfo():
    data_list = []
    file = xlrd.open_workbook(r'D:\Python Projects\miguoa\data_file\登录用户数据.xlsx')
    # sheet表名
    # print(file.sheet_names())
    # 使用sheet1的数据
    sheet = file.sheet_by_index(0)
    # 获取每行的数据
    for i in range(sheet.nrows - 1):
        # print(sheet.row_values(i + 1))
        username = sheet.row_values(i + 1)[0]
        password = sheet.row_values(i + 1)[1]
        expect = sheet.row_values(i + 1)[2]
        # 构造数据格式=元组为参数的列表:[(a,b),(1,2)]
        data_list.append((str(username), str(password), str(expect)))
    return data_list


@pytest.mark.parametrize('username,password,expect', _read_userinfo())
# 使用公共文件中browser_driver的driver对象
def test_login(browser_driver, username, password, expect):
    browser_driver.get('https://oazsctest.migu.cn/portal/login/index')
    bd = browser_driver
    # bd.implicitly_wait(10)
    bd.switch_to_frame('logonIframe')
    bd.find_element_by_id('wenb').send_keys(username)
    bd.find_element_by_id('worde').send_keys(password)
    bd.find_element_by_id('submitButton').click()
    time.sleep(3)
    if expect == '成功':
        assert bd.title == '咪咕门户' or bd.title == '密码变更'
    elif expect == '失败':
        assert bd.title != '咪咕门户'
    bd.close()


if __name__ == '__main__':
    pytest.main(['-s', 'test_login.py'])
