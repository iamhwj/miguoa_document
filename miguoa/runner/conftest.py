#!/usr/bin/env python
# -*- coding:utf-8 -*-
# author:HeWenjun
# datetime:2022/7/7 9:47
# software: PyCharm
import time, os

import pytest
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# 仅启动浏览器
@pytest.fixture(scope='function')
def browser_driver():
    wd = webdriver.Chrome()
    print('###执行启动浏览器###')
    wd.get('https://oazsctest.migu.cn/portal/login/index')
    return wd


# 提前完成登录
@pytest.fixture(scope='function')
def logged_in_driver():
    # 设置webdriver
    option = webdriver.ChromeOptions()
    # 向启动的浏览器添加NTKO插件
    extension_path = r'C:\Users\HeWenjun\AppData\Local\Google\Chrome\User Data\Default\Extensions\lppkeogbkjlmmbjenbogdndlgmpiddda\1.8.3_0.crx'
    option.add_extension(extension=extension_path)
    wd = webdriver.Chrome(options=option,executable_path=r'C:\Users\HeWenjun\Downloads\chromedriver_win32\chromedriver.exe')
    wd.maximize_window()
    wd.implicitly_wait(10)
    print('###执行启动浏览器###')
    # 执行登录
    wd.get('https://oazsctest.migu.cn/portal/login/index')
    wd.switch_to.frame('logonIframe')
    wd.find_element_by_id('wenb').send_keys('hechusheng')
    wd.find_element_by_id('worde').send_keys('M19u%20213')
    wd.find_element_by_id('submitButton').click()
    time.sleep(2)
    print(wd.current_url)
    return wd
