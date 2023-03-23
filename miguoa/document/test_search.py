#!/usr/bin/env python
# -*- coding:utf-8 -*-
# author:HeWenjun
# datetime:2022/7/15 15:01
# software: PyCharm
import random

from selenium.webdriver.support.select import Select


def test_search(logged_in_driver):
    wd = logged_in_driver
    wd.find_element_by_id('business').click()
    # 切换到OA公文页面
    wd.switch_to_window(wd.window_handles[1])
    assert wd.title == 'OA办公'
    wd.switch_to.frame('iframe_a')
    # 测标题
    wd.find_element_by_id('title_name').send_keys('测试')
    wd.find_element_by_id('query_Btn').click()
    form_title_name = wd.find_element_by_xpath(
        '/html/body/div[1]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[1]/td[2]/div/div').text
    assert '测试' in form_title_name
    wd.find_element_by_id('reset_Btn').click()
    # 测文号
    wd.find_element_by_id('doc_No').send_keys('2022')
    wd.find_element_by_id('query_Btn').click()
    form_doc_no = wd.find_element_by_xpath(
        '/html/body/div[1]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[1]/td[5]/div').text
    assert '2022' in form_doc_no
    wd.find_element_by_id('reset_Btn').click()
    # 测文种
    ele = wd.find_element_by_id('docId')
    options = Select(ele).options
    select_index = random.randint(1, len(options) - 1)
    Select(ele).select_by_index(select_index)
    wd.find_element_by_id('query_Btn').click()
    form_doc_id = wd.find_element_by_xpath(
        '/html/body/div/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[1]/td[4]/div').text
    print(Select(ele).all_selected_options[0].text)
    assert form_doc_id == Select(ele).all_selected_options[0].text

