#!/usr/bin/env python
# -*- coding:utf-8 -*-
# author:HeWenjun
# datetime:2022/7/8 10:06
# software: PyCharm

import time
import pytest
import pyautogui
import xlrd
from selenium.common.exceptions import NoSuchElementException


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
# 公司签报起草
def test_flow(logged_in_driver, title, article):
    wd = logged_in_driver
    # 进入公文
    # WebDriverWait(wd, 60).until(lambda wd: wd.find_element_by_id("business")).click()
    # wd.find_element_by_id('business').click()
    wd.get('https://oazsctest.migu.cn/framework-miguOA-web/navigation')
    # 切换到OA公文页面
    assert wd.title == 'OA办公'
    wd.switch_to.frame('iframe_a')
    wd.find_element_by_id('xzgw_content_2').click()
    # 切换到公文起草页面
    wd.switch_to.window(wd.window_handles[1])
    # 输入手机号码
    wd.find_element_by_id('phone').clear()
    wd.find_element_by_id('phone').send_keys('13333333333')
    # 点击单选框
    wd.find_element_by_css_selector("label:nth-child(2) span").click()
    # 输入公文标题
    wd.find_element_by_id('title').send_keys(title)
    # 选取主送人员
    wd.find_element_by_xpath('//*[@id="lordSentHtml"]/button').click()
    wd.switch_to.frame('layui-layer-iframe1')
    wd.find_element_by_id('groupWordsId_32_switch').click()
    time.sleep(1)
    wight, height = pyautogui.size()
    pyautogui.moveTo(wight * 0.9, height * 0.9)
    pyautogui.scroll(-10000)
    wd.find_element_by_id('groupWordsId_84_span').click()
    time.sleep(2)
    wd.switch_to.default_content()
    wd.find_element_by_xpath('/html/body/div[3]/div[3]/a[1]').click()
    # NTKO编辑正文
    while True:
        times = 1
        try:
            wd.switch_to.frame('NTKOFORM')
            # 首次进入循环，没有编辑过正文，必然没有该元素，直接进入except执行编辑操作
            # 为避免浪费时间，首次进入循环修改隐式等待时间10秒为3秒
            if times == 1:
                wd.implicitly_wait(3)
            else:
                wd.implicitly_wait(10)
            ntko_form = wd.find_element_by_xpath('/html/body/div/div/p/span').text
            if ntko_form != '':
                wd.switch_to.default_content()
                submit(wd)
                break
            else:
                raise NoSuchElementException
        except NoSuchElementException:
            wd.switch_to.default_content()
            wd.find_element_by_link_text("编辑正文").click()
            ntko(article)
    # 切换到OA办公窗口
    wd.switch_to.window(wd.window_handles[0])
    # 切换到待办
    wd.find_element_by_xpath('/html/body/div[7]/div[1]/div/div/ul/li[2]/a/span[2]').click()
    submit_to_clerk(wd, title)
    change_user(wd, title)
    change_back(wd, title)
    submit_to_department(wd, title)
    distribute(wd, title)
    time.sleep(3)
    wd.quit()
    # wd.find_element_by_xpath('/html/body/div[7]/div[1]/div/div/ul/li[4]/a/span[2]').click()
    # wd.switch_to_frame('iframe_a')
    # done_title = wd.find_element_by_xpath(
    #     '/html/body/div/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[1]/td[2]/div/div').text
    # assert done_title == title


# 起草完成提交流程部门领导审核
def submit(driver):
    # 点击提交按钮
    driver.find_element_by_id('submit').click()
    driver.switch_to.frame('layui-layer-iframe1')
    # 点击直接用一键提交何初升
    driver.find_element_by_xpath('/html/body/div/div[2]/div/div[1]/div[2]/div[2]/div/span/input').click()
    driver.switch_to.default_content()
    driver.find_element_by_xpath('/html/body/div[3]/div[3]/a[1]').click()


# 操作NTKO编辑正文
def ntko(article):
    # 使用第三方库pyautogui操作NTKO编辑正文
    pg = pyautogui
    # 获取屏幕尺寸
    width, height = pg.size()
    print(f'wight:{width},height:{height}')
    # 计算【保存并关闭】按钮的坐标
    to_wight = width * 0.07625
    to_height = height * 0.055
    print(f'to_wight:{to_wight},to_height:{to_height}')
    time.sleep(5)
    # 点击NTKO正文区域，用以转移文本焦点
    pg.click(width / 2, height / 2)
    # 尽量删除文本
    for i in range(len(article)):
        pg.press('backspace')
        pg.press('delete')
    # 模拟键盘输入“test”文本。注：typewrite不支持中文输入
    # 需要中文输入可以使用复制粘贴https://blog.csdn.net/weixin_42551921/article/details/122846980
    pg.typewrite(article)
    # 点击保存并关闭
    pg.click(to_wight, to_height)
    time.sleep(3)
    pg.press('enter')


# 部门领导送部门文书核稿
def submit_to_clerk(wd, title):
    wd.switch_to.frame('iframe_a')
    # 查询出对应公文
    wd.find_element_by_id('title').clear()
    wd.find_element_by_id('title').send_keys(title)
    wd.find_element_by_id('query_Btn').click()
    # 打开公文
    wd.find_element_by_xpath('/html/body/div/div[3]/div[2]/div[1]/div[2]/table/tbody/tr/td[2]/div/div').click()
    time.sleep(5)
    wd.switch_to.window(wd.window_handles[1])
    print(wd.current_url)
    wd.find_element_by_id('submit').click()
    wd.switch_to.frame('layui-layer-iframe1')
    wd.find_element_by_xpath('/html/body/div/div[2]/div/div[1]/div[2]/div[2]/div[1]/span/input').click()
    wd.switch_to.default_content()
    wd.find_element_by_xpath('/html/body/div[3]/div[3]/a[1]').click()
    time.sleep(1)


# 切换文书登录审批
def change_user(wd, title):
    wd.switch_to.window(wd.window_handles[0])
    # 点击退出
    wd.find_element_by_xpath('//*[@id="zhuxiao"]/a').click()
    # 登录文书账号
    wd.switch_to.frame('logonIframe')
    wd.find_element_by_id('wenb').send_keys('huminli')
    wd.find_element_by_id('worde').send_keys('M19u%20213')
    wd.find_element_by_id('submitButton').click()
    time.sleep(2)
    wd.find_element_by_xpath('/html/body/div[7]/div[1]/div/div/ul/li[2]/a/span[2]').click()
    wd.switch_to.frame('iframe_a')
    # 查询出对应公文
    wd.find_element_by_id('title').clear()
    wd.find_element_by_id('title').send_keys(title)
    wd.find_element_by_id('query_Btn').click()
    # 打开公文
    wd.find_element_by_xpath('/html/body/div/div[3]/div[2]/div[1]/div[2]/table/tbody/tr/td[2]/div/div').click()
    time.sleep(5)
    wd.switch_to.window(wd.window_handles[1])
    print(wd.current_url)
    wd.find_element_by_id('submit').click()
    # 切换到提交窗口
    wd.switch_to.frame('layui-layer-iframe1')
    wd.find_element_by_xpath('/html/body/div/div[2]/div/div[1]/div[2]/div[2]/div/span/input').click()
    wd.switch_to.default_content()
    wd.find_element_by_xpath('/html/body/div[3]/div[3]/a[1]').click()
    time.sleep(1)


# 切换回原用户
def change_back(wd, title):
    wd.switch_to.window(wd.window_handles[0])
    # 点击退出
    wd.find_element_by_xpath('//*[@id="zhuxiao"]/a').click()
    # 登录何初升账号
    wd.switch_to.frame('logonIframe')
    wd.find_element_by_id('wenb').send_keys('hechusheng')
    wd.find_element_by_id('worde').send_keys('M19u%20213')
    wd.find_element_by_id('submitButton').click()
    time.sleep(2)
    wd.find_element_by_xpath('/html/body/div[7]/div[1]/div/div/ul/li[2]/a/span[2]').click()


# 核稿送综合部分发
def submit_to_department(wd, title):
    wd.switch_to.window(wd.window_handles[0])
    print('当前链接', wd.current_url)
    wd.switch_to.frame('iframe_a')
    # 查询出对应公文
    wd.find_element_by_id('title').clear()
    wd.find_element_by_id('title').send_keys(title)
    wd.find_element_by_id('query_Btn').click()
    # 打开公文
    wd.find_element_by_xpath('/html/body/div/div[3]/div[2]/div[1]/div[2]/table/tbody/tr/td[2]/div/div').click()
    time.sleep(5)
    wd.switch_to.window(wd.window_handles[1])
    wd.find_element_by_id('submit').click()
    wd.switch_to.frame('layui-layer-iframe1')
    wd.find_element_by_xpath('/html/body/div/div[2]/div/div[1]/div[2]/div[2]/div[1]/span/input').click()
    wd.switch_to.default_content()
    wd.find_element_by_xpath('/html/body/div[3]/div[3]/a[1]').click()
    time.sleep(1)


# 实现分发
def distribute(wd, title):
    wd.switch_to.window(wd.window_handles[0])
    wd.get('https://oazsctest.migu.cn/framework-miguOA-web/navigation')
    # 切换到OA公文页面
    # wd.switch_to.window(wd.window_handles[1])
    wd.find_element_by_xpath('/html/body/div[7]/div[1]/div/div/ul/li[2]/a').click()
    wd.switch_to.frame('iframe_a')
    # 查询出对应公文
    wd.find_element_by_id('title').clear()
    wd.find_element_by_id('title').send_keys(title)
    wd.find_element_by_id('query_Btn').click()
    # 打开公文
    wd.find_element_by_xpath('/html/body/div/div[3]/div[2]/div[1]/div[2]/table/tbody/tr/td[2]/div/div').click()
    time.sleep(5)
    wd.switch_to.window(wd.window_handles[1])
    # 点击分发
    wd.find_element_by_id('splitSend').click()
    time.sleep(5)
    # 使用第三方库pyautogui操作NTKO套红正文
    pg = pyautogui
    # 获取屏幕尺寸
    wight, height = pg.size()
    print(f'wight:{wight},height:{height}')
    # 计算【套红并关闭】按钮的坐标
    pg.click(wight * 0.138, height * 0.055)
    # 等待套红结束和网页更新
    time.sleep(15)
    print('点击分发')
    wd.find_element_by_xpath('/html/body/header/ul/li[8]').click()
    time.sleep(2)
    print('点击确认')
    wd.find_element_by_xpath('/html/body/div[3]/div[3]/a[1]').click()
    time.sleep(5)


if __name__ == '__main__':
    print(_read_doc_info())
