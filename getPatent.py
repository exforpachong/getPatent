import numpy as np ,pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import logging
import re
import requests
import os
import pyautogui

#输入默认的全局参数
fileName = '数据总表-深市沪市去重后数据汇总 - 合并为一个表.xlsx'   #需要处理的文件名
fileSave = 'results.txt'   #保存结果的文件名
waitSecond = 10   #这里设置载入一个页面等待页面响应的最大时间

def init_browser():   #初始化浏览器的函数，该函数的作用是配置请求的用户浏览器信息，以及用户信息
    chromeOptions = webdriver.ChromeOptions()
    # 设置代理
    # chromeOptions.add_argument("--proxy-server=http://121.233.207.225:9999")  #这里注释因为没有花钱买代理ip
    user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36'
    user_data = 'H:/github/qichacha/User Data'  # 这里面填你的user data对应的文件夹
    # user_data = 'C:/Users/dell/AppData/Local/Google/Chrome/User Data'   #测试自己的数据
    # user_agent = random.choice(agents)  #有些服务器需多次登录需要设置不同用户浏览器信息来反爬虫
    chromeOptions.add_argument('user-agent=%s' % user_agent)
    chromeOptions.add_argument('--user-data-dir=%s' % user_data)
    # cookies = {'IS_LOGIN': 'true',
    #            'JSESSIONID': 'i7SHo_6qZ_Diyg7ivNGCbBrwtQksogdhCt7Nhqlljpi2JFRZFjZ7!1491867009!1065980549',
    #            'WEE_SID': 'i7SHo_6qZ_Diyg7ivNGCbBrwtQksogdhCt7Nhqlljpi2JFRZFjZ7!1491867009!1065980549!1569938734762',
    #            'wee_password':'bGlqaWU1MjE%3D',
    #            'wee_username':'d2VuamllMTg3'
    #            }
    # chromeOptions.add_argument('Cookie=%s'% cookies)    #这一块没有搞懂，cookie还是不会设置，所以前面通过配置user_data信息来达到同样效果
    browser = webdriver.Chrome(chrome_options=chromeOptions)
    return browser

def clickByXPath(browser,xpath):
    WebDriverWait(browser, waitSecond, 0.5).until(EC.presence_of_element_located((By.XPATH, xpath)))  # 等待元素出现
    button = browser.find_element_by_xpath(xpath)  # 得到按钮
    ActionChains(browser).move_to_element(button).perform();  # 滚动到可以看到按钮
    WebDriverWait(browser, waitSecond, 0.5).until(EC.element_to_be_clickable((By.XPATH, xpath)))  # 等待元素可点击
    button.send_keys(Keys.ENTER)  # 点击回车代替点击好，否则会出错！

def getYearNumByXPath(browser,xpath):
    WebDriverWait(browser, waitSecond, 0.5).until(EC.presence_of_element_located((By.XPATH, xpath)))  # 等待元素出现
    lis = browser.find_elements_by_xpath(xpath)
    data = []
    for li in lis:
        y_n = li.text
        year = y_n[:y_n.index('(')]
        num = y_n[y_n.index('(') + 1:y_n.index(')')]
        # data += [[year, num]]  #方式一
        data += [y_n.strip()]  #方式二
    return data

def save(content):
    with open(fileSave,'a+') as f:
        f.write(content[0]+'\t'+content[1]+'\t'+content[2]+'\t'+content[3]+'\t'+content[4]+'\n')

if __name__ == '__main__':
    dataPD = pd.read_excel(fileName, sheet_name='汇总表')
    company_names = list(dataPD['被参控公司'].values)
    try:
        with open(fileSave, 'r') as f:
            lines = f.readlines()
            name = lines[-1].split('\t')[0]
        w = company_names.index(name)
        print('已经搜索到['+company_names[w]+']，从['+company_names[w+1]+']开始检索')
    except Exception as e:
        print(e)
        print('没有检索匹配到公司名称，从头开始搜索！')
        w = -1
    company_names = company_names[w + 1:]

    url = 'http://pss-system.cnipa.gov.cn/sipopublicsearch/patentsearch/tableSearch-showTableSearchIndex.shtml'
    browser = init_browser()  # 初始化浏览器，配置请求的用户浏览器信息，以及用户信息
    browser.get(url)  # 进入主页
    print('waiting for login user!')
    time.sleep(10)  # 打开浏览器的瞬间需要手动登录，所以等待了10s，主要手速要快，否则这里设置时间长一点！
    print('begin to search!')

    # company_names = ['平安银行股份有限公司','万科企业股份有限公司','上海凌云实业发展股份有限公司']
    for company_name in company_names:
        try:
            browser.refresh()  # 刷新当前页面准备下一次搜索
            #寻找输入框并输入需要搜索的公司名称
            xpath_input = '//*[@id="tableSearchItemIdIVDB020"]'   #申请（专利权）人的XPath
            WebDriverWait(browser, waitSecond, 0.5).until(EC.presence_of_element_located((By.XPATH, xpath_input)))  #等待’申请（专利权）人‘出现
            browser.find_element_by_xpath(xpath_input).send_keys(company_name)  #输入公司名称
            #寻找检索按钮并单击检索按钮
            xpath_search = '/html/body/div[3]/div[3]/div/div[2]/div[3]/a[3]'  #搜索按钮的XPath
            clickByXPath(browser,xpath_search)  #点击搜索
            #寻找并记录申请日统计的数据
            applyDayStas = '//*[@id="itemsList"]//a[@sort=\'ADY\']'  #申请日统计栏的XPath
            clickByXPath(browser, applyDayStas)  #点击展开申请日统计栏
            applyYearsStas = '//*[@id="itemsList"]//ul[@sort=\'ADY\']/li'  #申请日统计栏年(申请数)的XPath
            applyData = getYearNumByXPath(browser,applyYearsStas)  #得到申请日统计数据
            # 寻找并记录公开日统计的数据
            publicDayStas = '//*[@id="itemsList"]//a[@sort=\'PDY\']'  #公开日统计栏的XPath
            clickByXPath(browser, publicDayStas)  # 点击展开公开日统计栏
            publicYearsStas = '//*[@id="itemsList"]//ul[@sort=\'PDY\']/li'  # 公开日统计栏年(申请数)的XPath
            publicData = getYearNumByXPath(browser, publicYearsStas)  # 得到公开日统计数据

            #打印结果测试一下
            # print(applyData)
            # print(publicData)
            save([company_name,'申请日统计', '-'.join(applyData),'公开日统计', '-'.join(publicData)])

            # print('finished!')
        except Exception as e:
            print(e)
            print(company_name+'没有搜索到结果')
