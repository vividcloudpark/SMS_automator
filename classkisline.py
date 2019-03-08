import os.path
import time
import json
from bs4 import BeautifulSoup
from collections import OrderedDict, defaultdict
import numpy as np
import pandas as pd
import collections
import traceback
import datetime
import os
import sys
import pywinauto

from selenium import webdriver
from selenium.webdriver.ie.options import Options
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait # available since 2.4.0
from selenium.webdriver.support import expected_conditions as EC # available since 2.26.0
from selenium.webdriver.common.keys import Keys

class JsonDump:
    def __init__(self):
        dir = os.path.dirname(os.path.abspath(__file__))
        jsonfile = dir + r'\\asset\\static\\login_infomation.json'
        with open(jsonfile) as f:
            login_data = json.load(f)
        self.login_data = login_data

    def get_login_info(self):
        self.id = self.login_data['ID']
        self.password = self.login_data['Password']

    def get_MU_info(self):
        muinfo = {'MU' : self.login_data['MU'],
                  'Countrycode' : self.login_data['Countrycode'],
                  'inumber' : self.login_data['inumber']}
        return muinfo

class SelDriver(JsonDump):
    def __init__(self):
        super(SelDriver, self).__init__()
        opts = Options()
        opts.ignore_protected_mode_settings = True
        opts.ignore_zoom_level = True
        dir = os.path.dirname(os.path.abspath(__file__))
        path = dir + r'\\asset\\selenium\\chromedriver.exe'
        self.driver = webdriver.Chrome(executable_path=path, chrome_options=opts)
        self.access()
        self.login()

    def control(self):
        return self.driver

    def access(self):
        pageurl = "http://intraportal.gachon.ac.kr/intraportal/main.jsp"
        self.driver.get(pageurl)

    def put_id(self):
        self.driver.find_element_by_name('uid').click()
        self.driver.find_element_by_name('uid').send_keys('manage')

    def put_password(self):
        self.driver.find_element_by_name('password').click()
        self.driver.find_element_by_name('password').send_keys('716rkcjs!')

    def login(self):
            print('****LOGIN****')
            self.put_id()
            self.put_password()
            self.driver.find_element_by_xpath('//*[@id="content"]/div[2]/div/div[2]/form/div/input').click()

if __name__ == "__main__":
    seldriver = SelDriver()
    driver = seldriver.control()
    driver.find_element_by_xpath('//*[@id="Map3"]/area[3]').click()
    time.sleep(1)
    window_after = driver.window_handles[1]
    driver.switch_to.window(window_after)
    driver.find_element_by_xpath('//*[@id="mtm_layerpop"]/div[2]/div/a').click()
    driver.find_element_by_xpath('//*[@id="header"]/div/ul/li[2]/a').click()
    driver.find_element_by_xpath('//*[@id="mtm_layerpop"]/div[2]/div/a').click()
    driver.switch_to.window(driver.window_handles[-1])
    filepath = 'D:\workspace\SMS_automator\send_sample_general.xlsx'
    driver.find_element_by_xpath('//*[@id="uploadForm"]/div/table/tbody/tr/td/div').click()


    def findapp():
        app = pywinauto.application.Application()
        print(app)
        # mainWindow = app['Envoi du fichier'] # main windows' title
        # ctrl=mainWindow['Edit']
        # mainWindow.SetFocus()
        # ctrl.ClickInput()
        # ctrl.TypeKeys(Name_of_File)
        # ctrlBis = mainWindow['Ouvrir'] # open file button
        # ctrlBis.ClickInput()
    findapp()

    driver.find_element_by_xpath('//*[@id="uploadForm"]/div/table/tbody/tr/td/div').send_keys(filepath)
    driver.find_element_by_xpath('//*[@id="editor"]').click()
    print('''
    **** 필요한 문자를 입력해주세요 ****
    입력후 '전송파일 업로드하기를 클릭해주세요'
    ''')
