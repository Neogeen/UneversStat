import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException, UnexpectedAlertPresentException
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep

driver = webdriver.Firefox()
path = 'ВУЗы и ОП.xlsx'
table = pd.read_excel(path, sheet_name='Все')

Vuzi = table.columns.to_list()
Op = table['ГОП'].to_list()
spec = table['Спец'].to_list()
l = 0
# spec = table['Спец'].to_list()
try:
    for i in range(1000):
        driver.get('https://univision.kz/edu-program/group/')
        sleep(1)
        driver.find_element_by_xpath(f"//*[contains(text(), '{Op[i].split(' ')[0]}')]").click()
        sleep(1)
        try:
            driver.find_element_by_xpath('//a[@href="#eps"]').click()
        except:
            print(f'net spec {i}')
            continue
        sleep(1)

        try:
            for t in range(1, 1000):
                # table.loc[l, 'ГОП'] = Op[i]
                # table.loc[l, 'ОП'] = driver.find_element_by_xpath(f'/html/body/main/div/div[4]/div[2]/div/a[{t}]/div').text
                # table.loc[l, 'ВУЗ'] = driver.find_element_by_xpath(f'/html/body/main/div/div[4]/div[2]/div/a[{t}]/p').text
                # l += 1
                # table._set_value(i, 'ГОП', Op[i])
                textop = driver.find_element_by_xpath(f'/html/body/main/div/div[4]/div[2]/div/a[{t}]/div').text
                vuz = driver.find_element_by_xpath(f'/html/body/main/div/div[4]/div[2]/div/a[{t}]/p').text
                # print(spec[i])
                # print(textop)
                if spec[i] in textop:
                    table._set_value(i, vuz, 1)
                    # print(2)
        except:
            end = 0

        #/html/body/main/div/div[4]/div[2]/div/a[{t}]/div
        #/html/body/main/div/div[4]/div[2]/div/a[2]/div
        # break
        # j = ''
        # for k in Vuzi:
        #     if j == k:
        #         table._set_value(i, k, 1)
        # # for j in range(120):
    table.to_excel('out2.xlsx', index=False)
except:
    table.to_excel('out2.xlsx', index=False)
