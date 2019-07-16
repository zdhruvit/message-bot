import time
import os
import re
import png
import random
import selenium
from tqdm import tqdm
import urllib.parse
import xlrd
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

driver = webdriver.Firefox( executable_path=r'geckodriver.exe')
driver.get('http://web.whatsapp.com')
driver.set_window_size(1400,900)
time.sleep(1)
loc = ("users.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
num_users = int(sheet.nrows)

try:
    no = []
    msg = []
    for i in range(0,sheet.nrows):
        while True:
            try:
                x = int(sheet.cell_value(i, 0))
                pat = re.compile(r'^[0-9]\d{9}$')
                p = re.findall(pat,str(x))
                if p:
                    no.append(x)
                    msg.append(str(sheet.cell_value(i, 1)))
                    break
                else:
                    print('pattern dosent match! at row ' + str(i+1))
            except ValueError:
                print('Invalid Number at row ' + str(i+1))
                continue       
except ValueError:
    print('Invalid Input ! ')

time.sleep(20)
for i in tqdm(range(0,sheet.nrows)):
    elm = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[3]/div/div[1]')
    driver.execute_script("arguments['0'].innerHTML = '<a href=\"https://api.whatsapp.com/send?phone="+"+91"+str(no[i])+"&message="+urllib.parse.quote(msg[i])+"\"id=\"contact"+str(i+1)+"\">"+str(i+1)+"</a>';", elm)
    msgele = driver.find_element_by_xpath("/html/body/div[1]/div/div/div[3]/div/div[1]/a")
    msgele.click()
    time.sleep(2)
    focus = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
    for letter in msg[i]:
        time.sleep(random.randint(1, 3) / 10)
        if(letter == '\n'):
            focus.send_keys(Keys.SHIFT, Keys.ENTER)
        else:
            focus.send_keys(letter)
        
    focus.send_keys(Keys.ENTER)
    time.sleep(random.randint(60, 120))
    
print('done')
time.sleep(1)
dot = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[3]/div/header/div[2]/div/span/div[3]/div')
dot.click()
time.sleep(1)
logout = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[3]/div/header/div[2]/div/span/div[3]/span/div/ul/li[6]/div')
logout.click()
print('logged out')

driver.quit()