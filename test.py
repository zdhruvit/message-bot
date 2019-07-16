import time
import os
import re
import png
import selenium
from tqdm import tqdm
import xlrd
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

# options = Options()
# options.set_headless(True)
driver = webdriver.Firefox( executable_path=r'geckodriver.exe')
driver.get('http://web.whatsapp.com')
delay = 3 # seconds
driver.set_window_size(1400,900)
time.sleep(1)
token = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div[1]/div/div[2]/div').get_attribute('data-ref')
img = pyqrcode.create(token)
img.png('QRcode.png',scale=5)
loc = ("users.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
num_users = int(sheet.nrows)

try:
    no = []
    msg = []
    for i in range(0,sheet.nrows-1):
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

for i in tqdm(range(0,sheet.nrows-1)):
    try:
        elm = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[3]/div/div[1]')))
        driver.execute_script("arguments['0'].innerHTML = '<a href=\"https://api.whatsapp.com/send?phone="+"+91"+str(no[i])+"&message="+str(msg[i])+"id=\"contact"+str(i+1)+">"+str(i+1)+"</a>';", elm)
        msgele = driver.find_element_by_xpath("/html/body/div[1]/div/div/div[3]/div/div[1]/a")
        msgele.click()
        time.sleep(1)
        focus = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
        focus.send_keys(msg)
        focus.send_keys(Keys.ENTER)
    except TimeoutException:
        print("Loading took too much time!")
    

print('done')
time.sleep(1)
dot = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[3]/div/header/div[2]/div/span/div[3]/div')
dot.click()
time.sleep(1)
logout = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[3]/div/header/div[2]/div/span/div[3]/span/div/ul/li[6]/div')
logout.click()
print('logged out')

if os.path.exists('QRcode.png'):
    os.remove('QRcode.png')


driver.quit()