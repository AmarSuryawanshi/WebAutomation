from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
import sys
from glob import glob
import os

chromeOptions = Options()
chromeOptions.add_experimental_option("prefs", {
                                      "download.default_directory": "D:\Computer Aplications Storage(Users)\Documents\Paladin_Backup\paladin_avail"})

deleteFiles = input(
    "Please, press (Y)es to delete all present .csv and .xlsx files. ")
if (deleteFiles == 'Y' or deleteFiles == 'y'):
    directory = 'D:/Computer Aplications Storage(Users)/Documents/Paladin_Backup/paladin_avail'
    os.chdir(directory)
    files = glob('*.cgi')
    for filename in files:
        os.unlink(filename)

    directory1 = 'D:/Computer Aplications Storage(Users)/Documents/Paladin_Backup/paladin_excel'
    os.chdir(directory1)
    files1 = glob('*.xlsx')
    for filename in files1:
        os.unlink(filename)


web = webdriver.Chrome(options=chromeOptions)
web.get('http://127.0.0.1:5500/index.html')

# for Hosts

type = web.find_element_by_xpath('//*[@id="type"]')
drp = Select(type)

drp.select_by_visible_text('Host')

submit1 = web.find_element_by_xpath('//*[@id="btnClick"]')
submit1.click()

submit2 = web.find_element_by_xpath('/html/body/header/div/form/div/button')
submit2.click()

CSVradio = web.find_element_by_xpath('//*[@id="checkCSV"]')
CSVradio.click()

submit3 = web.find_element_by_xpath('/html/body/header/div/form/div/button/a')
submit3.click()

home = web.find_element_by_xpath('/html/body/header/nav/ul/li/a')
home.click()


# for services

type = web.find_element_by_xpath('//*[@id="type"]')
drp = Select(type)

drp.select_by_visible_text('Service')

submit1 = web.find_element_by_xpath('//*[@id="btnClick"]')
submit1.click()

submit2 = web.find_element_by_xpath('/html/body/header/div/form/div/button')
submit2.click()

CSVradio = web.find_element_by_xpath('//*[@id="checkCSV"]')
CSVradio.click()

submit3 = web.find_element_by_xpath('/html/body/header/div/form/div/button/a')
submit3.click()

time.sleep(2)

web.close()
time.sleep(0.5)
print("Deleted and downloaded new files Successfully...")
time.sleep(0.5)
print("Converting files to xlsx, Please wait...")

df = pd.read_csv('E:/Studies/MyProjects/ATOSAutomation/atos web/Host/avail.cgi', sep=',\s+',
                 engine='python').replace('"', '', regex=True)
writer = pd.ExcelWriter(
    'D:/Computer Aplications Storage(Users)/Documents/Paladin_Backup/paladin_excel/PRD_HOST.xlsx')
df.to_excel(writer, index=False)
writer.save()

df = pd.read_csv('E:/Studies/MyProjects/ATOSAutomation/atos web/Service/avail.cgi', sep=',\s+',
                 engine='python').replace('"', '', regex=True)
writer = pd.ExcelWriter(
    'D:/Computer Aplications Storage(Users)/Documents/Paladin_Backup/paladin_excel/PRD_SERVICES.xlsx')
df.to_excel(writer, index=False)
writer.save()

print("\nProcess has been completed Successful. You can close this window.")
print("Thank you for using this app!!! Developed by Amruta Suryawanshi @2021\n")
time.sleep(2)
sys.exit()
