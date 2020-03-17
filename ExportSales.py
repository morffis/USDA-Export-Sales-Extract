#DATA PREPROCESSING
import pandas as pd #pip3 install pandas
import numpy as np #pip3 install numpy
from xlrd import open_workbook
from bs4 import BeautifulSoup
from urllib.request import urlopen
import csv
import datetime

#SQL SERVER
import pyodbc
import sqlalchemy
import urllib


#OS
from os import listdir
from os import remove
from os.path import isfile, join, splitext

#SELENIUM
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains 
from selenium.webdriver.common.keys import Keys

import time

import pyautogui as GUI


def delete_all_files_in_folder(folder):
    for f in listdir(folder):
        if isfile(join(folder, f)):
            remove(join(folder, f))



optionsEditted = Options();
optionsEditted.set_preference("browser.download.folderList",2);
optionsEditted.set_preference("browser.download.manager.showWhenStarting", False)


#DOWNLOAD FOLDER LOCATION

download_path = 'Export Sales.csv'

output_path = download_path

#Clear Output and Download Folders
delete_all_files_in_folder(output_path)
delete_all_files_in_folder(download_path)

optionsEditted.set_preference("browser.download.dir",download_path);

mime_types = [
    'text/plain', 
    'application/vnd.ms-excel', 
    'text/csv', 
    'application/csv', 
    'text/comma-separated-values', 
    'application/download', 
    'application/octet-stream', 
    'binary/octet-stream', 
    'application/binary', 
    'application/x-unknown'
]

optionsEditted.set_preference("browser.helperApps.neverAsk.saveToDisk", ",".join(mime_types));


binary = FirefoxBinary("C:/Program Files/Mozilla Firefox/firefox.exe")

cap = DesiredCapabilities().FIREFOX
cap["marionette"] = True


#PARAMETERS FOLDER LOCATION

paramsPath = r'./QueryParams'

paramFiles = []

for f in listdir(paramsPath):
    #if it is a file..
    if(isfile(join(paramsPath,f))): 
        #add it to the list
        paramFiles.append(f)

print(paramFiles)


with open(paramsPath+'/'+paramFiles[0],'r') as f:
    commodities = [line.strip() for line in f]

print(commodities)


with open(paramsPath+'/'+paramFiles[1],'r') as f:
    countries = [line.strip() for line in f]
print(countries)


#Start browser and go to query page
QUERY_URL = 'https://apps.fas.usda.gov/esrquery/esrq.aspx/'
#RESOURCE FILE LOCATION
browser = webdriver.Firefox(capabilities=cap, executable_path=r"N:\BI\BI New Database\02 Automation Routines\00. Resources\SeleniumWebDrivers\geckodriver.exe", firefox_binary=binary, options=optionsEditted)

browser.get(QUERY_URL)


commoditySelector = browser.find_element_by_id("ctl00_MainContent_lbCommodity")
commodityOptions = commoditySelector.find_elements_by_tag_name('option')

#then, for each param listed in commodities file click if the option matches
for commodityOption in commodityOptions:
    for commodityParam in commodities:
        if( commodityOption.get_attribute('innerHTML') == commodityParam or int(commodityOption.get_attribute('value')) == 901):
            browser.execute_script("arguments[0].selected='selected';",commodityOption)
            break
        else:
            browser.execute_script("arguments[0].selected='';",commodityOption)


countrySelector = browser.find_element_by_id("ctl00_MainContent_lbCountry")
countryOptions = countrySelector.find_elements_by_tag_name('option')

#then, for each param listed in commodities file click if the option matches
for countryOption in countryOptions:
    for countryParam in countries:
        if( countryOption.get_attribute('innerHTML') == countryParam ):
            browser.execute_script("arguments[0].selected='selected';",countryOption)
            break
        else:
            browser.execute_script("arguments[0].selected='';",countryOption)



lastDateString = '01/01/2020'
lastDate = str(lastDateString.strftime('%m/%d/%Y'))

today = str(time.strftime('%m/%d/%Y'))
print (today)


#change Start and End Dates
startDate = browser.find_element_by_id("ctl00_MainContent_tbStartDate")
endDate = browser.find_element_by_id("ctl00_MainContent_tbEndDate")

startDate.clear()
startDate.send_keys(lastDate)


endDate.clear()
endDate.send_keys(today)


outputFormat = browser.find_element_by_id('ctl00_MainContent_rblOutputType_2')
sbtBtn = browser.find_element_by_id('ctl00_MainContent_btnSubmit')


outputFormat.click()
sbtBtn.click()


browser.quit()