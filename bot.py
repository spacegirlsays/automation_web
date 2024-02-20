from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium import webdriver
from pathlib import Path
import pandas as pd
import time
import os.path

browser_options = Options()
folder_path = str(Path(__file__).parent)
browser_options.add_experimental_option("prefs", {
  "download.default_directory": folder_path
  })

#Iniciate browser
browser = webdriver.Chrome(options=browser_options)
browser.maximize_window()

#Opening Site
browser.get("https://rpachallenge.com/")

#Download Excel File
file_path = Path(f"{folder_path}\\challenge.xlsx")
time.sleep(4)

if not os.path.isfile(file_path):
  browser.find_element(By.XPATH, "//a[@class=' col s12 m12 l12 btn waves-effect waves-light uiColorPrimary center']").click()
  time.sleep(4)


#Reading Excel
excel_file = pd.read_excel("challenge.xlsx", header=None, skiprows=1)
for row in excel_file.itertuples():
    #First Name
    browser.find_element(By.XPATH, "//input[@ng-reflect-name='labelFirstName']").send_keys(row[1])
    time.sleep(1)
    #Last Name
    browser.find_element(By.XPATH, "//input[@ng-reflect-name='labelLastName']").send_keys(row[2])
    time.sleep(1)
    #Company Name
    browser.find_element(By.XPATH, "//input[@ng-reflect-name='labelCompanyName']").send_keys(row[3])
    time.sleep(1)
    #Role
    browser.find_element(By.XPATH, "//input[@ng-reflect-name='labelRole']").send_keys(row[4])
    time.sleep(1)
    #Address
    browser.find_element(By.XPATH, "//input[@ng-reflect-name='labelAddress']").send_keys(row[5])
    time.sleep(1)
    #Email
    browser.find_element(By.XPATH, "//input[@ng-reflect-name='labelEmail']").send_keys(row[6])
    time.sleep(1)
    #Phone
    browser.find_element(By.XPATH, "//input[@ng-reflect-name='labelPhone']").send_keys(row[7])
    time.sleep(1)
    browser.find_element(By.XPATH, "//input[@type='submit']").click()
    time.sleep(1)

browser.close()