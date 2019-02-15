from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import os
import time
from time import sleep
from selenium.webdriver.firefox.firefox_profile \
import FirefoxProfile
from pynput.keyboard import Key, Controller




# National Horticulture Board URL
url = 'http://nhb.gov.in/OnlineClient/categorywiseallvarietyreport.aspx?enc=3ZOO8K5CzcdC' \
      '/Yq6HcdIxJ4o5jmAcGG5QGUXX3BlAP4= '



# You need to install geckodriver first in order to run the automation.
# Change the path to where you have geckodriver installed to run Firefox
path = "/home/chinmay/Downloads/geckodriver"




# Set FireFox Profile
profile = FirefoxProfile()
profile.set_preference("browser.download.panel.shown", False)
profile.set_preference("browser.helperApps.neverAsk.openFile","text/plain, application/vnd.ms-excel, text/csv, text/comma-separated-values, application/octet-stream, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.ms-excel, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
profile.set_preference("browser.download.folderList", 2)
profile.set_preference("browser.download.dir", "/home/chinmay/Downloads/")



# Create a webdriver
driver = webdriver.Firefox(firefox_profile=profile, executable_path=path)
driver.get(url)



# Enter Inputs
date = driver.find_element_by_id("ctl00_ContentPlaceHolder1_txtdate")
date.send_keys('10/12/2018')

category = driver.find_element_by_id("ctl00_ContentPlaceHolder1_drpCategoryName")
category.send_keys('FLOWERS')
# Select All States
states = driver.find_element_by_id("ctl00_ContentPlaceHolder1_btSelectAll").click()




# Submit the Details by button click
click = WebDriverWait(driver,40).until(EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnSearch")))
ActionChains(driver).move_to_element(click).click().perform()



# Page loads new URL
# Explicitly wait until the Buttons are visible and clickable
element = WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnExcel"]'))).click()



# Manual Input on Download Manager to Download the Excel File to Downloads Folder
keyboard = Controller()
keyboard.press(Key.down)
keyboard.release(Key.down)
keyboard.press(Key.enter)
keyboard.release(Key.enter)

# You can find the Downloaded file in the Downloads folder.
# If you want to download it in the current working directory you can use the (os.getcwd()) to download the file there.


