### Import Dependencies
import os
import pandas as pd
import numpy as np
from bs4 import BeautifulSoup as soup
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pyautogui
from datetime import datetime
import shutil
import ERP_Credentials

### Initiate Counter
start_time = datetime.now()

### Remove Existing Files from the Specified Path
if os.path.isfile("C:\Downloads\DMR NC Report.xls") == True:
    os.remove("C:\Downloads\DMR NC Report.xls")
if os.path.isfile("C:\Downloads\DMR NC Report (1).xls") == True:
    os.remove("C:\Downloads\DMR NC Report (1).xls")
if os.path.isfile("C:\Downloads\DMR List.xls") == True:
    os.remove("C:\Downloads\DMR List.xls")
if os.path.isfile("C:\Downloads\CAR Report - 2019.xls") == True:
    os.remove("C:\Downloads\CAR Report - 2019.xls")
if os.path.isfile("C:\Downloads\CAR Report - 2020.xls") == True:
        os.remove("C:\Downloads\CAR Report - 2020.xls")
else:
    print("No existing files to remove. Ready to initialize the script.")

### Initialize Automated Browser
driver = webdriver.Chrome(executable_path='C:\\Users\MXC04\Downloads\chromedriver\chromedriver')
driver.execute_script("document.body.style.zoom='zoom 125%'")
driver.get('ERP SYSTEM URL')
driver.maximize_window()
time.sleep(1)

### Apply Credentials and Log in to eTraveler
username = driver.find_element_by_xpath("//input[@type='text'][@class='searchBlockInput']")
username.send_keys(ERP_Credentials.Username)
password = driver.find_element_by_xpath("//input[@type='password'][@class='searchBlockInput']")
password.send_keys(ERP_Credentials.Password)
password.send_keys(Keys.ENTER)
time.sleep(1)

### Click on the NCR Report
ncrreport = driver.find_element_by_id("rp3")
ncrreport.click()

### Click on Excel Report via Location by Pixels
time.sleep(2)
#pyautogui.position()
pyautogui.click(252, 322)

### Click on Excel Report via Location by Pixels
time.sleep(2)
pyautogui.click(914, 548)
pyautogui.click(252, 322)
time.sleep(3)
driver.switch_to.window(driver.window_handles[0])

### Click on the NCR Report
ncrreport = driver.find_element_by_id("rp1")
ncrreport.click()
time.sleep(1)
pyautogui.click(853, 517)
time.sleep(1)
pyautogui.moveTo(891, 590)
pyautogui.click(891, 590)
time.sleep(1)
pyautogui.click(254, 327)

### Switch to the Main tab
driver.switch_to.window(driver.window_handles[0])

### Click on Excel Report via Location by Pixels
pyautogui.click(109, 544)
time.sleep(1)
pyautogui.moveTo(165, 844)
time.sleep(1)
pyautogui.click(165, 844)
time.sleep(2)
image = driver.find_elements_by_css_selector('img[src="/images/excel.png"]')
image[0].click() #Click on 2020 CAR Report
time.sleep(1)
image[1].click() #Click on 2019 CAR Report
time.sleep(1)
driver.quit()

### Read & Convert HTML to Excel File
df = pd.read_html("C:\Downloads\DMR NC Report.xls")
df1 = df[0]
df2 = df1.to_excel("DMR NCR.xlsx")

### Read & Convert HTML to Excel File
df = pd.read_html("C:\Downloads\DMR NC Report (1).xls")
df1 = df[0]
df2 = df1.to_excel("DMR NCR(SERIAL).xlsx")
print("Conversion of DMR NCR file has been completed for both hollistic and serialized formats.")

### Read & Convert HTML to Excel File
df = pd.read_html("C:\Downloads\DMR List.xls")
df1 = df[0]
df2 = df1.to_excel("DMR NCR(CLOSED).xlsx")
print("Conversion of Closed DMRs has been completed.")

### Read & Convert HTML to Excel File
df = pd.read_html("C:\Downloads\CAR Report - 2019.xls")
df1 = df[0]
df2 = df1.to_excel("CAR Report.xlsx")
df3 = pd.read_html("C:\Downloads\CAR Report - 2020.xls")
df4 = df3[0]
df5 = df4.to_excel("CAR Report 2.xlsx")
print("Conversion of CAR Report has been completed.")

# Allocate new Excel files to appropriate file path
if os.path.isfile("C:\\Users\MXC04\Desktop\dist\AU\DMR NCR.xlsx") == True:
    shutil.move(os.path.join("C:\\Users\MXC04\Desktop\dist\AU", "DMR NCR.xlsx"), os.path.join("C:\\Users\MXC04\Desktop\Python Projects", "DMR NCR.xlsx"))
    shutil.move(os.path.join("C:\\Users\MXC04\Desktop\dist\AU", "DMR NCR(SERIAL).xlsx"), os.path.join("C:\\Users\MXC04\Desktop\Python Projects", "DMR NCR(SERIAL).xlsx"))
    shutil.move(os.path.join("C:\\Users\MXC04\Desktop\dist\AU", "DMR NCR(CLOSED).xlsx"), os.path.join("C:\\Users\MXC04\Desktop\Python Projects", "DMR NCR(CLOSED).xlsx"))
    shutil.move(os.path.join("C:\\Users\MXC04\Desktop\dist\AU", "CAR Report.xlsx"), os.path.join("C:\\Users\MXC04\Desktop\Python Projects", "CAR Report.xlsx"))
    shutil.move(os.path.join("C:\\Users\MXC04\Desktop\dist\AU", "CAR Report 2.xlsx"), os.path.join("C:\\Users\MXC04\Desktop\Python Projects", "CAR Report 2.xlsx"))
else:
    print('No Files in Desktop/Dist/AU')

if os.path.isfile("C:\\Users\MXC04\Desktop\DMR NCR.xlsx") == True:
    shutil.move(os.path.join("C:\\Users\MXC04\Desktop", "DMR NCR.xlsx"), os.path.join("C:\\Users\MXC04\Desktop\Python Projects", "DMR NCR.xlsx"))
    shutil.move(os.path.join("C:\\Users\MXC04\Desktop", "DMR NCR(SERIAL).xlsx"), os.path.join("C:\\Users\MXC04\Desktop\Python Projects", "DMR NCR(SERIAL).xlsx"))
    shutil.move(os.path.join("C:\\Users\MXC04\Desktop", "DMR NCR(CLOSED).xlsx"), os.path.join("C:\\Users\MXC04\Desktop\Python Projects", "DMR NCR(CLOSED).xlsx"))
    shutil.move(os.path.join("C:\\Users\MXC04\Desktop", "CAR Report.xlsx"), os.path.join("C:\\Users\MXC04\Desktop\Python Projects", "CAR Report.xlsx"))
    shutil.move(os.path.join("C:\\Users\MXC04\Desktop", "CAR Report 2.xlsx"), os.path.join("C:\\Users\MXC04\Desktop\Python Projects", "CAR Report 2.xlsx"))
else:
    print('No Files in Desktop')

if os.path.isfile("C:\\Users\MXC04\Desktop\Py Scripts\dist\AU\DMR NCR.xlsx") == True:
    shutil.move(os.path.join("C:\\Users\MXC04\Desktop\Py Scripts\dist\AU", "DMR NCR.xlsx"), os.path.join("C:\\Users\MXC04\Desktop\Python Projects", "DMR NCR.xlsx"))
    shutil.move(os.path.join("C:\\Users\MXC04\Desktop\Py Scripts\dist\AU", "DMR NCR(SERIAL).xlsx"), os.path.join("C:\\Users\MXC04\Desktop\Python Projects", "DMR NCR(SERIAL).xlsx"))
    shutil.move(os.path.join("C:\\Users\MXC04\Desktop\Py Scripts\dist\AU", "DMR NCR(CLOSED).xlsx"), os.path.join("C:\\Users\MXC04\Desktop\Python Projects", "DMR NCR(CLOSED).xlsx"))
    shutil.move(os.path.join("C:\\Users\MXC04\Desktop\Py Scripts\dist\AU", "CAR Report.xlsx"), os.path.join("C:\\Users\MXC04\Desktop\Python Projects", "CAR Report.xlsx"))
    shutil.move(os.path.join("C:\\Users\MXC04\Desktop\Py Scripts\dist\AU", "CAR Report 2.xlsx"), os.path.join("C:\\Users\MXC04\Desktop\Python Projects", "CAR Report 2.xlsx"))
else:
    print('No Files in Desktop/Py Scripts/dist/AU')

if os.path.isfile("C:\\Users\MXC04\Desktop\Py Scripts\DMR NCR.xlsx") == True:
    shutil.move(os.path.join("C:\\Users\MXC04\Desktop\Py Scripts", "DMR NCR.xlsx"), os.path.join("C:\\Users\MXC04\Desktop\Python Projects", "DMR NCR.xlsx"))
    shutil.move(os.path.join("C:\\Users\MXC04\Desktop\Py Scripts", "DMR NCR(SERIAL).xlsx"), os.path.join("C:\\Users\MXC04\Desktop\Python Projects", "DMR NCR(SERIAL).xlsx"))
    shutil.move(os.path.join("C:\\Users\MXC04\Desktop\Py Scripts", "DMR NCR(CLOSED).xlsx"), os.path.join("C:\\Users\MXC04\Desktop\Python Projects", "DMR NCR(CLOSED).xlsx"))
    shutil.move(os.path.join("C:\\Users\MXC04\Desktop\Py Scripts", "CAR Report.xlsx"), os.path.join("C:\\Users\MXC04\Desktop\Python Projects", "CAR Report.xlsx"))
    shutil.move(os.path.join("C:\\Users\MXC04\Desktop\Py Scripts", "CAR Report 2.xlsx"), os.path.join("C:\\Users\MXC04\Desktop\Python Projects", "CAR Report 2.xlsx"))
else:
    print('No Files in Desktop/Py Scripts')

print("Files have been moved to their appropriate paths.")
print("All tasks have been completed.")

### End Timer $ Print Runtime
end_time = datetime.now()

print('Script Completed in: {}'.format(end_time - start_time))
input('Press ENTER to close the command prompt.')