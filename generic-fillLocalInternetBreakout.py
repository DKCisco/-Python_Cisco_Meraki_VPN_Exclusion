"""
Before running this script correctly you will need to do the following:
Install python on PC
Open powershell 
pip install keyring
pip install autogui
pip install selenium
pip install openpyxl
Open python terminal and set your credentials with keyring.set_password("nameofacct", "username", "pass")
The credentials variable calls this
This script should be run on a single monitor
Define the size for your resolution in the pyautogui.size()
Disable mfa for your meraki account
Download the ChromeDriver 104.0.5112.20, https://chromedriver.chromium.org/downloads
Change the driver location for driver = webdriver.Chrome(executable_path = r"C:\Users\user\Documents\chromedriver.exe")
Change the url for url_sd_wan_and_traffic_rules
Change the url for driver.get('https://orgID.meraki.com/login/dashboard_login')
Change the IPandPorts file to the IP addresses you need to add to the vpn exclusion save it as IPandPorts.xlsx in the same folder as this script
You will only need pyautogui to click the protocol once if you are doing the same protocol for each IP
The destination port will default to any
Open visual studio and run/debug the code, don't scroll at all
Let the script run until it opens the Security & SD-WAN tab
You will need to change the x,y coordinates for pyautogui, this can be done w/ a program called faststone capture
Install FastStone Caputre 9.3 and use the cntrl+PrtSc shortcut to get the coordinates for each button or feild you need clicked
You will need to change the scroll value to fit your needs
"""

from lib2to3.pgen2.driver import Driver
import openpyxl
from telnetlib import IP
import keyring
import pyautogui
import pyautogui as py
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.action_chains import ActionChains

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

import selenium.webdriver
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.common.exceptions import TimeoutException
import configparser

credentials = keyring.get_password('Meraki', 'username')
url_sd_wan_and_traffic_rules = 'https://x.meraki.com/your_template_here/n/W1niQavc/manage/configure/traffic_shaping/0'
dashboardPass = credentials
email = 'yourlogin@email.com'

driver = webdriver.Chrome(executable_path = r"C:\Users\user\Documents\chromedriver.exe") #Opens the Chrome driver from your previous download
driver.maximize_window()

pyautogui.size()
(1920, 1080)

driver.get('https://orgID.meraki.com/login/dashboard_login') #Opens login page
username = driver.find_element(By.ID, "email")  #Finds the email field
username.send_keys(email)   #Inserts email pulled from the email variable above
driver.find_element(By.ID, "next-btn").click()  #Clicks the next button


password = driver.find_element(By.ID, "password")   #Finds the password field
password.send_keys(dashboardPass)   #Enters password pulled from variable dashboardPass that pulls from credentials file
driver.find_element(By.ID, "login-btn").click() #Clicks the login button

# wait the ready state to be complete
WebDriverWait(driver=driver, timeout=10).until(
    lambda x: x.execute_script("return document.readyState === 'complete'")
)
error_message = "Incorrect username or password."
#get the errors (if there are)
errors = driver.find_elements(By.CLASS_NAME, "flash-error")

if any(error_message in e.text for e in errors):
    print("[!] Login failed")
else:
    print("[+] Login successful")

driver.get(url_sd_wan_and_traffic_rules)
time.sleep(3)
pyautogui.click(1400, 900) #Click in whitespace
print('Clicked whitespace')
pyautogui.scroll(-1700) #Scroll down
time.sleep(1)
print('Scrolled once')
button = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//a[@data-toggle='dropdown']")))
button.click()
time.sleep(3)
pyautogui.click(760, 530) #Protocol dropdown
time.sleep(3)
pyautogui.click(760, 650) #Selecting any
time.sleep(3)

# load excel with its path
wrkbk = openpyxl.load_workbook("IPandPorts.xlsx")
  
sh = wrkbk.active
  
# iterate through excel and display data
for row in sh.iter_rows(min_row=1, min_col=1, max_row=302, max_col=1): #Your max row will be the last ip row's number in the IPandPorts.xlsx
    for cell in row:
        time.sleep(2)
        print(cell.value, end=" ")
        pyautogui.click(760, 590) #Destination IP dropdown
        print('Just clicked the destination IP dropdown.')
        time.sleep(5)
        pyautogui.click(770, 650) #IP address feild
        print('Just clicked the IP address feild.')
        time.sleep(5)
        pyautogui.typewrite(cell.value)
        print('Just entered the IP address')
        time.sleep(5)
        pyautogui.click(780, 680) #Add IP address entered
        print('Just clicked Add for the IP address just entered')
        time.sleep(10)
        #pyautogui.click(950, 700) #Click add expression
        print('Just clicked to activate window')
        element = driver.find_element(By.XPATH, "//tbody/tr/td[@class='cfgo']/div[@id='vpn_exclusion_div']/div[@class='tclass']/div[@class='traffic_shaper']/div[@id='rule_menu_vpn_exclusion_shaper']/div[@class='pillbox-dropdown dropdown-widget iwan_l7']/div[@class='category sub custom_pane iwan_l7']/div[@class='custom expressions']/div[@class='custom custom-net expression iwan_l7']/div[@class='iwan_l7 descriptor']/button[1]")
        driver.execute_script("arguments[0].click();", element)
        #WebDriverWait(driver, 40).until(EC.element_to_be_clickable((By.XPATH, "//tbody/tr/td[@class='cfgo']/div[@id='vpn_exclusion_div']/div[@class='tclass']/div[@class='traffic_shaper']/div[@id='rule_menu_vpn_exclusion_shaper']/div[@class='pillbox-dropdown dropdown-widget iwan_l7']/div[@class='category sub custom_pane iwan_l7']/div[@class='custom expressions']/div[@class='custom custom-net expression iwan_l7']/div[@class='iwan_l7 descriptor']/button[1]"))).click()
    print()

pyautogui.click(1760, 1000) #Click save
print('Just clicked save')
time.sleep(8)
print('Script complete')

