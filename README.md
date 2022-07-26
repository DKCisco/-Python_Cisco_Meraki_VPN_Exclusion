# -Python_Cisco_Meraki_VPN_Exclusion
Automating the act of adding IP addresses to the VPN Exclusion for the 2022 Meraki Dashboard
This project was forked from https://github.com/oborys/Selenium_automation_Adding_Cisco_Meraki_VPN_exclusion_rules

Before running this script correctly you will need to do the following:
Install python on PC,
Open powershell ,
pip install keyring,
pip install autogui,
pip install pydirectinput,
pip install opencv-python,
pip install pillow,
pip install selenium,
Open python terminal and set your credentials with keyring.set_password("nameofacct", "username", "pass"),
The credentials variable calls this,
This script should be run on a single monitor,
Define the size for your resolution in the pyautogui.size(),
Disable mfa for your meraki account,
Download the ChromeDriver 104.0.5112.20, https://chromedriver.chromium.org/downloads,
Change the driver location for driver = webdriver.Chrome(executable_path = r"C:\Users\user\Documents\chromedriver.exe"),
Change the url for url_sd_wan_and_traffic_rules,
Change the url for driver.get('https://orgID.meraki.com/login/dashboard_login'),
Change the IPandPorts file to the IP addresses you need to add to the vpn exclusion save it as IPandPorts.xlsx in the same folder as this script,
You will only need pyautogui to click the protocol once if you are doing the same protocol for each IP,
The destination port will default to any,
Open visual studio and run/debug the code, don't scroll at all,
Let the script run until it opens the Security & SD-WAN tab,
You will need to change the x,y coordinates for pyautogui, this can be done w/ a program called faststone capture,
Install FastStone Caputre 9.3 and use the cntrl+PrtSc shortcut to get the coordinates for each button or feild you need clicked,
You will need to change the scroll value to fit your needs

