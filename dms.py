import os
from selenium import webdriver
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from datetime import datetime
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

import time

from senhas import dmsLogin, dmsPassword

# defining geckodriver and nav
path = r"C:\Users\guilh\AppData\Local\Programs\Python\Python39"
download_dir = r"D:\JAC\ATUALIZACAO_STATUS_GARANTIA"

# checking directory if file already exists
if os.path.exists(download_dir + "\DMS.xls"):
    os.remove(download_dir + "\DMS.xls")

fp = webdriver.FirefoxProfile()
fp.set_preference("browser.download.folderList",2)
fp.set_preference("browser.download.manager.showWhenStarting", False)
# fp.set_preference("browser.helperApps.alwaysAsk.force", False)
fp.set_preference("browser.helperApps.neverAsk.saveToDisk","application/octet-stream,application/vnd.ms-excel")
fp.set_preference("browser.download.dir", download_dir)
navegador = webdriver.Firefox(firefox_profile=fp)
navegador.maximize_window()

# main data
dms_login = dmsLogin
dms_password = dmsPassword

def logging_in_dms():
    username = navegador.find_element_by_id("userName")
    password = navegador.find_element_by_id("password")
    code_place = navegador.find_element_by_id("code")
    username.send_keys(dms_login)
    password.send_keys(dms_password)
    language = Select(navegador.find_element_by_id("LanguageType"))
    language.select_by_visible_text("English")

    # user input
    code = input("Enter verification code: ")
    code_place.send_keys(code)

    navegador.execute_script("submitform()")

def navigating_trough_dms(opt):
    def claim_edit_page():
        service_mgmt_button = navegador.find_element_by_id("1013")
        service_mgmt_button.click()
        time.sleep(3)

        navegador.switch_to.frame(navegador.find_element_by_id("menuIframe"))
        navegador.find_element_by_link_text("Warranty Claim Management").click()
        time.sleep(2)
        navegador.find_element_by_id("sd1").click()
        navegador.switch_to.default_content()

    def warranty_claim_tracking_page():
        service_mgmt_button = navegador.find_element_by_id("1013")
        service_mgmt_button.click()
        time.sleep(5)

        navegador.switch_to.frame(navegador.find_element_by_id("menuIframe"))
        navegador.find_element_by_link_text("Warranty Claim Management").click()
        time.sleep(5)
        navegador.find_element_by_id("sd4").click()
        navegador.switch_to.default_content()

    options = {
        0: claim_edit_page,
        1: warranty_claim_tracking_page
    }

    options[opt]()

def getting_warranty_tacking_report(start_date, end_date):
    navegador.switch_to.frame(navegador.find_element_by_id("inIframe"))

    # cleaning up audit approved dates
    # audit_approved_first = navegador.find_element_by_id("START_DATE")
    navegador.execute_script("document.getElementById('START_DATE').value =''")
    navegador.execute_script("document.getElementById('END_DATE').value =''")

    navegador.execute_script(f"document.getElementById('CREATE_DATE_START').value ='{start_date}'")
    navegador.execute_script(f"document.getElementById('CREATE_DATE_END').value ='{end_date}'")
    navegador.find_element_by_id("queryBtn").click()
    time.sleep(15)

    navegador.find_element_by_id("export").click()


# main logic
navegador.get("http://dms.jac.com.cn/JACINTERDMS/")
# time to click warning
time.sleep(10)

try:
    WebDriverWait(navegador, 3).until(EC.alert_is_present())
    alert = navegador.switch_to.alert
    alert.accept()
except TimeoutException:
    print("no alert")

# login logic
logging_in_dms()
time.sleep(10)

navigating_trough_dms(1)
time.sleep(10)
startingDate = "2021-01-01"
endDate = datetime.today().strftime('%Y-%m-%d')
getting_warranty_tacking_report(startingDate, endDate)
print("report downloaded!")

# wait for file download
time.sleep(5)

# file rename
downloaded_file = max([download_dir + "\\" + f for f in os.listdir(download_dir)],key=os.path.getctime)
os.rename(downloaded_file, download_dir + "\DMS.xls")
navegador.quit()
