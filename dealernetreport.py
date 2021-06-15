import os
from selenium import webdriver
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from datetime import datetime
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException

import time
from senhas import dealernetLogin, dealernetPassword

# defining geckodriver and nav
path = r"C:\Users\guilh\AppData\Local\Programs\Python\Python39"
download_dir = r"D:\JAC\ATUALIZACAO_STATUS_GARANTIA"

# checking directory if file already exists
if os.path.exists(download_dir + "\dealernet.xlsx"):
    os.remove(download_dir + "\dealernet.xlsx")

fp = webdriver.FirefoxProfile()
fp.set_preference("browser.download.folderList",2)
fp.set_preference("browser.download.manager.showWhenStarting", False)
# fp.set_preference("browser.helperApps.alwaysAsk.force", False)
fp.set_preference("browser.helperApps.neverAsk.saveToDisk","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
fp.set_preference("browser.download.dir", download_dir)
navegador = webdriver.Firefox(firefox_profile=fp)
navegador.maximize_window()

# defining login and password from Dealernet
meu_login_dealernet = dealernetLogin
minha_senha_dealernet = dealernetPassword


# login function
def login_in_dealernet():
    # finding where to put data
    login = navegador.find_element_by_id("vUSUARIO_IDENTIFICADORALTERNATIVO")
    senha = navegador.find_element_by_id("vUSUARIOSENHA_SENHA")
    confirm_button = navegador.find_element_by_id("IMAGE3")    

    # send keys and confirming
    login.send_keys(meu_login_dealernet)
    senha.send_keys(minha_senha_dealernet)
    confirm_button.send_keys(Keys.ENTER)

# change to iframe function
def listing_iframes(id_searched):
    number_of_frames = len(navegador.find_elements_by_tag_name("iframe"))
    array_tries = 0
    for x in range(number_of_frames):
        # print(array_tries)
        navegador.switch_to.frame(x)
        try:
            os_num_placeholder = navegador.find_element_by_id(id_searched)
            navegador.switch_to.default_content()
            array_tries = x
            return array_tries

        except NoSuchElementException:
            navegador.switch_to.default_content()
            # print("no OS num placeholder in " + str(x))

    return print("no frame found")

# getting report logic
def getting_os_report(enddate):
    oficina_button = navegador.find_element_by_id("W5|_253_|Oficina")
    oficina_button.click()
    time.sleep(3)
    first_hover_over = navegador.find_element_by_id("x-menu-el-W5|_305_|Relatórios")
    ActionChains(navegador).move_to_element(first_hover_over).perform()
    time.sleep(2)
    second_hover_over = navegador.find_element_by_id("W5|_373_|O.S.")
    ActionChains(navegador).move_to_element(second_hover_over).perform()
    time.sleep(2)
    os_button = navegador.find_element_by_id("W5|_378_|O.S.")
    os_button.click()

    time.sleep(5)
    iframe_num = listing_iframes("vDATAINICIO")
    
    if iframe_num != False:
        navegador.switch_to.frame(iframe_num)

        # navegador.find_element_by_id("vDATAINICIO_dp_container").send_keys("01/06/2021")
        navegador.execute_script('document.getElementById("vDATAINICIO_dp_container").setAttribute("value", "01/06/2021")')
        # navegador.find_element_by_id("vDATAFIM_dp_container").send_keys("14/06/2021")
        navegador.execute_script(f'document.getElementById("vDATAFIM_dp_container").setAttribute("value", "{enddate}")')
        navegador.find_element_by_xpath('//*[@id="TABLE6"]/tbody/tr/td/table/tbody/tr[7]/td/input').click()
        time.sleep(1)
        navegador.find_element_by_id("BTNGERAR").click()
        navegador.switch_to.default_content()
        time.sleep(10)

        iframe_num2 = listing_iframes("IMGMALADIRETA")
        if iframe_num2 != False:
            navegador.switch_to.frame(iframe_num2)
            navegador.find_element_by_id("IMGMALADIRETA").click()
            print("dealernet report downloaded!")


def renaming_file():
    downloaded_file = max([download_dir + "\\" + f for f in os.listdir(download_dir)],key=os.path.getctime)
    os.rename(downloaded_file, download_dir + "\dealernet.xlsx")
    print("File is avaliable at D:/JAC/ATUALIZACAO_STATUS_GARANTIA \n Please look up!")


#main logic
navegador.get("http://dealernet.shcnet.com.br/LoginAux.aspx?Windows")
time.sleep(5)
login_in_dealernet()
time.sleep(12)
end_date = datetime.today().strftime('%d/%m%Y')
getting_os_report(end_date)
time.sleep(5)
renaming_file()
navegador.quit()