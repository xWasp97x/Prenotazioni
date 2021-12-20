from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.firefox.options import Options
from datetime import datetime, timedelta
import yaml
from time import sleep
import os


def download():
    now = datetime.now()
    weekday = now.isoweekday()
    monday = now - timedelta(days=weekday - 1)
    friday = monday + timedelta(days=4)

    link = 'https://web.spaggiari.eu/home/'
    link_reservations = 'https://web.spaggiari.eu/cvv/app/default/agenda.php?mode=prenotazione&aula_id=8207'
    link_file = f"https://web.spaggiari.eu/cvv/app/default/xml_export.php?stampa=%3Astampa%3A&report_name=&tipo=agenda&data={now.strftime('%d+%m+%Y')}&autore_id=5343519&tipo_export=EVENTI_PRENOTAZIONE&quad=%3Aquad%3A&materia_id=&classe_id=8207&gruppo_id=&ope=RPT&dal={monday.strftime('%Y-%m-%d')}&al={friday.strftime('%Y-%m-%d')}&tipologia=corrente&formato=xml"

    with open('/config/config.yaml') as file:
        config = yaml.safe_load(file)

    profile = webdriver.FirefoxProfile()
    profile.set_preference('browser.download.folderList', 2)
    profile.set_preference('browser.download.manager.showWhenStarting', False)
    profile.set_preference('browser.download.dir', '/home/wasp97/Downloads')
    profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'application/vnd.oasis.opendocument.spreadsheet')
    options = Options()
    options.headless = True

    driver = webdriver.Firefox(profile, executable_path=GeckoDriverManager().install(), options=options)

    driver.get(link)

    id_box = driver.find_element(by=By.ID, value='login')
    id_box.send_keys(config['username'])

    pass_box = driver.find_element(by=By.ID, value='password')
    pass_box.send_keys(config['password'])

    '''
    login_button = driver.find_element(by=By.CLASS_NAME, value='accedi btn btn-primary')
    login_button.click()
    '''

    driver.find_element(By.ID, 'fform').submit()

    load_check = EC.presence_of_element_located((By.ID, 'data_table'))
    timeout = 10
    WebDriverWait(driver, timeout).until(load_check)

    driver.set_page_load_timeout(10)
    try:
        driver.get(link_file)
    except:
        pass
    driver.quit()
