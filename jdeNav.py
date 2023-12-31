# jdeNav.py
# This file holds the functions that navigate JD Edwards to pull necessary report with minimal input from the user
# Automate navigating Oracle JD Edwards
# Code to be imported into engine to work with GUI


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from datetime import date as dt
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC 


website = r'COMPANY_URL' # The URL for the login page for company's JD Edwards site, typically saved in seperate text file for security purposes


# Tools
# Wait for elements to exist and interact
def justWait(driverName, method, locator):
    WebDriverWait(driverName, 5).until(  # Chose 5 seconds as default
        EC.presence_of_all_elements_located((method, locator))  # method will be either By.XPATH, or By.ID
    )
def waitAndClick(driverName, method, locator):
    justWait(driverName, method, locator)
    if method == By.ID:
        drop = driverName.find_element(By.ID, locator)
        time.sleep(1)
        selectDrop = Select(drop)
        selectDrop.select_by_value('Literal')
    elif method == By.XPATH:
        driverName.find_element(By.XPATH, locator).click()
    else:
        raise ValueError("Wrong locator type, only use By.ID or By.XPATH: Method '{}' is not supported.".format(method))


# Automate navigating JD Edwards for Accounting Tenant Improvement Report for some project
# User only needs to provide project job number for requested report
def getATR(projNum):
    driver = webdriver.Chrome()
    driver.get(website)

    # Network login for JDE, must be on vpn or on company wifi to allow network login (no need to code user's login credentials)
    waitAndClick(driver, By.XPATH, "/html/body/div/table/tbody/tr[2]/td/form/table/tbody/tr/td/div/table/tbody/tr/td/p/a")

    # Access 'Enter Contract' through dropdown menu
    waitAndClick(driver, By.XPATH, '//*[@id="drop_mainmenu"')
    waitAndClick(driver, By.XPATH, "//*[@tasklabel='COMPANY Properties Menu']")
    waitAndClick(driver, By.XPATH, '/html/body/div[10]/table/tbody/tr/td/div/div[4]')
    waitAndClick(driver, By.XPATH, '/html/body/div[11]/table/tbody/tr/td/div/div[3]')
    waitAndClick(driver, By.XPATH, '/html/body/div[12]/table/tbody/tr/td/div/div[1]')
    time.sleep(1)

    # Switch to iframe
    iframe = driver.switch_to.frame('e1menuAppIframe')
    waitAndClick(driver, By.XPATH, '/html/body/form[3]/div/table/tbody/tr/td/div/span[2]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[2]/td/div/table/tbody/tr/td[2]/table/tbody/tr[1]/td/table/tbody/tr/td[1]/div/input')
    waitAndClick(driver, By.XPATH, '/html/body/form[3]/table[2]/tbody/tr/td/table/tbody/tr/td[1]')

    # Access Data Selection page
    waitAndClick(driver, By.XPATH, '/html/body/form[3]/div/table/tbody/tr/td/div/span[1]/table[2]/tbody/tr/td/div/span[1]/div/input')
    waitAndClick(driver, By.XPATH, '/html/body/form[3]/table[2]/tbody/tr/td/table/tbody/tr/td[3]/table/tbody/tr/td[3]/a/div')
    waitAndClick(driver, By.ID, 'RightOperand1')
    litVal = driver.find_element('xpath','/html/body/form[2]/div/table/tbody/tr/td/table[2]/tbody/tr/td/table/tbody/tr[2]/td[2]/input')
    litVal.send_keys(projNum)
    waitAndClick(driver, By.XPATH, '/html/body/table[2]/tbody/tr/td/table/tbody/tr/td[1]/a/img')
    waitAndClick(driver, By.XPATH, '/html/body/table[2]/tbody/tr/td/table/tbody/tr/td[1]/a/img')
    time.sleep(1)
    today = str(dt.today())
    y = today[2:4]  # Current year
    m = today[5:7]  # Current month
    d = today[-2:]  # Current day (of the month)
    perNum = driver.find_element('xpath','/html/body/div/form/table/tbody/tr/td/table/tbody/tr/td/table[2]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[4]/td[2]/input')
    perNum.clear()
    perNum.send_keys(m)  # Current period number (month)
    fisYea = driver.find_element('xpath','/html/body/div/form/table/tbody/tr/td/table/tbody/tr/td/table[2]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[6]/td[2]/input')
    fisYea.clear()
    fisYea.send_keys(y)  # Populate current fiscal year
    waitAndClick(driver, By.XPATH, '/html/body/table[2]/tbody/tr/td/table/tbody/tr/td[1]/a/img')
    waitAndClick(driver, By.XPATH, '/html/body/form[3]/table[2]/tbody/tr/td/table/tbody/tr/td[1]/a/img')
    time.sleep(200)  # 200 Seconds to navigate to results page for necessary document(s)



# Automate navigating JDE to obtain a specific contract report
# User only needs to provide the assigned contract number
def printContract(contractNum):
    driver = webdriver.Chrome()
    driver.get(website)

    # Network login for JDE, must be on vpn or COMPANY wifi
    login = driver.find_element('xpath','//*[@id="boxcontent"]/p/a')
    login.click()

    # Access 'Enter Contract' page through dropdown menu
    waitAndClick(driver, By.XPATH, '//*[@id="drop_mainmenu"]')
    waitAndClick(driver, By.XPATH, "//*[@tasklabel='COMPANY Properties Menu']")
    waitAndClick(driver, By.XPATH, '/html/body/div[10]/table/tbody/tr/td/div/div[4]')
    waitAndClick(driver, By.XPATH, '/html/body/div[11]/table/tbody/tr/td/div/div[1]')
    waitAndClick(driver, By.XPATH, '/html/body/div[12]/table/tbody/tr/td/div/div[2]/div')
    time.sleep(1)

    # Switch to iframe
    iframe = driver.switch_to.frame('e1menuAppIframe')

    # Populate remaining data input page for report
    dataSelect = driver.find_element('xpath','/html/body/form[3]/div/table/tbody/tr/td/div/span[1]/table[2]/tbody/tr/td/div/span[1]/div/input')
    dataSelect.click()
    waitAndClick(driver, By.XPATH, '/html/body/form[3]/table[2]/tbody/tr/td/table/tbody/tr/td[3]/table/tbody/tr/td[3]/a/div')
    waitAndClick(driver, By.ID, 'RightOperand2')
    litVal = driver.find_element('xpath','/html/body/form[2]/div/table/tbody/tr/td/table[2]/tbody/tr/td/table/tbody/tr[2]/td[2]/input')
    litVal.clear()
    litVal.send_keys(contractNum)
    waitAndClick(driver, By.XPATH, '/html/body/table[2]/tbody/tr/td/table/tbody/tr/td[1]/a/img')
    waitAndClick(driver, By.XPATH, '/html/body/table[2]/tbody/tr/td/table/tbody/tr/td[1]/a/img')
    waitAndClick(driver, By.XPATH, '/html/body/form[3]/table[2]/tbody/tr/td/table/tbody/tr/td[1]/a/img')
    time.sleep(200)  # 200 Seconds to navigate to results page for necessary document(s)


# Automate navigating JD Edwards for a contract's change order report
# User only needs to provide contract number and the change order number
def printCO(contractNum, coNum):
    driver = webdriver.Chrome()
    driver.get(website)

    # Network login for JDE, must be on vpn or COMPANY wifi
    login = driver.find_element('xpath','//*[@id="boxcontent"]/p/a')
    login.click()

    # Access 'Enter Contract' through dropdown menu
    waitAndClick(driver, By.XPATH, '//*[@id="drop_mainmenu"')
    waitAndClick(driver, By.XPATH, "//*[@tasklabel='COMPANY Properties Menu']")
    waitAndClick(driver, By.XPATH, '/html/body/div[10]/table/tbody/tr/td/div/div[4]')
    waitAndClick(driver, By.XPATH, '/html/body/div[11]/table/tbody/tr/td/div/div[1]')
    waitAndClick(driver, By.XPATH, '/html/body/div[12]/table/tbody/tr/td/div/div[3]/div')
    time.sleep(1)

    # Switch to iframe for data input
    iframe = driver.switch_to.frame('e1menuAppIframe')

    # Populate data selection page
    waitAndClick(driver, By.XPATH, '/html/body/form[3]/div/table/tbody/tr/td/div/span[1]/table[2]/tbody/tr/td/div/span[1]/div/input')
    waitAndClick(driver, By.XPATH, '/html/body/form[3]/table[2]/tbody/tr/td/table/tbody/tr/td[3]/table/tbody/tr/td[3]/a/div')
    waitAndClick(driver, By.ID, 'RightOperand2')
    litVal = driver.find_element('xpath','/html/body/form[2]/div/table/tbody/tr/td/table[2]/tbody/tr/td/table/tbody/tr[2]/td[2]/input')
    litVal.clear()
    litVal.send_keys(contractNum)
    waitAndClick(driver, By.XPATH, '/html/body/table[2]/tbody/tr/td/table/tbody/tr/td[1]/a/img')
    justWait(driver, By.ID, 'Comparison3')
    dropDown = driver.find_element('id', 'Comparison3')
    time.sleep(1)
    select = Select(dropDown)
    select.select_by_value('0')
    waitAndClick(driver, By.ID, 'RightOperand3')
    litVal = driver.find_element('xpath','/html/body/form[2]/div/table/tbody/tr/td/table[2]/tbody/tr/td/table/tbody/tr[2]/td[2]/input')
    litVal.clear()
    litVal.send_keys(coNum)
    waitAndClick(driver, By.XPATH, '/html/body/table[2]/tbody/tr/td/table/tbody/tr/td[1]/a/img')
    waitAndClick(driver, By.XPATH, '/html/body/table[2]/tbody/tr/td/table/tbody/tr/td[1]/a/img')
    waitAndClick(driver, By.XPATH, '/html/body/form[3]/table[2]/tbody/tr/td/table/tbody/tr/td[1]/a/img')
    time.sleep(200)  # 200 Seconds to navigate to results page for necessary document(s)


# Automate navigating JD Edwards to obtain the development summary report for a capital project
# User only needs to provide the project job number for the reqeusted report
def developmentSummaryReport(projNum):
    driver = webdriver.Chrome()
    driver.get(website)

    # Network login for JDE, must be on vpn or company wifi
    login = driver.find_element('xpath','//*[@id="boxcontent"]/p/a')
    login.click()
    time.sleep(1)

    # Access 'Enter Contract' through dropdown menu
    waitAndClick(driver, By.XPATH, '//*[@id="drop_mainmenu"]')
    gpMenu = driver.find_element('xpath',"//*[@tasklabel='COMPANY Properties Menu']")
    gpMenu.click()
    waitAndClick(driver, By.XPATH, '/html/body/div[10]/table/tbody/tr/td/div/div[4]')
    waitAndClick(driver, By.XPATH, '/html/body/div[11]/table/tbody/tr/td/div/div[3]')
    jobCostReporting = driver.find_element('xpath','/html/body/div[11]/table/tbody/tr/td/div/div[3]')
    jobCostReporting.click()
    waitAndClick(driver, By.XPATH, '/html/body/div[12]/table/tbody/tr/td/div/div[5]')
    time.sleep(1)

    # switch to iframe for data input
    iframe = driver.switch_to.frame('e1menuAppIframe')
    waitAndClick(driver, By.XPATH, '/html/body/form[3]/div/table/tbody/tr/td/div/span[2]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[2]/td/div/table/tbody/tr/td[2]/table/tbody/tr[1]/td/table/tbody/tr/td[1]/div/input')
    waitAndClick(driver, By.XPATH, '/html/body/form[3]/table[2]/tbody/tr/td/table/tbody/tr/td[1]')

    # data selection
    waitAndClick(driver, By.XPATH, '/html/body/form[3]/div/table/tbody/tr/td/div/span[1]/table[2]/tbody/tr/td/div/span[1]/div/input')
    waitAndClick(driver, By.XPATH, '/html/body/form[3]/table[2]/tbody/tr/td/table/tbody/tr/td[3]/table/tbody/tr/td[3]/a/div')
    waitAndClick(driver, By.ID, 'RightOperand1')
    litVal = driver.find_element('xpath','/html/body/form[2]/div/table/tbody/tr/td/table[2]/tbody/tr/td/table/tbody/tr[2]/td[2]/input')
    litVal.send_keys(projNum)
    waitAndClick(driver, By.XPATH, '/html/body/table[2]/tbody/tr/td/table/tbody/tr/td[1]/a/img')
    waitAndClick(driver, By.XPATH, '/html/body/table[2]/tbody/tr/td/table/tbody/tr/td[1]/a/img')
    time.sleep(1)
    today = str(dt.today())
    y = today[2:4]
    m = today[5:7]
    d = today[-2:]
    perNum = driver.find_element('xpath','/html/body/div/form/table/tbody/tr/td/table/tbody/tr/td/table[2]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[4]/td[2]/input')
    perNum.clear()
    perNum.send_keys(m)
    fisYea = driver.find_element('xpath','/html/body/div/form/table/tbody/tr/td/table/tbody/tr/td/table[2]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[6]/td[2]/input')
    fisYea.clear()
    fisYea.send_keys(y)
    waitAndClick(driver, By.XPATH, '/html/body/table[2]/tbody/tr/td/table/tbody/tr/td[1]/a/img')
    waitAndClick(driver, By.XPATH, '/html/body/form[3]/table[2]/tbody/tr/td/table/tbody/tr/td[1]/a/img')
    time.sleep(200)  # 200 Seconds to navigate to results page for necessary document(s)


# Automate navigating JD Edwards to obtain the AR Aging report (no further input needed)
def arAge():
    driver = webdriver.Chrome()
    driver.get(website)

    # network login for JDE, must be on vpn or COMPANY wifi
    login = driver.find_element('xpath','//*[@id="boxcontent"]/p/a')
    login.click()
    # access A/R Aging with Details through dropdown menu
    waitAndClick(driver, By.XPATH, '//*[@id="drop_mainmenu"]')
    navigator = driver.find_element('xpath','//*[@id="drop_mainmenu"]')
    navigator.click()
    gpMenu = driver.find_element('xpath',"//*[@tasklabel='COMPANY Properties Menu']")
    gpMenu.click()
    waitAndClick(driver, By.XPATH, "//*[@tasklabel='Property Management")
    waitAndClick(driver, By.XPATH, "//*[@tasklabel='Tenant &#38; Lease Information']")
    waitAndClick(driver, By.XPATH, "//*[@tasklabel='A/R Aging with Details']")
    time.sleep(1)

    # Switch to iframe
    iframe = driver.switch_to.frame('e1menuAppIframe')

    # Popoulate data in the data selection page
    waitAndClick(driver, By.XPATH, '/html/body/form[3]/div/table/tbody/tr/td/div/span[2]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[2]/td/div/table/tbody/tr/td[2]/table/tbody/tr[1]/td/table/tbody/tr/td[1]/div/input')
    waitAndClick(driver, By.XPATH, '/html/body/form[3]/table[2]/tbody/tr/td/table/tbody/tr/td[1]')
    waitAndClick(driver, By.XPATH, '/html/body/form[3]/div/table/tbody/tr/td/div/span[1]/table[2]/tbody/tr/td/div/span[1]/div/input')
    waitAndClick(driver, By.XPATH, '/html/body/form[3]/table[2]/tbody/tr/td/table/tbody/tr/td[3]/table/tbody/tr/td[3]/a/div')
    time.sleep(1)
    compDropDown = driver.find_element('id','Comparison1')
    select = Select(compDropDown)
    select.select_by_value('5')
    waitAndClick(driver, By.ID, 'RightOperand1')
    litVal = driver.find_element('xpath','/html/body/form[2]/div/table/tbody/tr/td/table[2]/tbody/tr/td/table/tbody/tr[2]/td[2]/input')
    litVal.send_keys('16')
    waitAndClick(driver, By.XPATH, '/html/body/table[2]/tbody/tr/td/table/tbody/tr/td[1]/a/img')
    justWait(driver, By.ID, 'Comparison2')
    glOffsetDropDown = driver.find_element('id', 'Comparison2')
    select = Select(glOffsetDropDown)
    select.select_by_value('0')
    glRighOp = driver.find_element('id', 'RightOperand2')
    time.sleep(1)
    select = Select(glRighOp)
    select.select_by_value('Literal')
    litVal = driver.find_element('xpath','/html/body/form[2]/div/table/tbody/tr/td/table[2]/tbody/tr/td/table/tbody/tr[2]/td[2]/input')
    litVal.send_keys('TFO')
    justWait(driver, By.ID, 'hc_Select')
    ok = driver.find_element('id','hc_Select')
    ok.click()
    ok = driver.find_element('id','hc_Select')
    ok.click()
    waitAndClick(driver, By.NAME, 'E2')
    ageDate = driver.find_element('name', 'E2')
    ageDate.clear()
    today = str(dt.today())
    y = today[0:4]
    m = today[5:7]
    d = today[-2:]
    enteredDate = (f'{m}/{d}/{y}')
    ageDate.send_keys(enteredDate)
    waitAndClick(driver, By.XPATH, '//*[@id="modelessTabHeaders"]/tbody/tr/td/table/tbody/tr/td[3]/table/tbody/tr/td/table/tbody/tr/td/span/a')
    justWait(driver, By.ID, 'PO17T2')
    dateEnt = driver.find_element('id', 'PO17T2')
    dateEnt.clear()
    dateEnt.send_keys(enteredDate)
    waitAndClick(driver, By.XPATH, '/html/body/table[2]/tbody/tr/td/table/tbody/tr/td[1]/a/img')
    waitAndClick(driver, By.XPATH, '/html/body/form[3]/table[2]/tbody/tr/td/table/tbody/tr/td[1]/a/img')
    time.sleep(200)  # 200 Seconds to navigate to results page for necessary document(s)


