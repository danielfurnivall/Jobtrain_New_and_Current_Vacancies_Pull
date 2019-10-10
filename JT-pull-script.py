from selenium import webdriver
import time
import datetime
import configparser
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from calendar import monthrange

today = datetime.datetime.now()
current = today.strftime("01/%m/%Y" + " 00:00:00")
today = today.replace(day=1)
today = today.replace(month = today.month - 1)
today = today.replace(day = monthrange(today.year, today.month)[1])
startdate = (today.strftime("01/%m/%Y" + " 00:00:00"))
enddate = (today.strftime("%d/%m/%Y" + " 00:00:00"))
print(current + "\n" + startdate + "\n" + enddate)

config = configparser.ConfigParser()
config.read(r'W:\\Python\Danny\SSTS Extract\SSTSConf.ini')
print("W:/Workforce Monthly Reports/Monthly_Reports/" +today.strftime('%b-%y ')+ "Snapshot/Vacancies/")
chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory": 'W:\Workforce Monthly Reports\Monthly_Reports\\' + today.strftime('%b-%y ')+ 'Snapshot\Vacancies\\',
         'safebrowsing.disable_download_protection': True}
chromeOptions.add_experimental_option("prefs", prefs)
browser = webdriver.Chrome(executable_path="W:/Danny/Chrome Webdriver/chromedriver.exe",
                           options=chromeOptions)

browser.get('https://reporting52.jobtrain.co.uk')
uname = browser.find_element_by_xpath('/html/body/div[1]/div/form/table/tbody/tr/td[3]/table/tbody/tr[1]/td[2]/div/input')
pword = browser.find_element_by_xpath('/html/body/div[1]/div/form/table/tbody/tr/td[3]/table/tbody/tr[1]/td[3]/div/input')
uname.clear()
pword.clear()
uname.send_keys(config.get('JobTrain', 'uname'))
pword.send_keys(config.get('JobTrain','pword'))
browser.find_element_by_id('logonButton').click()
WebDriverWait(browser, 90).until(
    ec.element_to_be_clickable((By.XPATH, '//*[@id="headerContainer"]/div[1]/div/div/input')))
searchbar = browser.find_element_by_xpath('//*[@id="headerContainer"]/div[1]/div/div/input')
searchbar.clear()
time.sleep(1)
searchbar.send_keys('WI "Current" Vacancies')
browser.find_element_by_xpath('//*[@id="headerContainer"]/div[1]/div/div/div[2]/div[2]/img').click()
x = '//*[@id="browsePageBody"]/form/div[10]/div/div[3]/div/div/div/div[1]/div/div[1]/div[2]/div[2]'
WebDriverWait(browser, 90).until(
    ec.element_to_be_clickable((By.XPATH, x)))
time.sleep(1)
actionChains = ActionChains(browser)
actionChains.double_click(browser.find_element_by_xpath(x))
actionChains.perform()
WebDriverWait(browser, 90).until(
    ec.element_to_be_clickable((By.XPATH, '//*[@id="pagecontent"]/div[2]/div[3]/div[3]/div[3]/div/div/div[2]/div/table')))
dateinput = browser.find_element_by_xpath('//*[@id="132506"]/div/div[4]/div/div/table/tr/td[1]/input')
dateinput.send_keys(current)
browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[3]/div[1]/'+
                              'div/div/div[3]/div[1]/table/tbody/tr/td/div/table/tbody/tr/td[2]').click()
browser.find_element_by_id('reportexport').click()

WebDriverWait(browser, 90).until(
    ec.element_to_be_clickable((By.ID, 'rptDataOverlayPanelContent')))
browser.find_element_by_xpath('//*[@id="rptDataOverlayPanelContent"]/div/div[1]/table/tbody/tr[1]/td[2]/a').click()
time.sleep(2)
browser.find_element_by_xpath('//*[@id="csvExportBtnContainer"]/button/span').click()
x = '//*[@id="rptOutputTopToolbar"]/div/div[1]/div[2]/table[4]/tbody/tr/td/img'
WebDriverWait(browser, 90).until(
    ec.element_to_be_clickable((By.XPATH, x)))
browser.find_element_by_xpath(x).click() #quit button
WebDriverWait(browser, 90).until(
    ec.element_to_be_clickable((By.XPATH, '//*[@id="browsePageBody"]/form/div[10]/div/div[1]/div[2]/div/input')))
searchbar = browser.find_element_by_xpath('//*[@id="browsePageBody"]/form/div[10]/div/div[1]/div[2]/div/input')
searchbar.clear()
searchbar.send_keys('WI New Vacancies')
browser.find_element_by_xpath('//*[@id="browsePageBody"]/form/div[10]/div/div[1]/div[2]/div/div[2]/div[2]/img').click()
x = '//*[@id="browsePageBody"]/form/div[10]/div/div[3]/div/div/div/div[1]/div/div[1]'
WebDriverWait(browser, 90).until(
    ec.element_to_be_clickable((By.XPATH, x)))
actionChains = ActionChains(browser)
actionChains.double_click(browser.find_element_by_xpath(x))
actionChains.perform()
WebDriverWait(browser, 90).until(
    ec.element_to_be_clickable((By.XPATH, '//*[@id="116454"]/div/div[4]/div/div/div/div[1]/input')))
browser.find_element_by_xpath('//*[@id="116454"]/div/div[4]/div/div/div/div[1]/input').click()

dateinput1 = browser.find_element_by_xpath('/html/body/div[5]/div[2]/div[2]/div[1]/div[1]/input')
dateinput2 = browser.find_element_by_xpath('/html/body/div[5]/div[2]/div[2]/div[1]/div[2]/input')
WebDriverWait(browser, 90).until(
    ec.element_to_be_clickable((By.XPATH, '/html/body/div[5]/div[2]/div[2]/div[1]/div[2]/input')))
dateinput1.clear()
dateinput2.clear()
dateinput1.send_keys(startdate)
dateinput2.send_keys(enddate)


WebDriverWait(browser, 90).until(
    ec.element_to_be_clickable((By.ID, 'reportexport')))
browser.find_element_by_id('reportexport').click()

WebDriverWait(browser, 90).until(
    ec.element_to_be_clickable((By.ID, 'rptDataOverlayPanelContent')))
browser.find_element_by_xpath('//*[@id="rptDataOverlayPanelContent"]/div/div[1]/table/tbody/tr[1]/td[2]/a').click()
browser.find_element_by_xpath('//*[@id="csvExportBtnContainer"]/button/span').click()