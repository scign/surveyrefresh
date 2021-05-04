import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import JavascriptException, TimeoutException, StaleElementReferenceException

TIMEOUT = 120

def set_options(download_path):
  options = Options()
  options.headless = True
  options.set_preference("browser.download.dir", download_path)
  options.set_preference("browser.download.folderList", 2)
  options.set_preference("browser.download.useDownloadDir", True)
  options.set_preference("browser.download.manager.showWhenStarting", False)
  options.set_preference("browser.download.viewableInternally.enabledTypes", "")
  options.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/csv, text/tab-separated-values, application/vnd.ms-excel, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
  return options

def login(params):
  driver = webdriver.Firefox(options=params['options'], service_log_path=os.devnull)
  driver.get(params['url'])
  WebDriverWait(driver, TIMEOUT).until(EC.visibility_of_element_located((By.ID, params['username'][0])))
  driver.find_element_by_id(params['username'][0]).send_keys(params['username'][1])
  WebDriverWait(driver, TIMEOUT).until(EC.visibility_of_element_located((By.ID, params['password'][0])))
  driver.find_element_by_id(params['password'][0]).send_keys(params['password'][1])
  driver.find_element_by_id(params['submit']).click()
  return driver

def go_to(driver, url):
  driver.get(url)

def wait_for_id(driver, id):
  try:
    WebDriverWait(driver, TIMEOUT).until(EC.visibility_of_element_located((By.ID, id)))
    return True
  except TimeoutException:
    return False

def get_id(driver, id):
  try:
    return driver.find_element_by_id(id)
  except TimeoutException:
    return None

def wait_for_class(driver, class_):
  try:
    WebDriverWait(driver, TIMEOUT).until(EC.visibility_of_element_located((By.CLASS_NAME, class_)))
    return True
  except TimeoutException:
    return False

def get_class(driver, class_):
  try:
    return driver.find_elements_by_class_name(class_)
  except TimeoutException:
    return None
