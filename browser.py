import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

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

def set_preferences():
  profile = webdriver.FirefoxProfile()
  profile.set_preference("browser.startup.homepage", "about:blank")
  profile.set_preference("browser.chrome.toolbar_style", 1) # Text on Toolbar instead of icons
  profile.set_preference("browser.display.show_image_placeholders", False) # Don't show thumbnails on not loaded images.
  profile.set_preference("browser.display.use_document_colors", False) # Don't show document colors.
  profile.set_preference("browser.display.use_document_fonts", 0) # Don't load document fonts.
  profile.set_preference("browser.display.use_system_colors", True) # Use system colors.
  profile.set_preference("browser.formfill.enable", False) # Autofill on forms disabled.
  profile.set_preference("browser.shell.checkDefaultBrowser", False)
  profile.set_preference("browser.startup.homepage", "about:blank")
  profile.set_preference("browser.startup.page", 0) # blank
  profile.set_preference("browser.tabs.forceHide", True) # Disable tabs, We won't need that.
  profile.set_preference("browser.urlbar.autoFill", False) # Disable autofill on URL bar.
  profile.set_preference("browser.urlbar.autocomplete.enabled", False) # Disable autocomplete on URL bar.
  profile.set_preference("browser.urlbar.showPopup", False) # Disable list of URLs when typing on URL bar.
  profile.set_preference("browser.urlbar.showSearch", False) # Disable search bar.
  profile.set_preference("extensions.checkCompatibility", False) # Addon update disabled
  profile.set_preference("extensions.checkUpdateSecurity", False)
  profile.set_preference("extensions.update.autoUpdateEnabled", False)
  profile.set_preference("extensions.update.enabled", False)
  profile.set_preference("general.startup.browser", False)
  profile.set_preference("plugin.default_plugin_disabled", False)
  profile.set_preference("permissions.default.image", 2) # Image load disabled again
  profile.update_preferences()
  return profile

def login(
    url='http://localhost',
    download_path='.',
    username=[],
    password=[],
    submit='',
  ):
  driver = webdriver.Firefox(
    options=set_options(download_path),
    firefox_profile=set_preferences(),
    service_log_path=os.devnull
  )
  wait = WebDriverWait(driver, TIMEOUT)
  driver.get(url)
  wait.until(EC.visibility_of_element_located((By.ID, username[0])))
  driver.find_element_by_id(username[0]).send_keys(username[1])
  wait.until(EC.visibility_of_element_located((By.ID, password[0])))
  driver.find_element_by_id(password[0]).send_keys(password[1])
  driver.find_element_by_id(submit).click()
  return driver

def wait_for_id(driver, id):
  wait = WebDriverWait(driver, TIMEOUT)
  try:
    wait.until(EC.visibility_of_element_located((By.ID, id)))
    return True
  except TimeoutException:
    return False

def get_id(driver, id):
  try:
    return driver.find_element_by_id(id)
  except TimeoutException:
    return None

def wait_for_class(driver, class_):
  wait = WebDriverWait(driver, TIMEOUT)
  try:
    wait.until(EC.visibility_of_element_located((By.CLASS_NAME, class_)))
    return True
  except TimeoutException:
    return False

def get_class(driver, class_):
  try:
    return driver.find_elements_by_class_name(class_)
  except TimeoutException:
    return None
