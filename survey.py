import os
import logging
import pandas as pd
from time import sleep
from tempfile import TemporaryDirectory

if os.name=='posix':
  from pyvirtualdisplay import Display

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options
from selenium.common.exceptions import JavascriptException, TimeoutException, StaleElementReferenceException

from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.files.file_creation_information import FileCreationInformation

"""
  Turn on logging to get some progress output
"""
logging.basicConfig(
  format='%(asctime)s %(levelname)-8s %(message)s',
  datefmt='%Y-%m-%d %H:%M:%S',
  level=logging.INFO
)

"""
  Set up survey datasets to download
  Format:
    - Node ID (number from the download URL)
    - Destination filename (what the file should be called when uploaded to SharePoint)
    - Survey title (What the survey should be named in Power BI)
    - Completion threshold (responses with the same answers and within this number of seconds from each other will be removed)
"""
nodes = {
  '46099': ('202105_pp_p_teacher.xlsx', 'Pre-Primary and Primary teachers 2021 year end', 60),
  '45400': ('202105_g4_g6_student.xlsx', 'Grade 4-6 student 2021 year end', 60),
  '45404': ('202105_g4_g6_parent.xlsx', 'Grade 4-6 parent 2021 year end', 60),
  '45402': ('202105_pp_g3_family.xlsx', 'Pre-Primary - Grade 3 family 2021 year end', 60),
  '45401': ('202105_g7_g12_student.xlsx', 'Secondary student 2021 year end', 60),
  '45403': ('202105_g7_g12_parent.xlsx', 'Secondary parent 2021 year end', 60),
  #'42518': ('old_survey_test_file.xlsx', 'Old survey - test 2021 year end', 60),
}

survey_username = os.environ.get('SURVEY_USERNAME','')
survey_password = os.environ.get('SURVEY_PASSWORD','')

url = 'https://balmoralaid.sharepoint.com/sites/Planner_C6xE'
sharepoint_username = os.environ.get('ITREB_USERNAME','')
sharepoint_password = os.environ.get('ITREB_PASSWORD','')
relative_url = '/sites/Planner_C6xE/Power%20BI%20Dashboard/Survey%202021-05/'

if os.name=='posix':
  display = Display(visible=0, size=(800, 600))
  display.start()

with TemporaryDirectory() as download_path:
  options = Options()
  options.headless = True
  options.set_preference("browser.download.dir", download_path)
  options.set_preference("browser.download.folderList", 2)
  options.set_preference("browser.download.useDownloadDir", True)
  options.set_preference("browser.download.manager.showWhenStarting", False)
  options.set_preference("browser.download.viewableInternally.enabledTypes", "")
  options.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/csv, text/tab-separated-values, application/vnd.ms-excel, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

  for node in nodes.keys():
    survey_file = nodes[node][0]
    survey_name = nodes[node][1]
    completion_threshold = nodes[node][2]

    try:
      logging.info(f'Starting browser ({node} - {survey_name})')
      driver = webdriver.Firefox(options=options, service_log_path=os.devnull)
      logging.info('Logging in')
      driver.get('https://iicanada.org/user')
      WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, 'edit-name')))
      driver.find_element_by_id('edit-name').send_keys(survey_username)
      WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, 'edit-pass')))
      driver.find_element_by_id('edit-pass').send_keys(survey_password)
      driver.find_element_by_id('edit-submit').click()

      logging.info('Navigating to survey responses')
      driver.get('https://iicanada.org/node/' + node + '/webform-results/download')  #test
      WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, 'edit-format-delimited')))
      WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.CLASS_NAME, 'fieldset-title')))

      # find how many results
      driver.find_elements_by_class_name('fieldset-title')[2].click()
      num_responses = int(driver.find_element_by_id('edit-range-range-type').text.split('(')[1].split(' ')[0])

      logging.info(f'Responses to date: {num_responses}')

      new_file = ''
      if num_responses:
        _,_,current_files = next(os.walk(download_path))
        logging.info('Downloading responses')
        driver.find_element_by_id('edit-submit').click()
        WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, 'edit-format-excel')))
        logging.info('Waiting for download to complete')

        for i in range(20):
          _,_,latest_files = next(os.walk(download_path))
          if len(latest_files) > len(current_files):
            break
          sleep(1)
        if len(latest_files) == len(current_files):
          logging.error('No new files detected within time limit - exiting (no data)')

        new_file = os.path.join(download_path, (set(latest_files) - set(current_files)).pop())
        logging.info(f'File downloaded: {new_file}')

      else:
        logging.info('No responses found - exiting (no data)')

      driver.quit()

      if num_responses and new_file:
        logging.info('Opening file')
        df = pd.read_excel(new_file, skiprows=2, sheet_name="Sheet1", engine='openpyxl')
        df['survey'] = survey_name
        logging.info(f'Found [rows, columns]: {df.shape}')

        logging.info('Removing duplicates')
        df['all_cols'] = df[df.columns[4:]].apply(
          lambda x: ''.join(x.dropna().astype(str)), axis=1
        )
        dfx = []
        for n,g in df.groupby('all_cols'):
          dfg = g.sort_values('Submitted Time')
          dfg['s_diff'] = dfg['Submitted Time'].diff().dt.total_seconds()
          dfx.append(dfg[~(dfg.s_diff < completion_threshold)].drop(columns=['all_cols','s_diff']))

        dups_removed = pd.concat(dfx)
        dups = df.shape[0] - dups_removed.shape[0]
        dups_pct = dups / df.shape[0]
        logging.info(f'Final [rows, columns]: {dups_removed.shape} - {dups}({dups_pct:0.1%}) duplicates removed')

        common_cols = ['Serial','SID','Submitted Time','Completed Time','Draft','IP Address','UID','Username','survey']
        melted = dups_removed.melt(common_cols, var_name='Question', value_name='Response')
        output_file = os.path.join(download_path, survey_file)
        melted.to_excel(output_file)

        logging.info('Authenticating to SharePoint')
        ctx_auth = AuthenticationContext(url)
        if ctx_auth.acquire_token_for_user(sharepoint_username, sharepoint_password):
          ctx = ClientContext(url, ctx_auth)
          web = ctx.web
          ctx.load(web)
          ctx.execute_query()
          logging.info(f'Connected to SharePoint site {web.properties["Title"]}')
        else:
          logging.error(f'Error connecting to SharePoint: {ctx_auth.get_last_error()}')

        logging.info(f'Uploading {new_file}')
        fci = FileCreationInformation()
        fci.overwrite = True
        fci.url = survey_file
        with open(output_file, 'rb') as f:
          fci.content = f.read()
        upload = ctx.web.get_folder_by_server_relative_url(relative_url).files.add(fci)
        ctx.execute_query()

        logging.info(f'------ Completed ({node} - {survey_name}) ------')

    except Exception as e:
      logging.exception(f"Failed to process {survey_name}")

    finally:
      try:
        driver.quit()
      except:
        pass

      if os.name=='posix':
        try:
          display.stop()
        except:
          pass

