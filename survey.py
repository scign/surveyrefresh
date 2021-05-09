import os
import asyncio
import logging
import pandas as pd
from time import sleep
from tempfile import TemporaryDirectory

from browser import login, wait_for_id, get_id, wait_for_class, get_class
from sharepoint import upload_to_sharepoint, get_file_details
from telegram import send_message
import display

logging.basicConfig(
  format='%(asctime)s %(levelname)-8s %(message)s',
  datefmt='%Y-%m-%d %H:%M:%S',
  level=logging.INFO
)

def get_logged_in_browser(survey_id, download_path):
  survey_name = surveys[survey_id][1]
  logging.info(f'Starting browser ({survey_id} - {survey_name})')
  b = login(
    download_path=download_path,
    url='https://iicanada.org/user',
    username=('edit-name', os.environ.get('ITREB_USERNAME','')),
    password=('edit-pass', os.environ.get('ITREB_PASSWORD','')),
    submit='edit-submit'
  )
  logging.info('Navigating to survey responses')
  b.get('https://iicanada.org/node/' + survey_id + '/webform-results/download')
  wait_for_id(b, 'edit-format-delimited')
  return b

def get_num_responses(b):
  # find how many results
  elems = get_class(b, 'fieldset-title') if wait_for_class(b, 'fieldset-title') else []
  if len(elems)==3:
    elems[2].click()  # open the third section (download settings)

  e = get_id(b, 'edit-range-range-type')
  num_responses = int(e.text.split('(')[1].split(' ')[0]) if e else 0

  logging.info(f'Responses to date: {num_responses}')
  return num_responses

def download_responses(b, download_path):
  _,_,current_files = next(os.walk(download_path))
  logging.info('Downloading responses')
  
  get_id(b, 'edit-submit').click()
  wait_for_id(b, 'edit-format-excel')
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
  return new_file

def get_responses(survey_id, download_path):
  b = get_logged_in_browser(survey_id, download_path)
  num_responses = get_num_responses(b)
  new_file = download_responses(b, download_path) if num_responses else ''
  b.quit()
  return num_responses, new_file

def remove_duplicates(df, completion_threshold):
  logging.info(f'Found [rows, columns]: {df.shape}')

  logging.info('Removing duplicates')
  # ignore first 4 columns: SID, Serial, Completed Time, Submitted Time (all guaranteed unique)
  df['all_cols'] = df[df.columns[4:]].apply(
    lambda x: ''.join(x.dropna().astype(str)), axis=1
  )
  dfx = []
  # group by the remaining columns
  for n,g in df.groupby('all_cols'):
    dfg = g.sort_values('Submitted Time')
    dfg['s_diff'] = dfg['Submitted Time'].diff().dt.total_seconds()
    # filter where responses within the group are closer than the completion_threshold and drop them
    dfx.append(dfg[~(dfg.s_diff < completion_threshold)].drop(columns=['all_cols','s_diff']))

  # recombine
  dups_removed = pd.concat(dfx)
  dups = df.shape[0] - dups_removed.shape[0]
  dups_pct = dups / df.shape[0]
  logging.info(f'Final [rows, columns]: {dups_removed.shape} - {dups} ({dups_pct:0.1%}) duplicates removed')

  common_cols = ['Serial','SID','Submitted Time','Completed Time','Draft','IP Address','UID','Username','survey']
  return dups_removed.melt(common_cols, var_name='Question', value_name='Response')


@display.virtual_display
def download_surveys():
  with TemporaryDirectory() as download_path:
    for survey_id in surveys.keys():
      survey_file = surveys[survey_id][0]
      survey_name = surveys[survey_id][1]
      completion_threshold = surveys[survey_id][2]
      num_responses, new_file = get_responses(survey_id, download_path)

      if num_responses and new_file:
        logging.info('Opening file')
        df = pd.read_excel(new_file, skiprows=2, sheet_name="Sheet1", engine='openpyxl')
        df['survey'] = survey_name
        
        output_file = os.path.join(download_path, survey_file)
        melted = remove_duplicates(df, completion_threshold)
        melted.to_excel(output_file)

        upload_to_sharepoint(output_file, survey_file)

        logging.info(f'------ Completed ({survey_id} - {survey_name}) ------')

"""
  Set up survey datasets to download
  Format:
    - Node ID (number from the download URL)
    - Destination filename (what the file should be called when uploaded to SharePoint)
    - Survey title (What the survey should be named in Power BI)
    - Completion threshold (responses with the same answers and within this number of seconds from each other will be removed)
"""
surveys = {
  #'46099': ('202105_pp_p_teacher.xlsx', 'Pre-Primary and Primary teachers 2021 year end', 60),
  #'45400': ('202105_g4_g6_student.xlsx', 'Grade 4-6 student 2021 year end', 60),
  #'45404': ('202105_g4_g6_parent.xlsx', 'Grade 4-6 parent 2021 year end', 60),
  #'45402': ('202105_pp_g3_family.xlsx', 'Pre-Primary - Grade 3 family 2021 year end', 60),
  '45401': ('202105_g7_g12_student.xlsx', 'Secondary student 2021 year end', 60),
  #'45403': ('202105_g7_g12_parent.xlsx', 'Secondary parent 2021 year end', 60),
  #'42518': ('old_survey_test_file.xlsx', 'Old survey - test 2021 year end', 60),
}

if __name__ == '__main__':
  download_surveys()
  files = get_file_details()
  message = ''
  for name, modified in files.items():
    message += '<b>' + name + '</b>\n' + modified.strftime('%a %b %d, %I:%M:%S %p') + '\n\n'
  asyncio.run(send_message(message))
