import os
import logging
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.files.file_creation_information import FileCreationInformation

sharepoint_username = os.environ.get('ITREB_USERNAME','')
sharepoint_password = os.environ.get('ITREB_PASSWORD','')
url = 'https://balmoralaid.sharepoint.com/sites/Planner_C6xE'
relative_url = '/sites/Planner_C6xE/Power%20BI%20Dashboard/Survey%202021-05/'

def upload_to_sharepoint(src_filename, dst_filename):
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

  fci = FileCreationInformation()
  fci.overwrite = True
  fci.url = dst_filename
  with open(src_filename, 'rb') as f:
    fci.content = f.read()
  logging.info(f'Uploading {dst_filename}')
  ctx.web.get_folder_by_server_relative_url(relative_url).files.add(fci)
  ctx.execute_query()
