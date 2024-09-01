from google.oauth2 import service_account
from googleapiclient.discovery import build
import pandas as pd
import os
import re

# Google API setup
SERVICE_ACCOUNT_FILE = r'E:\Desktop\python\rent_receipt_auto_update\service_account.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive']

creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=creds)
sheet_service = build('sheets', 'v4', credentials=creds)

google_sheet_id = '1RB7FiUGABTNIUAjYbuNo-Dvdp5f9_Mo_HmyWvPdCFrM'
FOLDER_ID = '184bBAn5tNL901LInyWB1Ik9wujI7oj7H'

# Fetch Google Sheet headers and data
HEADER_RANGE = 'Sheet1!A1:F1'
header_result = sheet_service.spreadsheets().values().get(
    spreadsheetId=google_sheet_id, range=HEADER_RANGE).execute()
header_values = header_result.get('values', [])[0]

DATA_RANGE = 'Sheet1!A2:F'
data_result = sheet_service.spreadsheets().values().get(
    spreadsheetId=google_sheet_id, range=DATA_RANGE).execute()
data_values = data_result.get('values', [])

original_df = pd.DataFrame(data_values, columns=header_values)
df = pd.DataFrame(data_values, columns=header_values)


# Ensure 'link' column exists before updating
if 'link' not in df.columns:
    df['link'] = None


def clean_filename(name):
    # Attempt to extract only the room number part of the filename
    match = re.search(r'\d+[-/]\d+', name)
    if match:
        return match.group().replace('-', '/')
    return name


# Fetch files from Google Drive
query = f"'{FOLDER_ID}' in parents and mimeType != 'application/vnd.google-apps.folder'"
results = drive_service.files().list(
    q=query, fields="files(name, webViewLink)").execute()

slip_name = results.get('files', [])
file_name_and_link = {clean_filename(os.path.splitext(file['name'])[
    0]): file['webViewLink'] for file in slip_name}

# How to seperate keys and values
# file_names = list(file_name_and_link.keys())
# file_links = list(file_name_and_link.values())

for file_name, link in file_name_and_link.items():
    df.loc[df['unit_number'].str.contains(file_name, regex=False, na=False), [
        'paid', 'link']] = ['V', link]

change_cell = {}

# for row_idx, (original_df, df) in enmerate(zip(original_df.itertuples(index=False), df.itertuples(index=False)), )
print(df)
print(original_df)
