import pandas as pd
import pyautogui
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

start_row = 2
iterations = 1

SERVICE_ACCOUNT_FILE = 'YOUR JSON'  # ADD JSON
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=credentials)

SPREADSHEET_ID = 'YOUR SHEET ID'  # ADD SHEET ID

HEADER_RANGE = 'Sheet1!A1:E1'
header_result = service.spreadsheets().values().get(
    spreadsheetId=SPREADSHEET_ID, range=HEADER_RANGE).execute()
header_values = header_result.get('values', [])[0]

DATA_RANGE = f'Sheet1!A{str(start_row)}:E'
data_result = service.spreadsheets().values().get(
    spreadsheetId=SPREADSHEET_ID, range=DATA_RANGE).execute()
data_values = data_result.get('values', [])

df = pd.DataFrame(data_values, columns=header_values)

download_dir = r'E:\Desktop\python\GoogleSheet\download'
download_button_image = r'E:\Desktop\python\GoogleSheet\download_button.png'
drive_download_image = r'E:\Desktop\python\GoogleSheet\google_drive_download.png'
onedrive_error_image = r'E:\Desktop\python\GoogleSheet\drive_error_image.png'

chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": True,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
chrome_options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(options=chrome_options)
pyautogui.PAUSE = 0.3


def websearch(xpath, description):
    print(
        f'Start Searching XPath {description} on row {num}, Link {url}, Sku {sku}')
    try:
        element = WebDriverWait(driver, 9999).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        element.click()
    except Exception as e:
        print(f"Searching for XPath: {e}")


def search_picture(image, description):
    print(
        f'Start Searching image {description} on row {num}, Link {url}, Sku {sku}')
    image_location = None
    while image_location is None:
        try:
            image_location = pyautogui.locateOnScreen(image, confidence=0.8)
            if image_location:
                pyautogui.click(image_location)
                print(f'{button_image_path} is founded')
                break
        except Exception as e:
            time.sleep(0.3)
            print(f'Searching {description}:{e}')


def save_file_first_time():
    print('Searching for download box')
    while True:
        try:
            button_location = pyautogui.locateOnScreen(
                download_button_image, confidence=0.8)
            if button_location:
                print('Download box found')
                pyautogui.typewrite(str(sku))
                pyautogui.press('f4')
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.typewrite(download_dir)
                pyautogui.press('enter')
                pyautogui.click(button_location)
                print(f'Current No:{num}, Current SKU:{sku}')
                break
        except Exception as e:
            print(f'Searching "Save" Buttin: {e}')
        time.sleep(0.3)


def save_file():
    print('Searching for download box')
    while True:
        try:
            button_location = pyautogui.locateOnScreen(
                download_button_image, confidence=0.8)
            if button_location:
                print('Download box found')
                pyautogui.typewrite(str(sku))
                pyautogui.click(button_location)
                print(f'Current No:{num}, Current SKU:{sku}')
                break
        except Exception as e:
            print(f'Searching "Save" Buttin: {e}')
        time.sleep(0.3)


def update_google_sheet():
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f'Sheet1!B{index + start_row}',
        valueInputOption='RAW',
        body={'values': [['complete']]}
    ).execute()


for index, row in df.iterrows():
    url = row['property_photo']
    sku = row['sku']
    num = row['no']

    driver.get(url)
    time.sleep(1)
    if 'https://1drv.ms' in url:
        while True:
            try:
                onedrive_not_error = driver.find_element(
                    By.XPATH, '//button[@type="button" and @role="menuitem" and @name="Download"]')
                if onedrive_not_error:
                    break
            except NoSuchElementException:
                onedrive_error_message = pyautogui.locateOnScreen(
                    onedrive_error_image, confidence=0.8)
                if onedrive_error_message:
                    driver.refresh()
            time.sleep(1)
        websearch(
            '//button[@type="button" and @role="menuitem" and @name="Download"]', 'OneDrive')

    elif 'https://photos' in url:
        websearch(
            '//div[@role="button" and @class="U26fgb JRtysb WzwrXb YI2CVc G6iPcb"]', 'Google Photo')
        time.sleep(2)
        websearch('//span[@class="z80M1 o7Osof" and @jsname="j7LFlb" and @aria-label="Download all"]',
                  'Google Photo')

    elif 'https://drive.' in url:
        search_picture(drive_download_image, "Google Drive")

    time.sleep(1)

    if iterations == 1:
        save_file_first_time()
    else:
        save_file()

    print(
        f'Download Finished on row {num}, Link {url}, Sku {sku}')

    update_google_sheet()
    print(
        f'Sheet Updated on row {num}, Sku {sku}')
    iterations += 1

if __name__ == "__main__":
    main_loop()
