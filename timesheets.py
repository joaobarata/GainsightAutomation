import os
import sys

import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

# Load dotenv file ( .env file needs to be placed at the same level as the python script)

if len(sys.argv) == 6:
    EMAIL = sys.argv[1]
    PASS = sys.argv[2]
    TIMELINE_URL = sys.argv[3]
    GAINSIGHT_URL = sys.argv[4]
    ACTIVITY_TYPE = sys.argv[5]
else:
    load_dotenv()
    EMAIL = os.getenv('EMAIL', '')
    PASS = os.getenv('PASS', '')
    TIMELINE_URL = os.getenv('TIMELINE_URL', 'https://outsystems--jbcxm.eu43.visual.force.com/apex/GainsightNXT#timeline')
    GAINSIGHT_URL = os.getenv('GAINSIGHT_URL', 'https://outsystems.gainsightcloud.com')
    ACTIVITY_TYPE = 'TSM Update'

SETTINGS = (EMAIL, PASS, TIMELINE_URL, GAINSIGHT_URL)
PARAMS = ('EMAIL', 'PASS', 'TIMELINE_URL', 'GAINSIGHT_URL')

try:
    pos = SETTINGS.index(None)
    print('Please make sure you have a .env file with the a value for the propery ' + PARAMS[pos] + ' on the same folder as the python script')
    input('Press return to exit.')
    exit()
except ValueError:
    print('Starting')

if not os.path.isfile('./Gainsight_Log.xlsx'):
    print('Please create a Gainsight_Log.xlsx file on the same folder as the python script')
    input('Press return to exit.')
    exit()

# Login Selectors

login_Username_Inp = (By.CSS_SELECTOR, 'Input[name="loginfmt"]')
login_Btn = (By.CSS_SELECTOR, '.button_primary.button')
login_Pass_Inp = (By.NAME, "passwd")
login_BtnNext = (By.CSS_SELECTOR, 'Input[type="submit"]')

gainsight_iframe = (By.CSS_SELECTOR, 'iframe[src^="'+GAINSIGHT_URL+'"]')

# Activity Selectors
activity_CreateBtn = (By.CLASS_NAME, 'gs-add-ant')
activity_type_dropdown = (By.CSS_SELECTOR,'li.ant-select-dropdown-menu-item[title="'+ACTIVITY_TYPE+'"]')

# Company Selectors
company_dropdown = (By.XPATH, "//*[@placeholder='Search Companies and Relationships']")
company_dropdownItem = (By.CSS_SELECTOR, "a.gs-global-srch-res")

# Subject
subject_Inp = (By.XPATH, "//input[@aria-label='subject']")
# Add Date
date_inp = (By.CSS_SELECTOR, "input.ant-calendar-picker-input")
date_calendar_inp = (By.CLASS_NAME, 'ant-calendar-input')
date_calendar_selected = (By.CSS_SELECTOR, ".ant-calendar-date[aria-selected='true']")
# Add Time
time_inp_to_click=(By.CLASS_NAME,"ant-time-picker-input")
time_inp=(By.CLASS_NAME,"ant-time-picker-panel-input")
# Meeting Type
meeting_dropdown = (By.XPATH, "//nz-select[@aria-label='Ant__TSM_Meeting_Type__c']")
# Hours
hours_inp = (By.CLASS_NAME, "ant-input-number-input")
# Notes
notes_iframe = (By.XPATH, "//iframe[@aria-label='rte']")
notes_input = (By.CLASS_NAME, "gs-rte-container")
# Add
activity_save = (By.CLASS_NAME, "gs-nt-save")
activity_discard = (By.CLASS_NAME, "gs-nt-close")

# Supress webdriver experimental errors
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])

# Load chromium webdriver. It will automatically be updated if necessary
chrome_service = Service()
driver = webdriver.Chrome(service=chrome_service, options=options)

# Load the Gainsight timeline URL
driver.get(TIMELINE_URL)

wait = WebDriverWait(driver, 20, 1)

# Wait for form and then enter the username
wait.until(EC.element_to_be_clickable(login_Username_Inp)).send_keys(EMAIL, Keys.RETURN)

# Wait for form and then enter the password
if (len(PASS) > 0):
    wait.until(EC.element_to_be_clickable(login_Pass_Inp)).send_keys(PASS, Keys.RETURN)

print("\nWaiting for 2FA confirmation\n")
longwait = WebDriverWait(driver, 120, 1)

# Wait for the Timeline page to load
longwait.until(EC.frame_to_be_available_and_switch_to_it(gainsight_iframe))
print("Timeline loaded\n")

# switch to iframe
driver.switch_to.default_content()
driver.switch_to.frame(0)

# log item function


def logItem(string, subject, date, meetingtime, meetingtype, hours, notes):
    wait.until(EC.element_to_be_clickable(activity_CreateBtn)).click()
    driver.find_element(activity_type_dropdown[0], activity_type_dropdown[1]).click()
    driver.find_element(company_dropdown[0], company_dropdown[1]).send_keys(string)

    wait.until(EC.element_to_be_clickable(company_dropdownItem)).click()
    # Add Subject
    driver.find_element(subject_Inp[0], subject_Inp[1]).send_keys(subject)
    # Add Date
    driver.find_element(date_inp[0],date_inp[1]).click()
    driver.find_element(date_calendar_inp[0], date_calendar_inp[1]).clear()
    driver.find_element(date_calendar_inp[0], date_calendar_inp[1]).send_keys('{d.month}/{d.day}/{d.year}'.format(d=date))
    wait.until(EC.element_to_be_clickable(date_calendar_selected)).click()
    # Add Time
    driver.find_element(time_inp_to_click[0],time_inp_to_click[1])
    wait.until(EC.element_to_be_clickable(time_inp_to_click)).click()
    driver.find_element(time_inp[0],time_inp[1]).clear()
    wait.until(EC.element_to_be_clickable(time_inp)).click()
    driver.find_element(time_inp[0],time_inp[1]).send_keys(meetingtime, Keys.RETURN)
    # Meeting Type
    wait.until(EC.element_to_be_clickable(meeting_dropdown)).click()
    wait.until(EC.presence_of_element_located((By.XPATH,'//*/ul/li/span[contains(text(),"'+ meetingtype+'")]'))).click()
    # Hours
    driver.find_element(hours_inp[0], hours_inp[1]).send_keys(str(hours))
    # Notes
    driver.switch_to.frame(driver.find_element(notes_iframe[0], notes_iframe[1]))
    driver.find_element(notes_input[0], notes_input[1]).send_keys(notes)
    driver.switch_to.default_content()
    driver.switch_to.frame(0)
    # Add
    wait.until(EC.element_to_be_clickable(activity_save)).click()
    #wait.until(EC.element_to_be_clickable(activity_discard)).click()


df = pd.read_excel(r'./Gainsight_Log.xlsx')
df_parsed = df.dropna(subset=['Company', 'Subject', 'Date', 'Time', 'MeetingType', 'Hours', 'Notes'])
df_skipped = df[~df.index.isin(df_parsed.index)]

if (len(df_skipped)):
    print('Skipped '+str(len(df_skipped)) + ' rows.')
df = df_parsed
print('Valid rows to insert:')
print(df)
print()

for ind in df.index:
    company = df['Company'][ind]
    subject = df['Subject'][ind]
    date = pd.to_datetime(df['Date'][ind])
    meetingtime = df['Time'][ind]
    meetingtype = df['MeetingType'][ind]
    hours = pd.to_numeric(df['Hours'][ind])
    notes = df['Notes'][ind]
    logItem(company, subject, date, meetingtime, meetingtype, hours, notes)
    print("Added ", str(hours)+"h", meetingtype, "{d.month}/{d.day}/{d.year}".format(d=date), company)
    wait.until_not(EC.presence_of_element_located((By.CLASS_NAME, "gs-header-title__context")))

os.unsetenv('EMAIL')
os.unsetenv('PASS')

input('All entries added. Press return to exit.')
driver.quit()
sys.exit(0)