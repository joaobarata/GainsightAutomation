from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv

import os
import pandas as pd

#Load dotenv file ( .env file needs to be placed at the same level as the python script)
load_dotenv()
EMAIL = os.getenv('EMAIL')
PASS = os.getenv('PASS')
TIMELINE_URL = os.getenv('TIMELINE_URL')
GAINSIGHT_URL = os.getenv('GAINSIGHT_URL')

if None in (EMAIL, PASS, TIMELINE_URL, GAINSIGHT_URL ):
    print('Please create a .env file with your account credentials on the same folder as the python script')
    input('Press return to exit.')
    exit()

if not os.path.isfile('./Gainsight_Log.xlsx'):
    print('Please create a Gainsight_Log.xlsx file on the same folder as the python script')
    input('Press return to exit.')
    exit()

# Login Selectors
login_Username_Inp = (By.NAME, "loginfmt")
login_Btn = (By.CSS_SELECTOR,'.button_primary.button')
login_Pass_Inp = (By.NAME, "passwd")
login_BtnNext = (By.CSS_SELECTOR,'Input[type="submit"]')

gainsight_iframe = (By.CSS_SELECTOR ,'iframe[src^="'+GAINSIGHT_URL+'"]')

# Activity Selectors
activity_CreateBtn = (By.CLASS_NAME,'gs-add-ant')
activity_type_dropdown = (By.XPATH,"//li[contains(text(), 'TSM Update')]")

#Company Selectors
company_dropdown = (By.XPATH,"//*[@placeholder='Search Companies and Relationships']")
company_dropdownItem = (By.XPATH,"/html/body/div[1]/div[5]/div/div/div/div[1]/div[1]/ul/gs-global-search-item/li/a")

#Subject
subject_Inp= (By.XPATH,"//input[@aria-label='subject']")
#Add Date
date_inp=(By.XPATH,"//input[@placeholder='Select date']")
date_calendar_inp = (By.CLASS_NAME,'ant-calendar-input')
date_calendar_selected =(By.XPATH,"//div[@aria-selected='true']")
#Add Time
time_inp=(By.XPATH,"//input[@aria-label='time-picker-field activityDate']")
#Meeting Type
meeting_dropdown = (By.XPATH,"//nz-select[@aria-label='Ant__TSM_Meeting_Type__c']")
#Hours
hours_inp = (By.CLASS_NAME,"ant-input-number-input")
#Notes
notes_iframe=(By.XPATH,"//iframe[@aria-label='rte']")
notes_input=(By.CLASS_NAME,"gs-rte-container")
#Add
activity_save=(By.CLASS_NAME,"gs-nt-save")

#Supress webdriver experimental errors
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])

#Load chromium webdriver. It will automatically be updated if necessary
chrome_service=Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=chrome_service, options=options)

#Load the Gainsight timeline URL
driver.get(TIMELINE_URL)

wait = WebDriverWait(driver, 20)

#Wait for form and then enter the username
wait.until(EC.presence_of_element_located(login_Username_Inp)).send_keys(EMAIL)
wait.until(EC.element_to_be_clickable(login_Btn)).click()

#Wait for form and then enter the password
wait.until(EC.presence_of_element_located(login_Pass_Inp)).send_keys(PASS)
wait.until(EC.element_to_be_clickable(login_BtnNext)).click()

print("\nWaiting for 2FA confirmation\n")
longwait = WebDriverWait(driver, 120)

#Wait for the Timeline page to load
longwait.until(EC.frame_to_be_available_and_switch_to_it(gainsight_iframe))
print("Timeline loaded\n")

#switch to iframe
driver.switch_to.default_content()
driver.switch_to.frame(0)

#log item function
def logItem(string, subject, date, meetingtime, meetingtype, hours, notes):
    wait.until(EC.element_to_be_clickable(activity_CreateBtn)).click()
    driver.find_element(activity_type_dropdown[0],activity_type_dropdown[1]).click()
    driver.find_elements(company_dropdown[0],company_dropdown[1])[1].send_keys(string)
    
    wait.until(EC.element_to_be_clickable(company_dropdownItem)).click()
    #Add Subject
    driver.find_element(subject_Inp[0],subject_Inp[1]).send_keys(subject)
    #Add Date
    wait.until(EC.element_to_be_clickable(date_inp)).click()
    driver.find_element(date_calendar_inp[0],date_calendar_inp[1]).clear()
    driver.find_element(date_calendar_inp[0],date_calendar_inp[1]).send_keys('{d.month}/{d.day}/{d.year}'.format(d=date))
    wait.until(EC.element_to_be_clickable(date_calendar_selected)).click()
    #Add Time
    driver.find_element(time_inp[0],time_inp[1]).clear()
    wait.until(EC.element_to_be_clickable(time_inp)).click()
    driver.find_element(time_inp[0],time_inp[1]).send_keys(meetingtime)
    #Meeting Type
    wait.until(EC.element_to_be_clickable(meeting_dropdown)).click()
    wait.until(EC.presence_of_element_located((By.XPATH,"//li[contains(text(), ' " + meetingtype + "')]"))).click()
    #Hours
    driver.find_element(hours_inp[0],hours_inp[1]).send_keys(int(hours))
    #Notes
    driver.switch_to.frame(driver.find_element(notes_iframe[0],notes_iframe[1]))
    driver.find_element(notes_input[0],notes_input[1]).send_keys(notes)
    driver.switch_to.default_content()
    driver.switch_to.frame(0)
    #Add
    wait.until(EC.element_to_be_clickable(activity_save)).click()

df = pd.read_excel (r'./Gainsight_Log.xlsx')
df_parsed = df.dropna(subset=['Company','Subject', 'Date', 'Time', 'MeetingType', 'Hours', 'Notes'])
df_skipped = df[~df.index.isin(df_parsed.index)]

if(len(df_skipped)):
    print ('Skipped '+str(len(df_skipped)) +' rows.')
df = df_parsed
print('Valid rows to insert:')
print (df)
print()

for ind in df.index:
     company=df['Company'][ind]
     subject=df['Subject'][ind]
     date=pd.to_datetime(df['Date'][ind])
     meetingtime=df['Time'][ind]
     meetingtype=df['MeetingType'][ind]
     hours=df['Hours'][ind]
     notes=df['Notes'][ind]
     logItem(company, subject, date, meetingtime, meetingtype, hours, notes)
     print("Added ", str(hours)+"h", meetingtype, "{d.month}/{d.day}/{d.year}".format(d=date), company)
     wait.until_not(EC.presence_of_element_located((By.CLASS_NAME,"gs-header-title__context")))

os.unsetenv('EMAIL')
os.unsetenv('PASS')