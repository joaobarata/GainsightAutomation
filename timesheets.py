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
wait.until(EC.presence_of_element_located((By.NAME, "loginfmt"))).send_keys(EMAIL)
driver.find_element(By.CSS_SELECTOR,'.win-button.button_primary.button.ext-button.primary.ext-primary').click()

#Wait for form and then enter the password
wait.until(EC.presence_of_element_located((By.ID, "passwordInput"))).send_keys(PASS)
driver.find_element(By.ID,'submitButton').click()

print()
print("Waiting for 2FA confirmation")
longwait = WebDriverWait(driver, 120)

#Wait for the Timeline page to load
longwait.until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR ,'iframe[src^="'+GAINSIGHT_URL+'"]')))
wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.gs-create-activity button')))

print()
print("Timeline loaded")

#switch to iframe
driver.switch_to.default_content()
driver.switch_to.frame(0)

#log item function
def logItem(string, subject, date, meetingtime, meetingtype, hours, notes):
    driver.find_element(By.CLASS_NAME,'gs-add-ant').click()
    driver.find_element(By.XPATH,"//li[contains(text(), 'TSM Update')]").click()
    driver.find_elements(By.XPATH,"//*[@placeholder='Search Companies and Relationships']")[1].send_keys(string)
    
    wait.until(EC.element_to_be_clickable((By.XPATH,"/html/body/div[1]/div[5]/div/div/div/div[1]/div[1]/ul/gs-global-search-item/li/a"))).click()
    #Add Subject
    driver.find_element(By.XPATH,"//input[@aria-label='subject']").send_keys(subject)
    #Add Date
    driver.find_element(By.XPATH,"//input[@placeholder='Select date']").click()
    driver.find_element(By.CLASS_NAME,'ant-calendar-input').clear()
    driver.find_element(By.CLASS_NAME,'ant-calendar-input').send_keys('{d.month}/{d.day}/{d.year}'.format(d=date))
    driver.find_element(By.XPATH,"//div[@aria-selected='true']").click()
    #Add Time
    driver.find_element(By.XPATH,"//input[@aria-label='time-picker-field activityDate']").clear()
    driver.find_element(By.XPATH,"//input[@aria-label='time-picker-field activityDate']").click()
    driver.find_element(By.XPATH,"//input[@aria-label='time-picker-field activityDate']").send_keys(meetingtime)
    #Meeting Type
    driver.find_element(By.XPATH,"//nz-select[@aria-label='Ant__TSM_Meeting_Type__c']").click()
    wait.until(EC.presence_of_element_located((By.XPATH,"//li[contains(text(), ' " + meetingtype + "')]"))).click()
    #Hours
    driver.find_element(By.CLASS_NAME,"ant-input-number-input").send_keys(int(hours))
    #Notes
    driver.switch_to.frame(driver.find_element(By.XPATH,"//iframe[@aria-label='rte']"))
    driver.find_element(By.CLASS_NAME,"gs-rte-container").send_keys(notes)
    driver.switch_to.default_content()
    driver.switch_to.frame(0)
    #Add
    driver.find_element(By.CLASS_NAME,"gs-nt-save").click()

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
     date=df['Date'][ind]
     meetingtime=df['Time'][ind]
     meetingtype=df['MeetingType'][ind]
     hours=df['Hours'][ind]
     notes=df['Notes'][ind]
     logItem(company, subject, date, meetingtime, meetingtype, hours, notes)
     print("Added ", company, " ", date)
     wait.until_not(EC.presence_of_element_located((By.CLASS_NAME,"gs-header-title__context")))

os.unsetenv('EMAIL')
os.unsetenv('PASS')