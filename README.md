# Gainsight Automation

Python automation using selenium to input data on Gainsight

## Requirements

### Python dependencies

Run the following command to install the dependencies

pip3 install pandas openpyxl webdriver-manager selenium python-dotenv "python-dotenv[cli]"

### Credentials

Make sure you create your own .env file on the same folder as the script

Content should be:

EMAIL=youremail@domain.com\
PASS=yourpassword\
TIMELINE_URL=<https://timelineURL>\
GAINSIGHT_URL=<https://gainsightURL>

### Data Source

Use the included Gainsight_Log.xlsx example spreadsheet to enter your data
