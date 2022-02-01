import glob
from datetime import datetime

from selenium import webdriver
import os
import time

from webdriver_manager.chrome import ChromeDriverManager


# Get latest driver
def get_latest_driver():
    DL_DIR_LIST = "C:/Users/digiworker_biz_02/.wdm/drivers/chromedriver/win32/*"
    CHROME_FILE = "chromedriver.exe"
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
    # driver.get('https://google.com')
    driver.quit()
    list_of_files = glob.glob(DL_DIR_LIST)
    latest_fld = max(list_of_files, key=os.path.getmtime)
    latest_file = latest_fld + "/" + CHROME_FILE
    time.sleep(5)
    return latest_file
