import codecs
import time
from datetime import datetime
import pandas as pd


from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait

# MFZC_URL = "https://mflowz.daj.co.jp/MFZC/Account/Login"    # <--- 本チャンのサインイン画面
MFZC_URL = "https://mflowz.daj.co.jp/MFZC/100"

MFZC_VERIFY_URL = "https://mflowz-verify.daj.co.jp/MFZC/Account/Login"  # <--- 開発環境のサインイン画面（2022.01.12）

MIGRATION_COMMENT = "MFZデータ移行ロボです。無視してください。"

my_today = datetime.today().strftime("%Y-%m-%d")

def my_sleep_click(web_el):
    try:
        web_el.click()
    except NoSuchElementException:
        print(web_el + " is not available")
    time.sleep(3)   # 引数5であったが3へ短縮


def my_send_keys(web_el, key_name):
    try:
        web_el.send_keys(key_name)
    except NoSuchElementException:
        print(web_el + " is not available")
    time.sleep(2)


def click_submit_button(driver, logger, message, web_el):
    try:
        driver.execute_script("arguments[0].click();", web_el)
        logger.info(message + ": Success")
    except NoSuchElementException:
        logger.error(message + ": Error")
    time.sleep(3)   # 引数5であったが3へ短縮

# Method to read CSV file safely
def open_csv(csv_file):
    with codecs.open(csv_file, "r", "Shift-JIS", "ignore") as file:
        return pd.read_csv(file)

def click_main(driver, logger, message, web_el):
    try:
        driver.execute_script("arguments[0].click();", web_el)
        logger.info(message + ": Success")
    except NoSuchElementException:
        logger.error(message + ": Error")
    time.sleep(3)   # 引数5であったが3へ短縮

def click_open_all_button(driver, logger, message, web_el):
    try:
        driver.execute_script("arguments[0].click();", web_el)
        logger.info(message + ": Success")
    except NoSuchElementException:
        logger.error(message + ": Error")
    time.sleep(3)   # 引数5であったが3へ短縮

