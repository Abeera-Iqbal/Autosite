from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
import calendar
from datetime import datetime
import time
import os
import glob

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

load_dotenv()

import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import (
    DS_USERNAME,
    DS_PASSWORD,
    TARGET_MONTH,
    TARGET_YEAR,
    DATE_FROM_DAY,
    DATE_TO_DAY,
)

# Kept as `target_month` / `target_year` / `today_day` names below so the
# rest of this script (datepicker navigation) doesn't need to change.
target_month = TARGET_MONTH
target_year = TARGET_YEAR
today_day = DATE_TO_DAY


def move_to_current_month(driver, wait):
    """Navigate the jQuery datepicker to the current month."""
    while True:
        month_name = wait.until(
            EC.visibility_of_element_located(
                (By.CLASS_NAME, "ui-datepicker-month")
            )
        ).text.strip()

        year = int(driver.find_element(
            By.CLASS_NAME, "ui-datepicker-year"
        ).text.strip())

        month = list(calendar.month_name).index(month_name)
        print(f"Showing: {month_name} {year}")

        if month == target_month and year == target_year:
            print("Current month reached.")
            break

        if (year < target_year) or (year == target_year and month < target_month):
            btn_css = "a.ui-datepicker-next"
        else:
            btn_css = "a.ui-datepicker-prev"

        btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, btn_css)))
        driver.execute_script("arguments[0].click();", btn)
        time.sleep(0.5)

def wait_for_download(ARTIFACT_DIR, timeout=60):
    end = time.time() + timeout

    while time.time() < end:
        files = [
            f
            for f in glob.glob(os.path.join(ARTIFACT_DIR, "*.xlsx"))
            if not f.endswith(".crdownload")
            and not f.endswith(".tmp")
        ]

        if files:
            latest = max(files, key=os.path.getmtime)
            print("Download complete:", latest)
            return latest

        time.sleep(1)

    raise TimeoutError("Download timed out.")
    

# ── DRIVER ────────────────────────────────────────────────────────────────────

ARTIFACT_DIR = os.path.abspath("artifacts")
os.makedirs(ARTIFACT_DIR, exist_ok=True)

options = webdriver.ChromeOptions()

options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920,1080")

options.add_experimental_option(
    "prefs",
    {
        "download.default_directory": ARTIFACT_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
    },
)
ARTIFACT_DIR = "artifacts"
ARTIFACT_DIR = os.path.abspath("artifacts")
os.makedirs(ARTIFACT_DIR, exist_ok=True)

driver = webdriver.Chrome(
    service=Service("/usr/bin/chromedriver"),
    options=chrome_options
)
driver.maximize_window()
wait = WebDriverWait(driver, 30)


# ── LOGIN ─────────────────────────────────────────────────────────────────────

driver.get("https://dennys.onelogin.com/login2")

username_box = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input")))
username_box.clear()
username_box.send_keys(DS_USERNAME)

wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Continue')]"))).click()
print("Username entered")

password_box = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password']")))
password_box.clear()
password_box.send_keys(DS_PASSWORD)

wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Continue')]"))).click()
print("Password entered")

# ── ONELOGIN EXTENSION SKIP ───────────────────────────────────────────────────

try:
    skip_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//*[contains(text(),'skip') or contains(text(),'Skip')]")
        )
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", skip_btn)
    skip_btn.click()
    print("OneLogin extension page skipped")
except Exception:
    print("Skip page not shown")

# ── DASHBOARD ─────────────────────────────────────────────────────────────────

wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
print("Dashboard loaded")

local_tile = wait.until(
    EC.element_to_be_clickable(
        (By.XPATH, "//*[contains(text(),'Local Website and Facebook')]")
    )
)
driver.execute_script("arguments[0].scrollIntoView(true);", local_tile)
local_tile.click()
print("Local Website and Facebook clicked")

# ── MICROSOFT TRUST PAGE  (runs once only) ────────────────────────────────────

print("Waiting for Microsoft trust page...")
time.sleep(8)

if len(driver.window_handles) > 1:
    driver.switch_to.window(driver.window_handles[-1])
    print("Switched to new tab")

try:
    continue_btn = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable(
            (By.XPATH,
             "//button[contains(.,'Continue')] | //input[@value='Continue'] | //*[contains(text(),'Continue')]")
        )
    )
    driver.execute_script("arguments[0].click();", continue_btn)
    print("Microsoft Continue clicked")
except Exception as e:
    print("Continue button not found (may not be required):", e)

# ── DAC DASHBOARD ─────────────────────────────────────────────────────────────

time.sleep(15)
print("Current URL:", driver.current_url)
print("Current Title:", driver.title)

# ── NAVIGATE TO REVIEW FEED ───────────────────────────────────────────────────

reviews_menu = WebDriverWait(driver, 60).until(
    EC.element_to_be_clickable((By.XPATH, "//p[normalize-space()='Reviews']"))
)
driver.execute_script("arguments[0].scrollIntoView({block:'center'});", reviews_menu)
time.sleep(2)
driver.execute_script("arguments[0].click();", reviews_menu)
print("Reviews clicked")

reviews_feed = WebDriverWait(driver, 60).until(
    EC.element_to_be_clickable((By.XPATH, "//p[normalize-space()='Review Feed']"))
)
driver.execute_script("arguments[0].scrollIntoView({block:'center'});", reviews_feed)
time.sleep(2)
driver.execute_script("arguments[0].click();", reviews_feed)
print("Review Feed clicked")
time.sleep(5)

# ── FROM DATE ─────────────────────────────────────────────────────────────────

def open_datepicker(driver, wait, field_id):
    """
    Reliably open a jQuery datepicker on a field.
    Tries normal click first, then falls back to JS .focus() / datepicker('show').
    Returns True when the calendar div is visible.
    """
    field = wait.until(EC.presence_of_element_located((By.ID, field_id)))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", field)
    time.sleep(0.5)

    # Attempt 1 – regular JS click
    driver.execute_script("arguments[0].click();", field)
    time.sleep(0.8)

    # Check if picker appeared
    try:
        if driver.find_element(By.ID, "ui-datepicker-div").is_displayed():
            print(f"Calendar opened for #{field_id} (click)")
            return
    except Exception:
        pass

    # Attempt 2 – focus the field
    driver.execute_script("arguments[0].focus();", field)
    time.sleep(0.8)

    try:
        if driver.find_element(By.ID, "ui-datepicker-div").is_displayed():
            print(f"Calendar opened for #{field_id} (focus)")
            return
    except Exception:
        pass

    # Attempt 3 – trigger jQuery datepicker('show') directly
    driver.execute_script("$('#' + arguments[0]).datepicker('show');", field_id)
    time.sleep(0.8)

    # Final wait – up to 10 s
    WebDriverWait(driver, 10).until(
        lambda d: d.find_element(By.ID, "ui-datepicker-div").is_displayed()
    )
    print(f"Calendar opened for #{field_id} (jQuery trigger)")


open_datepicker(driver, wait, "dateFrom")
move_to_current_month(driver, wait)

wait.until(
    EC.element_to_be_clickable(
        (By.XPATH,
         f"//table[contains(@class,'ui-datepicker-calendar')]//a[text()='{DATE_FROM_DAY}']")
    )
).click()
print(f"From date selected: {DATE_FROM_DAY} of target month")

# ── TO DATE ───────────────────────────────────────────────────────────────────

# Picker sometimes hides after first selection — give it a moment
time.sleep(1)

open_datepicker(driver, wait, "dateTo")
print("To calendar opened.")

move_to_current_month(driver, wait)

wait.until(
    EC.element_to_be_clickable(
        (By.XPATH,
         f"//table[contains(@class,'ui-datepicker-calendar')]//a[text()='{today_day}']")
    )
).click()
print(f"To date selected: {today_day} of current month")

# ── EXPORT XLSX ───────────────────────────────────────────────────────────────

export_btn = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Export')]"))
)
driver.execute_script("arguments[0].click();", export_btn)

xlsx_btn = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//*[contains(text(),'XLSX')]"))
)
driver.execute_script("arguments[0].click();", xlsx_btn)

ok_button = wait.until(
    EC.element_to_be_clickable(
        (By.XPATH, "//div[contains(@class,'modal')]//button[normalize-space()='OK']")
    )
)
time.sleep(2)

ok_button.click()
print("Download started")

# Wait until the file actually lands in the download folder
wait_for_download(ARTIFACT_DIR, timeout=60)

print("Navigation completed successfully")
# driver.quit()
