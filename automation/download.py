"""
download.py
-----------
1. Logs into insights.bk.com via SSO.
2. Downloads the IRCR (In-Restaurant Complaint Ratio) guest responses XLSX.
3. Downloads the DCR (Digital Complaint Ratio) guest responses XLSX.
4. Uploads both files to Google Drive.
"""

import time
import os
from datetime import datetime
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from config import BK_USERNAME, BK_PASSWORD, TARGET_MONTH_LABEL, TARGET_YEAR, TARGET_MONTH


# ── Setup ─────────────────────────────────────────────────────────────────────
ARTIFACT_DIR = os.path.abspath("artifacts")
os.makedirs(ARTIFACT_DIR, exist_ok=True)

# ── Chrome (headless for GitHub Actions / Linux) ──────────────────────────────
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--disable-extensions")

prefs = {
    "download.default_directory": ARTIFACT_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
}
chrome_options.add_experimental_option("prefs", prefs)

# Use the system chromedriver installed by the workflow (avoids stale .wdm cache).
driver = webdriver.Chrome(
    service=Service("/usr/bin/chromedriver"),
    options=chrome_options
)

driver.maximize_window()
wait = WebDriverWait(driver, 30)


# ── Helpers ───────────────────────────────────────────────────────────────────

def close_popup_if_present():
    for _ in range(5):
        try:
            close_btn = driver.find_element(
                By.XPATH,
                "//div[contains(., 'Unlock more with BK Assistant')]//button"
            )
            driver.execute_script("arguments[0].click();", close_btn)
            return
        except Exception:
            time.sleep(1)


def click_sso():
    try:
        sso_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//button[contains(., 'Sign in with SSO')]"
            ))
        )
        driver.execute_script("arguments[0].click();", sso_btn)
        print("SSO button clicked")
        time.sleep(5)
        return True
    except Exception as e:
        print("SSO button not found:", e)
        return False


def click_export(button_id="FSS_IRCR_exportButton--0--trigger"):
    """
    Clicks the export dropdown button (identified by its ID) then selects XLSX.
    The button ID differs between IRCR and DCR panels.
    """
    print(f"Waiting for export button: #{button_id}")

    # Wait for the specific export button by ID
    export_btn = WebDriverWait(driver, 60).until(
        EC.element_to_be_clickable((By.ID, button_id))
    )
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});", export_btn
    )
    time.sleep(1)
    driver.execute_script("arguments[0].click();", export_btn)
    print("Export dropdown opened")

    # The button uses aria-controls to point at a listbox — wait for XLSX option
    # Try both text match and aria/role-based listbox option
    xlsx_btn = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((
            By.XPATH,
            "//*[@role='option' and contains(., 'XLSX')] | "
            "//*[contains(@class,'option') and contains(., 'XLSX')] | "
            "//*[contains(text(),'XLSX')]"
        ))
    )
    driver.execute_script("arguments[0].click();", xlsx_btn)
    print("Download started")


def wait_download(snapshot_before, timeout=180):
    """
    Wait until a NEW xlsx file appears and its size stops changing.
    Returns the downloaded file path.
    """
    start = time.time()

    while time.time() - start < timeout:

        files = [
            f for f in os.listdir(ARTIFACT_DIR)
            if f.endswith(".xlsx") and not f.startswith("~$")
        ]

        new_files = [f for f in files if f not in snapshot_before]

        if new_files:

            newest = max(
                [os.path.join(ARTIFACT_DIR, f) for f in new_files],
                key=os.path.getctime
            )

            size1 = os.path.getsize(newest)
            time.sleep(2)
            size2 = os.path.getsize(newest)

            if size1 == size2 and size1 > 0:
                print(f"Download completed: {os.path.basename(newest)} ({size2} bytes)")
                return newest

        time.sleep(1)

    raise TimeoutError("Download timeout.")


def click_close():
    try:
        close_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//div[contains(@class,'modal')]//button[last()]"
            ))
        )
        close_btn.click()
        print("Modal closed")
    except Exception as e:
        print("Close button not found:", e)


def click_franchise():
    print("Waiting for Franchise page...")

    # Wait until page is completely loaded
    WebDriverWait(driver, 60).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

    time.sleep(5)

    xpaths = [
        "//*[contains(normalize-space(),'Franchisee Success')]",
        "//button[contains(.,'Franchisee Success')]",
        "//a[contains(.,'Franchisee Success')]",
        "//span[contains(.,'Franchisee Success')]",
        "//*[contains(@aria-label,'Franchisee Success')]"
    ]

    for xpath in xpaths:
        try:
            print("Trying:", xpath)

            btn = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )

            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", btn
            )

            time.sleep(1)

            driver.execute_script("arguments[0].click();", btn)

            print("Franchise clicked")

            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((
                    By.XPATH,
                    "//*[contains(.,'Scorecard timeline')]"
                ))
            )

            print("Franchise page opened")

            return

        except Exception as e:
            print("Failed:", xpath)

    raise Exception("Could not locate Franchisee Success button.")

def click_currentmonth():
    current_month = TARGET_MONTH_LABEL
    print("Looking for:", current_month)
    time.sleep(3)

    clicked = False
    cards = driver.find_elements(By.XPATH, "//div")

    for card in cards:
        try:
            if current_month in card.text:
                open_btn = card.find_element(
                    By.XPATH,
                    ".//*[contains(text(),'Open')]"
                )
                driver.execute_script("arguments[0].click();", open_btn)
                print("Opened:", current_month)
                clicked = True
                break
        except Exception:
            continue

    if not clicked:
        raise Exception(f"{current_month} card not found on page")


def get_xlsx_files_before(snapshot):
    current = set()

    for f in os.listdir(ARTIFACT_DIR):
        if f.endswith(".xlsx") and not f.startswith("~$"):
            current.add(f)

    new_files = current - snapshot

    return [os.path.join(ARTIFACT_DIR, f) for f in new_files]

# ── Login ─────────────────────────────────────────────────────────────────────
driver.get("https://insights.bk.com/auth/login")

click_sso()

wait.until(
    EC.presence_of_element_located((By.NAME, "identifier"))
).send_keys(BK_USERNAME)

driver.find_element(By.NAME, "credentials.passcode").send_keys(BK_PASSWORD)
driver.find_element(By.CSS_SELECTOR, "input[data-type='save']").click()

time.sleep(10)
print("Login successful")

click_sso()

# ── Navigate ──────────────────────────────────────────────────────────────────
close_popup_if_present()
click_franchise()
close_popup_if_present()
click_currentmonth()

# ── Download IRCR ─────────────────────────────────────────────────────────────
snapshot_before_ircr = set(
    f for f in os.listdir(ARTIFACT_DIR)
    if f.endswith(".xlsx") and not f.startswith("~$")
)


try:
    ircr_card = wait.until(
        EC.presence_of_element_located((
            By.XPATH,
            "//div[contains(., 'In-Restaurant Complaint Ratio')]"
        ))
    )
    view_details_btn = ircr_card.find_element(
        By.XPATH,
        ".//*[contains(text(),'VIEW DETAILS')]"
    )
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});",
        view_details_btn
    )
    time.sleep(2)
    driver.execute_script("arguments[0].click();", view_details_btn)
    print("IRCR View Details clicked")
except Exception as e:
    print("Failed to click IRCR View Details:", e)

view_all_btn = wait.until(
    EC.element_to_be_clickable((
        By.XPATH,
        "//button[contains(., 'VIEW ALL GUEST RESPONSES')]"
    ))
)
driver.execute_script("arguments[0].click();", view_all_btn)
print("Guest Responses opened")
time.sleep(5)  # Allow guest responses panel to fully render

click_export("FSS_IRCR_exportButton--0--trigger")

ircr_latest = wait_download(snapshot_before_ircr)

print("Downloaded IRCR:", ircr_latest)

if ircr_latest:
    ircr_new = os.path.join(ARTIFACT_DIR, "IRCR_Reviews.xlsx")

    if os.path.exists(ircr_new):
        os.remove(ircr_new)

    os.rename(ircr_latest, ircr_new)

    print("IRCR downloaded and renamed to IRCR_Reviews.xlsx")

click_close()

# ── Download DCR ──────────────────────────────────────────────────────────────
# Build dynamic URL for the target month
current_month_url = TARGET_MONTH_LABEL.replace(" ", "%20")
dcr_url = (
    "https://app.prod.bkinsights.bkdata.rbi.tools"
    "/app/franchisee-success/my-score"
    f"#methodology=2026&restNo=&timePeriod={current_month_url}&timeWindow=month"
)

driver.get(dcr_url)
time.sleep(5)
print("DCR page loaded")

snapshot_before_dcr = set(
    f for f in os.listdir(ARTIFACT_DIR)
    if f.endswith(".xlsx") and not f.startswith("~$")
)

try:
    dcr_title = wait.until(
        EC.visibility_of_element_located((
            By.XPATH,
            "//*[contains(text(),'Digital Complaint Ratio')]"
        ))
    )
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});",
        dcr_title
    )
    time.sleep(2)

    dcr_button = dcr_title.find_element(
        By.XPATH,
        "./following::*[contains(text(),'VIEW DETAILS')][1]"
    )
    driver.execute_script("arguments[0].click();", dcr_button)
    print("DCR View Details clicked")
except Exception as e:
    print("Failed to click DCR View Details:", e)

view_all_btn = wait.until(
    EC.element_to_be_clickable((
        By.XPATH,
        "//button[contains(., 'VIEW ALL GUEST RESPONSES')]"
    ))
)
driver.execute_script("arguments[0].click();", view_all_btn)
print("Guest Responses opened")
time.sleep(5)  # Allow guest responses panel to fully render

click_export("FSS_DCR_exportButton--0--trigger")

dcr_latest = wait_download(snapshot_before_dcr)

print("Downloaded file:", dcr_latest)

try:
    temp = pd.read_excel(dcr_latest)
    print("Downloaded columns:", temp.columns.tolist())
except Exception as e:
    print("Cannot read downloaded file:", e)

if dcr_latest:
    dcr_new = os.path.join(ARTIFACT_DIR, "DCR_Reviews.xlsx")

    if os.path.exists(dcr_new):
        os.remove(dcr_new)

    os.rename(dcr_latest, dcr_new)

    print("DCR downloaded and renamed to DCR_Reviews.xlsx")

click_close()

driver.quit()
print("Download completed successfully.")
print("Artifacts stored in:", ARTIFACT_DIR)
