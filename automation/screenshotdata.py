"""
screenshotdata.py
-----------------
1. Logs into insights.bk.com via SSO.
2. Navigates to Franchisee Success → current month scorecard.
3. Takes a full-page screenshot (my_score.png).
4. Downloads the scorecard XLSX.
5. Removes unwanted columns from the XLSX.
6. Embeds the screenshot into the XLSX.
7. Uploads the updated XLSX to Google Drive.
"""


import time
import os
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage

from config import DOWNLOAD_DIR, BK_USERNAME, BK_PASSWORD
from drive_helper import upload_file

# ── Setup ─────────────────────────────────────────────────────────────────────
ARTIFACT_DIR = os.path.abspath("artifacts")
os.makedirs(ARTIFACT_DIR, exist_ok=True)

# ── Chrome (headless for GitHub Actions / Linux) ──────────────────────────────
# Use the system chromedriver installed by the workflow (avoids stale .wdm cache).
chrome_options = Options()
chrome_options.add_argument("--headless=new")
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


def wait_download():
    timeout = 120
    elapsed = 0
    while elapsed < timeout:
        files = os.listdir(ARTIFACT_DIR)
        if any(f.endswith(".crdownload") for f in files):
            time.sleep(1)
            elapsed += 1
        else:
            break
    if elapsed >= timeout:
        raise TimeoutError("Download did not complete within 120 seconds")


def get_latest_xlsx(exclude=None):
    """Return the most recently created .xlsx in DOWNLOAD_DIR."""
    exclude = exclude or []
    files = []
    for f in os.listdir(ARTIFACT_DIR):
        if f.endswith(".xlsx") and not f.startswith("~$") and f not in exclude:
            full = os.path.join(ARTIFACT_DIR, f)
            try:
                with open(full, "rb"):
                    pass
                files.append(full)
            except Exception:
                continue
    if not files:
        raise FileNotFoundError("No valid XLSX file found in downloads")
    return max(files, key=os.path.getctime)


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

# ── Navigate to Franchisee Success ────────────────────────────────────────────
close_popup_if_present()

franchise_btn = wait.until(
    EC.element_to_be_clickable((
        By.XPATH,
        "//*[contains(text(),'Franchisee Success')]"
    ))
)
franchise_btn.click()

wait.until(
    EC.presence_of_element_located((
        By.XPATH,
        "//*[contains(., 'Scorecard timeline')]"
    ))
)

close_popup_if_present()
print("Franchisee Success page loaded")

# ── Open current month ────────────────────────────────────────────────────────
current_month = datetime.now().strftime("%b %Y")
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

time.sleep(7)

# ── Screenshot ────────────────────────────────────────────────────────────────
screenshot_path = os.path.join(ARTIFACT_DIR, "my_score.png")

try:
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(2)
    driver.execute_script("document.body.style.zoom='80%'")
    time.sleep(2)
    driver.save_screenshot(screenshot_path)
    print("Screenshot captured:", screenshot_path)
except Exception as e:
    print("Screenshot failed:", e)

# ── Export XLSX ───────────────────────────────────────────────────────────────
export_btn = wait.until(
    EC.element_to_be_clickable((
        By.XPATH,
        "//button[contains(., 'Export')]"
    ))
)
driver.execute_script("arguments[0].click();", export_btn)

xlsx_btn = wait.until(
    EC.element_to_be_clickable((
        By.XPATH,
        "//*[contains(text(),'XLSX')]"
    ))
)
driver.execute_script("arguments[0].click();", xlsx_btn)
print("Download started")

wait_download()

# Remove temp lock files
for f in os.listdir(ARTIFACT_DIR):
    if f.startswith("~$"):
        try:
            os.remove(os.path.join(ARTIFACT_DIR, f))
        except Exception:
            pass

time.sleep(5)
latest = get_latest_xlsx()
print("Latest XLSX:", latest)

# ── Process XLSX: remove columns + embed screenshot ──────────────────────────
try:
    workbook = load_workbook(latest)
    sheet = workbook.active

    columns_to_delete = [
        "FZ Name", "Restaurant Address", "State", "Zip Code", "Country", "DMA",
        "Previous Overall Star Rating (2026-04)", "Previous IRCR (2026-04)",
        "Previous DCR (2026-04)", "Previous SOS (2026-04)",
        "Previous Basics Training (2026-04)", "Previous LTO Training (2026-04)",
        "Previous Brand Standards (2026-04)", "Previous GLV Score (2026-04)",
        "Grace Period", "Grace Start Time", "Grace End Time",
        "IRCR MSC", "DCR MSC", "IRCR - SM2", "SOS CCDT", "SOS SAI", "SOS MP",
        "SOS Cheating", "Training Cheating", "LTO Cheating",
    ]

    headers_map = {
        cell.value: cell.column
        for cell in sheet[1]
    }

    col_indexes = sorted(
        [headers_map[c] for c in columns_to_delete if c in headers_map],
        reverse=True
    )

    for col_idx in col_indexes:
        sheet.delete_cols(col_idx)

    print(f"Deleted {len(col_indexes)} columns")

    # Embed screenshot
    img = ExcelImage(screenshot_path)
    img.width = 900
    img.height = 450
    sheet.add_image(img, "B17")
    print("Screenshot embedded into XLSX")

    workbook.save(latest)
    workbook.close()
    print("XLSX saved:", latest)

except Exception as e:
    print("Excel processing failed:", e)

# ── Upload to Google Drive ────────────────────────────────────────────────────
print("Processed scorecard saved:", latest)

driver.quit()
print("screenshotdata.py completed successfully")
print("Artifacts saved in:", ARTIFACT_DIR)
