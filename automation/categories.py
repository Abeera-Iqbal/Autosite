"""
categories.py
-------------
1. Downloads the 3 latest XLSX files from the API, inspects columns,
   and picks only those with a 'Rating' column (IRCR + DCR review files).
2. Filters 1-2-3 star reviews, auto-detects complaint category via keyword matching.
3. Deduplicates against the existing 1_2_3_Reviews.xlsx (also fetched from API if present).
4. Stacks new reviews on top, saves and uploads to the API.
"""

import os
import time
import pandas as pd
import openpyxl

from config import DOWNLOAD_DIR, RESTAURANT_MAPPING, MANAGER_MAPPING
from drive_helper import upload_file, list_files, download_file

time.sleep(5)

ARTIFACT_DIR = os.path.abspath("artifacts")
os.makedirs(ARTIFACT_DIR, exist_ok=True)

OUTPUT_FILE = os.path.join(ARTIFACT_DIR, "1_2_3_Reviews.xlsx")

SKIP_FILES = {
    "1_2_3_Reviews.xlsx",
    "Combined_Reviews.xlsx",
    "BK_Category_Report.xlsx",
    "Final_BK_Report.xlsx",
}

RATING_COL = "Rating"


def has_rating_column(local_path: str) -> bool:
    """Returns True if the first sheet has a 'Rating' header."""
    try:
        wb = openpyxl.load_workbook(local_path, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        headers = [
            str(cell.value).strip()
            for cell in next(ws.iter_rows(max_row=1))
            if cell.value is not None
        ]
        wb.close()
        result = RATING_COL in headers
        print(f"  {'✓' if result else '✗'} {os.path.basename(local_path)} — "
              f"{'has' if result else 'no'} Rating column  |  headers: {headers[:6]}")
        return result
    except Exception as e:
        print(f"  Could not inspect {os.path.basename(local_path)}: {e}")
        return False


# ── Fetch file list and download 3 latest non-output XLSX files ───────────────
print("Loading review files...")

file1 = os.path.join(ARTIFACT_DIR, "IRCR_Reviews.xlsx")
file2 = os.path.join(ARTIFACT_DIR, "DCR_Reviews.xlsx")

if not os.path.exists(file1):
    raise FileNotFoundError(f"{file1} not found")

if not os.path.exists(file2):
    raise FileNotFoundError(f"{file2} not found")

print("IRCR:", file1)
print("DCR :", file2)

# ── Read + clean ──────────────────────────────────────────────────────────────
df1 = pd.read_excel(file1)

print("\n========== IRCR ==========")
print("Columns:", df1.columns.tolist())
print(df1.head())

if RATING_COL not in df1.columns:
    raise Exception(f"IRCR file does not contain '{RATING_COL}' column")

df1[RATING_COL] = pd.to_numeric(df1[RATING_COL], errors="coerce")


# ---------------- DCR ----------------

dcr_available = False

try:
    df2 = pd.read_excel(file2)

    print("\n========== DCR ==========")
    print("Columns:", df2.columns.tolist())
    print(df2.head())

    if not df2.empty and RATING_COL in df2.columns:
        df2[RATING_COL] = pd.to_numeric(df2[RATING_COL], errors="coerce")
        dcr_available = True
        print("DCR data found.")
    else:
        print("DCR file is empty. Continuing with IRCR only.")
        df2 = pd.DataFrame(columns=df1.columns)

except Exception as e:
    print("Unable to read DCR:", e)
    df2 = pd.DataFrame(columns=df1.columns)

# ── Filter 1, 2, 3 star reviews ───────────────────────────────────────────────
filtered1 = df1[df1[RATING_COL].isin([1, 2, 3])].copy()

if dcr_available:
    filtered2 = df2[df2[RATING_COL].isin([1, 2, 3])].copy()
else:
    filtered2 = pd.DataFrame(columns=filtered1.columns)

new_reviews = pd.concat([filtered1, filtered2], ignore_index=True)

# ── Category keyword detection ────────────────────────────────────────────────
CATEGORY_KEYWORDS = {
    "Food": [
        "tasty", "delicious", "bland", "stale", "fresh", "juicy",
        "overcooked", "undercooked", "flavor", "quality", "burger",
        "fries", "cold food", "hot food", "meal", "chicken", "drink",
        "burnt", "raw", "taste", "food",
    ],
    "Customer Service": [
        "rude staff", "friendly", "helpful", "behavior", "manager",
        "support", "attitude", "service", "cashier", "ignored",
        "customer service", "staff", "unprofessional", "respectful",
        "cooperative", "employee",
    ],
    "Order Compliance": [
        "wrong order", "missing item", "correct order", "customization",
        "mix-up", "forgot", "missing", "incorrect", "no sauce",
        "missing fries", "order accuracy", "didn't receive",
        "incomplete order",
    ],
    "Operations": [
        "dirty", "clean", "hygiene", "smell", "messy", "sanitary",
        "bathroom", "restroom", "long wait", "quick service", "slow",
        "queue", "delay", "waiting", "crowded", "environment",
        "maintenance", "ac not working", "table dirty",
    ],
    "Brand Expectation": [
        "disappointing", "terrible", "worst", "never again", "unhappy",
        "expected better", "not worth it", "poor experience",
        "below standard",
    ],
    "Delivery Experience": [
        "late delivery", "fast delivery", "rider", "packaging",
        "delayed", "on time", "tracking", "delivery",
    ],
    "Price & Value": [
        "expensive", "overpriced", "cheap", "value for money",
        "worth it", "pricing", "price", "portion size",
    ],
}


def detect_category(comment):
    if pd.isna(comment):
        return "Just Rating"
    comment = str(comment).lower().strip()
    if not comment:
        return "Just Rating"
    for category, keywords in CATEGORY_KEYWORDS.items():
        if any(kw in comment for kw in keywords):
            return category
    return "Brand Expectation"


# ── Enrich columns ────────────────────────────────────────────────────────────
new_reviews["Restaurant Name"] = new_reviews["Restaurant #"].map(RESTAURANT_MAPPING)
new_reviews["Area Manager"]    = new_reviews["Restaurant #"].map(MANAGER_MAPPING)
new_reviews["Category"]        = new_reviews["Comment"].apply(detect_category)

# ── Ensure all required columns exist ────────────────────────────────────────
REQUIRED_COLUMNS = [
    "Restaurant #", "Restaurant Name", "Area Manager",
    "Rating", "Comment", "Source", "ACR Inclusion",
    "Review Date", "Available on BK site",
    "Category", "Shift", "ACTION PLAN", "CLOSURE",
]

for col in REQUIRED_COLUMNS:
    if col not in new_reviews.columns:
        new_reviews[col] = ""

if "Date Exported" in new_reviews.columns:
    new_reviews["Available on BK site"] = new_reviews["Date Exported"]

new_reviews = new_reviews[REQUIRED_COLUMNS]

# ── Load existing 1_2_3_Reviews.xlsx from API (if it exists) ──────────────────
old_reviews = pd.DataFrame(columns=REQUIRED_COLUMNS)

existing_file = os.path.join(ARTIFACT_DIR, "1_2_3_Reviews.xlsx")

if os.path.exists(existing_file):
    try:
        old_reviews = pd.read_excel(existing_file)
        print(f"Loaded existing report ({len(old_reviews)} rows)")
    except Exception as e:
        print("Error reading existing report:", e)
else:
    print("No existing 1_2_3_Reviews.xlsx found. Starting fresh.")

# ── Deduplicate ───────────────────────────────────────────────────────────────
def make_key(df):
    return (
        df["Restaurant #"].astype(str).fillna("")
        + "|"
        + df["Review Date"].astype(str).fillna("")
        + "|"
        + df["Comment"].astype(str).fillna("")
    )


old_reviews["_key"] = make_key(old_reviews)
new_reviews["_key"] = make_key(new_reviews)

only_new = new_reviews[
    ~new_reviews["_key"].isin(old_reviews["_key"])
].copy()

print(f"New reviews to add: {len(only_new)}")

old_reviews.drop(columns=["_key"], inplace=True, errors="ignore")
only_new.drop(columns=["_key"], inplace=True, errors="ignore")

# ── Stack: new on top ─────────────────────────────────────────────────────────
final_df = pd.concat([only_new, old_reviews], ignore_index=True)

# ── Save ──────────────────────────────────────────────────────────────────────
try:
    final_df.to_excel(OUTPUT_FILE, index=False)
    print(f"Saved {len(final_df)} total rows to:", OUTPUT_FILE)
except PermissionError:
    print("ERROR: Please close the Excel file first")
    raise

# ── Upload to API ─────────────────────────────────────────────────────────────
print("categories.py completed successfully")
print("Output:", OUTPUT_FILE)
