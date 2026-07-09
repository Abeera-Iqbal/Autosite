"""
combinefiles.py
---------------
Reads Combined_Reviews.xlsx, 1_2_3_Reviews.xlsx and BK_Category_Report.xlsx
directly from the local downloads/ folder (written by earlier steps in the
same GitHub Actions job), merges them into Final_BK_Report.xlsx, and uploads
it to the API.
"""

import os
import pandas as pd

from config import DOWNLOAD_DIR, RUN_ID, TARGET_YEAR, TARGET_MONTH
from drive_helper import upload_file
from json_export import xlsx_to_json

ARTIFACT_DIR = os.path.abspath("artifacts")
os.makedirs(ARTIFACT_DIR, exist_ok=True)

reviews_file       = os.path.join(ARTIFACT_DIR, "Combined_Reviews.xlsx")
categories_file    = os.path.join(ARTIFACT_DIR, "1_2_3_Reviews.xlsx")
total_reviews_file = os.path.join(ARTIFACT_DIR, "BK_Category_Report.xlsx")
final_file         = os.path.join(ARTIFACT_DIR, "Final_BK_Report.xlsx")

# Filename tag exposed on the Test Records API so the frontend can find this
# specific run's output (e.g. Final_BK_Report_2026-07_bk-202607-1720000000.xlsx)
PERIOD_TAG = f"{TARGET_YEAR}-{TARGET_MONTH:02d}"
UPLOAD_BASENAME = f"Final_BK_Report_{PERIOD_TAG}_{RUN_ID}"

# ── Validate all source files exist locally ───────────────────────────────────
for path in [reviews_file, categories_file, total_reviews_file]:
    if not os.path.exists(path):
       raise FileNotFoundError(
    f"Expected artifact not found: {os.path.basename(path)}\n"
    f"Make sure the previous workflows downloaded all required artifacts."
)
    print(f"Found: {os.path.basename(path)}")

# ── Read ──────────────────────────────────────────────────────────────────────
reviews_df       = pd.read_excel(reviews_file)
categories_df    = pd.read_excel(categories_file)
total_reviews_df = pd.read_excel(total_reviews_file)

# ── Write multi-sheet workbook ────────────────────────────────────────────────
with pd.ExcelWriter(final_file, engine="openpyxl") as writer:
    reviews_df.to_excel(writer, sheet_name="Review Summary", index=False)
    categories_df.to_excel(writer, sheet_name="Review Details", index=False)
    total_reviews_df.to_excel(writer, sheet_name="Category Report", index=False)

print("Final_BK_Report.xlsx created:", final_file)

# ── JSON export ────────────────────────────────────────────────────────────────
final_json = xlsx_to_json(final_file)

# ── Upload to API (tagged with period + run id) ──────────────────────────────
upload_file(final_file, upload_name=f"{UPLOAD_BASENAME}.xlsx")
upload_file(final_json, upload_name=f"{UPLOAD_BASENAME}.json")

print("Final_BK_Report created and uploaded successfully.")
print("Output:", final_file, "/", final_json)
