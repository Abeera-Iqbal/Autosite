"""
drive_helper.py – File upload/download using the Test Records API.
Base URL: https://distapi.cybussolutions.com
"""

import os
import requests

API_BASE_URL = "https://distapi.cybussolutions.com/api/v1/test-records"

# Output files produced by this pipeline — never treat these as raw inputs
OUTPUT_FILE_NAMES = {
    "Combined_Reviews.xlsx",
    "1_2_3_Reviews.xlsx",
    "BK_Category_Report.xlsx",
    "Final_BK_Report.xlsx",
}


def upload_file(local_path: str, upload_name: str = None):
    """
    Uploads a local file to the API server.
    upload_name: override the filename sent to the API (e.g. to set IRCR_/DCR_ prefix).
    """
    filename = upload_name or os.path.basename(local_path)

    if local_path.endswith(".xlsx"):
        mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif local_path.endswith(".png"):
        mime_type = "image/png"
    elif local_path.endswith(".json"):
        mime_type = "application/json"
    else:
        mime_type = "application/octet-stream"

    try:
        with open(local_path, "rb") as f:
            response = requests.post(
                f"{API_BASE_URL}/upload",
                files={"file": (filename, f, mime_type)},
                timeout=120,
            )

        if response.status_code == 200:
            data = response.json().get("data", {})
            print(f"Uploaded to API: {filename} (saved as: {data.get('savedAs')})")
            return data
        else:
            print(f"Upload failed [{response.status_code}]: {response.text}")
            return None

    except Exception as e:
        print(f"Upload error for {filename}: {e}")
        return None



def list_files():
    """
    Returns all uploaded files from the API, sorted newest first.
    Each entry has: fileName, size, monthName, year, uploadedAt, downloadUrl
    NOTE: fileName = original name, downloadUrl contains the saved/internal name.
    """
    try:
        response = requests.get(f"{API_BASE_URL}/files", timeout=30)
        if response.status_code == 200:
            return response.json().get("data", [])
        else:
            print(f"List files failed [{response.status_code}]: {response.text}")
            return []
    except Exception as e:
        print(f"List files error: {e}")
        return []


def _saved_name(entry: dict) -> str:
    """
    Extracts the internal saved filename from the downloadUrl.
    e.g. '/api/v1/test-records/files/download/file-1781273016415-766373332.xlsx'
         → 'file-1781273016415-766373332.xlsx'
    """
    return entry["downloadUrl"].rsplit("/", 1)[-1]


def download_file(entry_or_saved_name, dest_path: str) -> bool:
    """
    Downloads a file to dest_path.
    Accepts either:
      - a file entry dict from list_files() (uses downloadUrl directly)
      - a plain saved filename string (builds the URL itself)
    Returns True on success, False on failure.
    """
    if isinstance(entry_or_saved_name, dict):
        url = f"https://distapi.cybussolutions.com{entry_or_saved_name['downloadUrl']}"
        label = entry_or_saved_name.get("fileName", url)
    else:
        url   = f"{API_BASE_URL}/files/download/{entry_or_saved_name}"
        label = entry_or_saved_name

    try:
        response = requests.get(url, stream=True, timeout=120)
        if response.status_code == 200:
            os.makedirs(os.path.dirname(os.path.abspath(dest_path)), exist_ok=True)
            with open(dest_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            print(f"Downloaded: {label} -> {dest_path}")
            return True
        else:
            print(f"Download failed [{response.status_code}]: {label}")
            return False
    except Exception as e:
        print(f"Download error for {label}: {e}")
        return False


def _has_rating_column(local_path: str) -> bool:
    """Returns True if the XLSX has a 'Rating' column — i.e. it's a review file."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(local_path, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        headers = [
            str(cell.value).strip()
            for cell in next(ws.iter_rows(max_row=1))
            if cell.value is not None
        ]
        wb.close()
        result = "Rating" in headers
        print(f"  {'OK' if result else '--'} {os.path.basename(local_path)} "
              f"({'has' if result else 'no'} Rating) | headers: {headers[:6]}")
        return result
    except Exception as e:
        print(f"  Could not inspect {local_path}: {e}")
        return False


def download_review_files(dest_dir: str) -> tuple:
    """
    Downloads the two XLSX files that contain a 'Rating' column (IRCR + DCR).
    Uses downloadUrl (not fileName) to fetch each file correctly.
    Returns (ircr_local_path, dcr_local_path).
    """
    os.makedirs(dest_dir, exist_ok=True)
    all_files = list_files()

    # Exclude known output files by their original fileName
    candidates = [
        f for f in all_files
        if f["fileName"].endswith(".xlsx")
        and os.path.basename(f["fileName"]) not in OUTPUT_FILE_NAMES
    ]

    if not candidates:
        raise FileNotFoundError("No candidate XLSX files found on the API server.")

    print(f"Found {len(candidates)} candidate file(s). Inspecting columns...")
    review_files = []

    for entry in candidates:
        if len(review_files) >= 2:
            break

        # Use the original fileName as the local filename so it's human-readable
        original_name = entry["fileName"]
        # Sanitize: replace spaces/parens for safe local path
        safe_name  = original_name.replace(" ", "_").replace("(", "").replace(")", "")
        local_path = os.path.join(dest_dir, safe_name)

        # Download using the entry dict (uses downloadUrl internally)
        if not os.path.exists(local_path):
            if not download_file(entry, local_path):
                continue

        if _has_rating_column(local_path):
            review_files.append(local_path)

    if len(review_files) < 2:
        raise FileNotFoundError(
            f"Could not find 2 review files with a 'Rating' column. "
            f"Found: {len(review_files)}"
        )

    print(f"IRCR: {review_files[0]}")
    print(f"DCR : {review_files[1]}")
    return review_files[0], review_files[1]


def download_latest_xlsx_files(dest_dir: str, count: int = 2,
                                skip_names: set = None,
                                name_prefix: str = None) -> list:
    """
    Downloads the `count` most recently uploaded XLSX files from the API.
    Uses downloadUrl for fetching — not fileName.
    """
    skip_names = skip_names or set()
    all_files  = list_files()

    xlsx_files = [
        f for f in all_files
        if f["fileName"].endswith(".xlsx")
        and os.path.basename(f["fileName"]) not in skip_names
        and (name_prefix is None or f["fileName"].startswith(name_prefix))
    ]

    if len(xlsx_files) < count:
        suffix = f" with prefix '{name_prefix}'" if name_prefix else ""
        raise Exception(
            f"Need at least {count} XLSX files on server{suffix}, "
            f"found: {len(xlsx_files)}"
        )

    downloaded = []
    for entry in xlsx_files[:count]:
        safe_name  = entry["fileName"].replace(" ", "_").replace("(", "").replace(")", "")
        local_path = os.path.join(dest_dir, safe_name)
        if download_file(entry, local_path):
            downloaded.append(local_path)

    return downloaded