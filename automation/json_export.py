"""
json_export.py – shared helper that mirrors a finished XLSX report into a
JSON file next to it. Used by the "final deliverable" scripts (combinefiles,
dennys-record, dennys-summary) so every report produced by the pipeline is
available in both formats.
"""

import json
import pandas as pd


def xlsx_to_json(xlsx_path: str, json_path: str = None, orient: str = "records") -> str:
    """
    Reads every sheet of xlsx_path and writes it out as JSON.

    - Single-sheet workbook -> JSON is a list of row objects.
    - Multi-sheet workbook  -> JSON is {sheet_name: [row objects, ...], ...}

    Returns the path of the JSON file that was written.
    """
    json_path = json_path or xlsx_path.rsplit(".", 1)[0] + ".json"

    sheets = pd.read_excel(xlsx_path, sheet_name=None)

    if len(sheets) == 1:
        data = next(iter(sheets.values())).to_dict(orient=orient)
    else:
        data = {name: df.to_dict(orient=orient) for name, df in sheets.items()}

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, default=str)

    print(f"JSON export written: {json_path}")
    return json_path
