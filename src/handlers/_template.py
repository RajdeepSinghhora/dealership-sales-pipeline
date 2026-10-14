"""
TEMPLATE for a new file type  (copy this file, rename it, e.g. test_drive.py)

1. Copy to  src/handlers/<your_name>.py
2. Fill in process() below: read the Excel, keep the columns you need, clean, and write
   the result to  <db_dir>/<your_name>.xlsx  (one row per unique key, update if it exists).
3. Register it in config/config.json, inside "PV" and/or "EV":
       {"name": "PV test drive", "handler": "<your_name>"}
   ("name" is the Excel file name without .xlsx; trailing digits like "(1)" are ignored.)
4. In Power BI, load <your_name>.xlsx and relate it to the Bookings table with Opty Id.
"""
import os

import pandas as pd

OUTPUT_FILE = "my_new_file.xlsx"      # <- change
KEY = "Opty Id"                       # <- the unique column of this file


def process(path, group, load_date, db_dir):
    df = pd.read_excel(path, sheet_name=0, dtype=object)

    # TODO 1: keep only the columns you need, e.g. df = df[["Opty Id", "Something"]]
    # TODO 2: clean (trim text, fix dates ...)
    df["Vehicle_Type"] = group
    df["Load_Date"] = pd.Timestamp(load_date)

    out = os.path.join(db_dir, OUTPUT_FILE)
    old = pd.read_excel(out, dtype=object) if os.path.exists(out) else pd.DataFrame(columns=df.columns)
    new_ids = set(df[KEY]) - set(old[KEY])
    merged = pd.concat([old[~old[KEY].isin(df[KEY])], df], ignore_index=True)
    merged.to_excel(out, index=False)    # TODO: raise handlers.FileLockedError if the file is open

    return {"rows_read": len(df), "new": len(new_ids), "updated": len(df) - len(new_ids),
            "unchanged": 0, "notes": []}
