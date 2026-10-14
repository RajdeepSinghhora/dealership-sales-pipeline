from datetime import datetime

import pandas as pd

from handlers import open_pipeline as op


def test_form_and_age_band():
    assert op._form("02 Open Green Form") == ("Green Form (enquiry)", 1)
    assert op._form("02 Open Pink Form")[1] == 2
    assert op._form("03 Open Yellow / Booked")[1] == 3
    assert op._form("04 Vehicle Allotted")[1] == 4
    assert op._band(0) == ("0-7 days", 1)
    assert op._band(31) == ("31-60 days", 4)
    assert op._band(400) == ("90+ days", 6)


def test_financial_year_starts_in_april():
    assert op.fy_start(datetime(2026, 3, 31)) == pd.Timestamp(2025, 4, 1)
    assert op.fy_start(datetime(2026, 4, 1)) == pd.Timestamp(2026, 4, 1)


def test_stale_enquiry_and_long_wait():
    df = pd.DataFrame({
        "Opty Id": ["E1", "E2", "B1", "B2"],
        "Sales Stage": ["02 Open Green Form", "02 Open Green Form", "03 Open Yellow / Booked", "04 Vehicle Allotted"],
        "Opty Created Date": [datetime(2026, 8, 1), datetime(2026, 10, 1), datetime(2026, 8, 1), datetime(2026, 9, 1)],
        "YF Date": [None, None, datetime(2026, 8, 2), datetime(2026, 10, 1)],
        "Employee Full Name": "ROHAN MEHTA", "PPL": "Kestrel", "PL": "Kestrel 1.2 Plus", "Dealer Code": "3999001",
        "Mobile No": "1234567890", "Division": "3999001-Sales-Central-Demo"})
    d = op.clean_open(df, "PV", datetime(2026, 10, 12)).set_index("Opty Id")
    assert d.loc["E1", "Is_Stale_Enquiry"] == 1 and d.loc["E2", "Is_Stale_Enquiry"] == 0
    assert d.loc["B1", "Is_Long_Wait"] == 1 and d.loc["B2", "Is_Long_Wait"] == 0
    assert d.loc["B1", "Wait_Days"] == 71
    assert d.loc["E1", "Lead Classification"] == "Unspecified"
    assert d.loc["E2", "Branch"] == "Central"
