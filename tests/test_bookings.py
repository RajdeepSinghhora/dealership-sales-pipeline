from datetime import datetime

import pandas as pd

from handlers import bookings as bk

DIV = "3999001-Sales-Central-Demo"


def raw(rows):
    base = {c: None for c in bk.KEEP_COLS}
    base.update({"Employee Full Name": "ROHAN MEHTA", "Source(PV)": "Referral", "Sales Stage": "03 Open Yellow / Booked",
                 "Opty Created Date": datetime(2026, 10, 1), "YF Date": datetime(2026, 10, 2), "LOB": "PG1",
                 "PPL": "Kestrel", "PL": "Kestrel 1.2 Plus", "Testdrive": "Not given", "Dealer Code": "3999001",
                 "Dealer": "DEMO MOTORS PVT LTD", "Month Name": "October", "Mobile No": "1234567890"})
    out = []
    for r in rows:
        d = dict(base)
        d.update(r)
        out.append(d)
    df = pd.DataFrame(out)
    df["Division"] = df.pop("Division") if "Division" in df else DIV
    return df


def test_name_typed_twice_is_flagged_not_fixed():
    assert bk.name_repeated("RAVI KUMAR RAVI KUMAR")
    assert bk.name_repeated("Mr. RAVI KUMAR RAVI KUMAR")
    assert not bk.name_repeated("RAVI KUMAR")
    assert not bk.name_repeated("RAVI RAVI KUMAR")


def test_clean_bookings_branch_mobile_and_stage():
    df = raw([{"Opty Id": "9-1AAA", "Division": "3999001 - Sales - Lakeside  - Demo", "Mobile No": "1234567890.0",
               "Sales Stage": "06 Sale Closed", "Sale Closed": datetime(2026, 10, 5)}])
    c = bk.clean_bookings(df, "PV", datetime(2026, 10, 6), "pv.xlsx")
    r = c.iloc[0]
    assert r["Division"] == "3999001-Sales-Lakeside-Demo"
    assert r["Branch"] == "Lakeside"
    assert r["Mobile No"] == "1234567890"
    assert r["Stage_No"] == 6 and r["Is_Sale_Closed"] == 1
    assert r["Lead_to_Booking_Days"] == 1
    assert r["Financial_Year"] == "FY2026-27"


def test_same_mobile_classification():
    df = raw([
        {"Opty Id": "A", "Mobile No": "1000000001", "PL": "Kestrel 1.2 Plus"},
        {"Opty Id": "B", "Mobile No": "1000000001", "PL": "Kestrel 1.2 Plus"},          # same variant -> duplicate
        {"Opty Id": "C", "Mobile No": "1000000002", "PPL": "Marlin", "PL": "Marlin 1.2T Plus"},
        {"Opty Id": "D", "Mobile No": "1000000002", "PPL": "Kestrel", "PL": "Kestrel 1.2 Plus"},  # other model
        {"Opty Id": "E", "Mobile No": "1000000003", "Sales Stage": "07 Closed Lost"},
        {"Opty Id": "F", "Mobile No": "1000000003"},                                     # booked again after a loss
    ])
    c = bk._classify_repeats(bk.clean_bookings(df, "PV", datetime(2026, 10, 6), "pv.xlsx"))
    kind = dict(zip(c["Opty Id"], c["Repeat_Type"]))
    assert kind["A"] == kind["B"] == "Possible duplicate - same variant"
    assert kind["C"] == kind["D"] == "Model change - different model"
    assert kind["F"] == "Rebooked after a lost booking"
    assert c["Dup_Extra"].sum() == 1                                 # only the second copy may be counted twice
    assert c.loc[c["Opty Id"] == "A", "Mobile_Masked"].iloc[0] == "XXXXXX0001"


def test_upsert_counts_and_missing_from_latest():
    day1 = bk.clean_bookings(raw([{"Opty Id": "A"}, {"Opty Id": "B"}]), "PV", datetime(2026, 10, 3), "d1.xlsx")
    master, hist, st = bk.upsert(pd.DataFrame(columns=bk.MASTER_COLS), pd.DataFrame(columns=bk.HISTORY_COLS),
                                 day1, datetime(2026, 10, 3), "PV")
    assert (st["new"], st["updated"], st["unchanged"]) == (2, 0, 0)
    day2 = bk.clean_bookings(raw([{"Opty Id": "A", "Sales Stage": "04 Vehicle Allotted"}]), "PV",
                             datetime(2026, 10, 4), "d2.xlsx")
    master, hist, st = bk.upsert(master, hist, day2, datetime(2026, 10, 4), "PV")
    assert (st["new"], st["updated"], st["unchanged"]) == (0, 1, 0)
    assert master.set_index("Opty Id").loc["B", "Missing_From_Latest"] == "Yes"   # gone from today's file
    assert list(hist["Sales Stage"]) == ["03 Open Yellow / Booked"] * 2 + ["04 Vehicle Allotted"]


def test_dq_checks():
    df = raw([{"Opty Id": "A", "Mobile No": "123456789"},
              {"Opty Id": "B", "Testdrive": "Given"},
              {"Opty Id": "C", "Employee Full Name": "UNSPECIFIED", "Dealer Code": "5999001"}])
    c = bk._classify_repeats(bk.clean_bookings(df, "PV", datetime(2026, 10, 6), "pv.xlsx"))
    checks = set(bk.dq_issues(c)["Check"])
    assert {"Mobile number not 10 digits", "Test drive 'Given' but no date", "Advisor is 'Unspecified'",
            "Dealer Code does not match PV/EV file"} <= checks
