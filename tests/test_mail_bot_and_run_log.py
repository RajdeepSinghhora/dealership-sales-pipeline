import mail_bot as mb
import run_log as rl

CFG = {"groups": {
    "PV": {"file_prefix": "PV", "files": [{"name": "pv Bookings- Current Month", "handler": "bookings"},
                                          {"name": "pv 1OPEN GF-PF-YF-ALLT", "handler": "open_pipeline"},
                                          {"name": "pv SSI Report (YTD)", "handler": "ssi", "required": False}]},
    "EV": {"file_prefix": "EV", "files": [{"name": "ev Bookings- Current Month", "handler": "bookings"}]}},
    "schedule": {"sla_ready_by": "13:00"}}


def test_files_are_recognised_by_name():
    assert mb.group_of(CFG, "pv 1OPEN GF-PF-YF-ALLT_2.xlsx") == "PV"
    assert mb.group_of(CFG, "EV - SSI Report (YTD).xlsx") == "EV"
    assert mb.group_of(CFG, "PVR report.xlsx") is None                # 'PVR' is not 'PV'
    assert mb.rule_for(CFG["groups"]["PV"], "pv 1OPEN GF-PF-YF-ALLT_2.xlsx")["handler"] == "open_pipeline"
    assert mb.rule_for(CFG["groups"]["PV"], "pv Bookings- Current Month (1).xlsx")["handler"] == "bookings"


def test_file_grid_rows_and_day_counts():
    report = [{"name": "PV", "checked": True, "received": [
        {"name": "pv Bookings- Current Month.xlsx", "status": "OK", "arrived": "09:10", "rows": 40, "new": 3,
         "updated": 1, "unchanged": 36}]},
        {"name": "EV", "checked": True, "received": []}]
    rows = rl.file_rows(CFG, "2026-10-12", report)
    st = {r["File_Label"]: r["Status"] for r in rows}
    assert st == {"PV Bookings": "OK", "PV Open pipeline": "Missing", "PV SSI": "Not sent", "EV Bookings": "Missing"}
    day = rl.daily_row(CFG, "2026-10-12", "RED", "2 file(s) missing", rows, True, None, False, "08:30", "Live")
    assert (day["Files_Expected"], day["Files_OK"], day["Files_Missing"]) == (3, 1, 2)
    assert day["SLA_Met"] == "No" and day["Weekday"] == "Mon"


def test_day_that_ends_red_at_the_cut_off_counts_as_sla_missed(tmp_path, monkeypatch):
    """Regression: at 13:00 with a file still missing the SLA was 'pending' and the day was left out of SLA %."""
    import datetime as _dt
    import json
    import logging

    import pandas as pd

    class At1300(_dt.datetime):
        @classmethod
        def now(cls, tz=None):
            return _dt.datetime(2026, 10, 12, 13, 0)

    monkeypatch.setattr(mb, "datetime", At1300)
    with open(f"{mb.ROOT}/config/config.demo.json", encoding="utf-8") as fh:
        cfg = json.load(fh)
    cfg["_root"] = str(tmp_path)
    (tmp_path / "data" / "test_inbox").mkdir(parents=True)              # nothing arrived
    state = {"done": False, "sent_sig": "", "results": {}}
    mb.run_cycle(cfg, _dt.datetime(2026, 10, 12), True, state, logging.getLogger("t"), send=False)
    runs = pd.read_excel(tmp_path / "data" / "database" / "run_log.xlsx", "Daily_Runs")
    assert runs.loc[0, "Status"] == "RED" and runs.loc[0, "SLA_Met"] == "No"
    html = next((tmp_path / "logs").glob("health_2026-10-12_*_RED.html")).read_text(encoding="utf-8")
    assert "SLA MISSED" in html and "4 file(s) missing" in html
