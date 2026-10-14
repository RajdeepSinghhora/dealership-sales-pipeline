"""Runs the real bot on a few days of fictional data - the same path a reviewer takes with demo/run_demo.py."""
import subprocess
import sys

import pandas as pd

from conftest import ROOT


def test_demo_pipeline(tmp_path):
    res = subprocess.run([sys.executable, f"{ROOT}/demo/run_demo.py", "--workdir", str(tmp_path),
                          "--end-date", "2026-10-12", "--days", "4"], capture_output=True, text=True, timeout=900)
    assert res.returncode == 0, res.stdout[-3000:] + res.stderr[-3000:]
    db = tmp_path / "data" / "database"
    for wb in ("master_bookings.xlsx", "open_pipeline.xlsx", "ssi.xlsx", "team_map.xlsx", "run_log.xlsx"):
        assert (db / wb).exists(), wb
    runs = pd.read_excel(db / "run_log.xlsx", "Daily_Runs")
    assert len(runs) == 3 and set(runs["Status"]) == {"GREEN"}
    assert list(runs["Bot_Started"]) == ["08:30", "08:30", "08:30"]
    bookings = pd.read_excel(db / "master_bookings.xlsx", "Bookings")
    assert len(bookings) > 500 and set(bookings["Vehicle_Type"]) == {"PV", "EV"}
    assert bookings["Mobile No"].astype(str).str.startswith("1").all()   # demo numbers are never real numbers
    surveys = pd.read_excel(db / "ssi.xlsx", "SSI_Surveys")
    assert {"Actual seller (GM register)", "CA ID list"} <= set(surveys["CA_Source"])
    team_keys = set(pd.read_excel(db / "team_map.xlsx", "Team_Map")["CA_Key"].str.upper())
    assert set(surveys["CA_Key"].str.upper()) <= team_keys          # every survey has a team row (no (Blank))
    issues = pd.read_excel(db / "run_log.xlsx", "Data_Issues")
    assert {"Bookings", "Open pipeline", "SSI", "Team list"} <= set(issues["Source"])
    assert len(list((tmp_path / "logs").glob("health_*_GREEN.html"))) == 3


def test_full_demo_plan(tmp_path):
    """--full: two financial years, last year's final SSI files and a month of mixed mornings (data only, no bot run)."""
    import generate_demo_data as gen
    from datetime import date
    meta = gen.generate(str(tmp_path), date(2026, 10, 14), seed=7, full=True, quiet=True)
    assert meta["history_from"] == "2025-04-01"
    assert [s["group"] for s in meta["ssi_backfill"]] == ["PV", "EV"]
    assert len(meta["run_days"]) >= 26
    kinds = [d["scenario"] for d in meta["run_days"]]
    assert "normal" in kinds and {k for _l, k in gen.FULL_SCENARIOS} <= set(kinds)   # every kind of bad morning
    assert kinds.count("missing_file") >= 2                            # two days that close RED
    notes = {s.get("note") for d in meta["run_days"] if d["scenario"] == "missing_file" for s in d["steps"]}
    assert notes >= {"cut-off: EV 1OPEN never sent", "cut-off: PV bookings never sent"}
    assert meta["run_days"][-1]["scenario"] == "normal"
    assert meta["bookings"] > 2000 and meta["surveys"] > 400


def test_full_demo_sla_history(tmp_path):
    """--full: six months of earlier mornings as the bot's state files; src/run_log.py rebuilds them into Daily_Runs."""
    import json
    import shutil
    from datetime import date

    import generate_demo_data as gen
    import run_log
    meta = gen.generate(str(tmp_path), date(2026, 10, 14), seed=7, full=True, quiet=True)
    hist = meta["sla_history"]
    assert hist["from"] == "2026-04-01" and hist["days"] >= 130
    with open(f"{ROOT}/config/config.demo.json", encoding="utf-8") as fh:
        cfg = json.load(fh)
    cfg["_root"] = str(tmp_path)
    state = tmp_path / "state"
    state.mkdir()
    for p in (tmp_path / hist["folder"]).glob("*.json"):
        shutil.copy2(p, state)
    assert run_log.update(cfg)["seeded"] == hist["days"]
    runs = pd.read_excel(tmp_path / "data" / "database" / "run_log.xlsx", "Daily_Runs")
    months = pd.to_datetime(runs["Run_Date"]).dt.to_period("M")
    assert months.nunique() >= 6 and set(runs["Status"]) == {"GREEN", "RED"}
    assert set(runs["SLA_Met"]) == {"Yes", "No"} and 0.7 < (runs["SLA_Met"] == "Yes").mean() < 0.97
    assert pd.to_datetime(runs["Run_Date"]).dt.weekday.max() < 6                  # no reports on Sundays
    files = pd.read_excel(tmp_path / "data" / "database" / "run_log.xlsx", "File_Loads")
    ok = files.loc[files["Status"] == "OK", "Rows_Read"]
    assert ok.notna().all() and (ok > 0).mean() > 0.95        # a new month can start with no EV booking yet
