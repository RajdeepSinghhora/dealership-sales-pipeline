r"""
run_log.py  -  the daily record behind the Power BI "Data Operations" page
=========================================================================
Writes data/database/run_log.xlsx with four sheets:

  Daily_Runs      one row per day: GREEN / RED, when the data was ready, SLA met?, files OK / missing / failed
  File_Loads      one row per day per file: arrived at, status, rows read / new / updated (+ cell text and colour
                  for the arrival grid)
  Data_Issues     every open data-quality problem from all workbooks, with WHERE to fix it
  Data_Freshness  one row per data source: records, latest data date, last loaded, days since load

The bot updates it at the end of every check (mail_bot.py calls record_day). If run_log.xlsx is open
in Excel it is skipped with a warning and written again at the next check - the bot itself never stops.

Run it by hand any time - the first time it builds the history of earlier days from the state/ folder:
    python src\run_log.py
"""
import glob
import json
import os
import re
import sys
from datetime import datetime

import pandas as pd
from handlers import temp_name, verify_xlsx

ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
LOG_FILE = "run_log.xlsx"

DAILY_COLS = ["Run_Date", "Weekday", "Status", "Status_Color", "Headline", "Bot_Started", "Ready_Time", "Ready_Hour",
              "SLA_Target", "SLA_Target_Hour", "SLA_Met", "Gmail_Reached", "Files_Expected", "Files_OK",
              "Files_Missing", "Files_Failed", "Rows_Read", "Recorded", "Last_Update"]
FILE_COLS = ["Run_Date", "Group", "File", "File_Label", "File_Order", "Required", "Arrived", "Status", "Status_Order",
             "Cell_Text", "Cell_Color", "Rows_Read", "New", "Updated", "Unchanged", "Note"]
ISSUE_COLS = ["Source", "Severity", "Severity_Order", "Check", "Record_Id", "Vehicle_Type", "Branch", "Person",
              "Detail", "Fix_Where"]
FRESH_COLS = ["Source", "Source_Order", "Workbook", "Records", "Latest_Data", "Last_Loaded", "Days_Since_Load",
              "Checked_On"]

FRIENDLY = {"bookings": "Bookings", "open_pipeline": "Open pipeline", "ssi": "SSI"}

# status of one file on one day -> (sort order, light background colour for the grid)
FILE_STATUS = {"OK": (1, "#E2F0D9"), "Not sent": (2, "#F2F2F2"), "Saved only": (3, "#DDEBF7"),
               "Not checked": (4, "#FFF2CC"), "Failed": (5, "#F4B6B6"), "Missing": (6, "#F8CBAD")}
DAY_STATUS_COLOR = {"GREEN": "#E2F0D9", "RED": "#F8CBAD", "Incomplete": "#FFF2CC"}

# open-pipeline checks that are SALES follow-ups (shown on the Sales Pipeline page), not data problems
FOLLOW_UP_CHECKS = ("Enquiry open >", "Booked customer waiting >")
SSI_FIX = {"Seller not found": "CA_ID_List.xlsx", "Shared ID - seller guessed": "GM register (Actual CA)",
           "Survey not completed": "Nothing - waiting for the OEM"}
TEAM_WARN = ("In Siebel, not in team list", "Conflict", "Skipped - no TL", "Team Type missing")


# --------------------------------------------------------------------------- #
# small helpers
# --------------------------------------------------------------------------- #
def _stem(name):
    """'PV current month (1).xlsx' -> 'pvcurrentmonth'  (same rule as mail_bot.stem_key)."""
    s = re.sub(r"[^a-z0-9]", "", os.path.splitext(str(name))[0].lower())
    return re.sub(r"\d+$", "", s)


def _hour(hhmm):
    if not hhmm or not re.fullmatch(r"\d{1,2}:\d{2}", str(hhmm)):
        return None
    h, m = str(hhmm).split(":")
    return round(int(h) + int(m) / 60, 2)


def _num(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


def _day(d):
    return pd.Timestamp(d).normalize()


def _sheet(path, name, cols):
    if not os.path.exists(path):
        return pd.DataFrame(columns=cols)
    try:
        with pd.ExcelFile(path) as xl:
            if name not in xl.sheet_names:
                return pd.DataFrame(columns=cols)
            return xl.parse(name)
    except Exception:
        return pd.DataFrame(columns=cols)


# --------------------------------------------------------------------------- #
# one day -> rows
# --------------------------------------------------------------------------- #
def _from_result(r):
    st = {"OK": "OK", "FAILED": "Failed"}.get(r.get("status"), "Saved only")
    order, colour = FILE_STATUS[st]
    text = r.get("arrived") if st == "OK" else ("Failed" if st == "Failed" else "Saved")
    return {"Arrived": r.get("arrived"), "Status": st, "Status_Order": order, "Cell_Text": text, "Cell_Color": colour,
            "Rows_Read": _num(r.get("rows")), "New": _num(r.get("new")), "Updated": _num(r.get("updated")),
            "Unchanged": _num(r.get("unchanged")), "Note": r.get("note") or None}


def _no_file(status):
    order, colour = FILE_STATUS[status]
    text = {"Missing": "Missing", "Not sent": "–", "Not checked": "?"}[status]
    return {"Arrived": None, "Status": status, "Status_Order": order, "Cell_Text": text, "Cell_Color": colour,
            "Rows_Read": None, "New": None, "Updated": None, "Unchanged": None,
            "Note": {"Missing": "Required file not received", "Not sent": "Optional file - not sent this day",
                     "Not checked": "Gmail could not be reached"}[status]}


def file_rows(cfg, day, groups_report, green_day=False):
    """One row per expected file (and per extra file received) for this day.
    green_day: an earlier day that ended GREEN - a required file with no trace was simply not expected then."""
    rows = []
    for gi, (g, gc) in enumerate(cfg["groups"].items()):
        grp = next((x for x in groups_report if x.get("name") == g), None)
        received = list(grp.get("received", [])) if grp else []
        checked = bool(grp.get("checked", True)) if grp else False
        used = set()
        for fi, rule in enumerate(gc.get("files", [])):
            hits = [r for r in received if _stem(r.get("name")) == _stem(rule["name"])]
            used.update(id(r) for r in hits)
            req = rule.get("required", True)
            base = {"Run_Date": _day(day), "Group": g, "File": rule["name"],
                    "File_Label": f"{g} {FRIENDLY.get(rule.get('handler'), rule['name'])}",
                    "File_Order": gi * 10 + fi + 1, "Required": "Yes" if req else "No"}
            if hits:
                rows.append({**base, **_from_result(hits[-1])})       # files run oldest first -> last one counts
            elif not checked:
                rows.append({**base, **_no_file("Not checked")})
            elif req and green_day:
                rows.append({**base, **_no_file("Not sent"), "Note": "Not expected by the bot on this day"})
            else:
                rows.append({**base, **_no_file("Missing" if req else "Not sent")})
        for r in received:
            if id(r) not in used:
                rows.append({"Run_Date": _day(day), "Group": g, "File": r.get("name"),
                             "File_Label": f"{g} {os.path.splitext(str(r.get('name')))[0]}",
                             "File_Order": gi * 10 + 9, "Required": "No", **_from_result(r)})
    return rows


def daily_row(cfg, day, status, headline, frows, gmail_ok, ready, met, started, recorded):
    # not expected that day -> not counted: "Not sent", or on a GREEN day "Saved only" (no rule for it yet)
    skip = ("Not sent", "Saved only") if status == "GREEN" else ("Not sent",)
    req = [f for f in frows if f["Required"] == "Yes" and f["Status"] not in skip]
    target = cfg.get("schedule", {}).get("sla_ready_by")
    return {"Run_Date": _day(day), "Weekday": _day(day).strftime("%a"), "Status": status,
            "Status_Color": DAY_STATUS_COLOR.get(status, "#F2F2F2"), "Headline": headline,
            "Bot_Started": started, "Ready_Time": ready, "Ready_Hour": _hour(ready),
            "SLA_Target": target, "SLA_Target_Hour": _hour(target),
            "SLA_Met": {True: "Yes", False: "No"}.get(met, None),
            "Gmail_Reached": "Yes" if gmail_ok else "No",
            "Files_Expected": len(req), "Files_OK": sum(f["Status"] == "OK" for f in req),
            "Files_Missing": sum(f["Status"] == "Missing" for f in frows),
            "Files_Failed": sum(f["Status"] == "Failed" for f in frows),
            "Rows_Read": int(sum(f["Rows_Read"] or 0 for f in frows if f["Status"] == "OK")),
            "Recorded": recorded, "Last_Update": datetime.now().replace(microsecond=0)}


def _from_report(cfg, day, report, state):
    frows = file_rows(cfg, day, report.get("groups", []))
    groups = report.get("groups", [])
    gmail_ok = bool(groups) and all(g.get("checked", True) for g in groups)
    sla = report.get("sla") or {}
    ready = sla.get("ready_at") if report["status"] == "GREEN" else None
    met = sla.get("met")
    if report["status"] == "GREEN" and met is None and ready and sla.get("target"):
        met = ready <= sla["target"]
    d = daily_row(cfg, day, report["status"], report.get("headline"), frows, gmail_ok, ready, met,
                  state.get("first_start"), "Live")
    return d, frows


def _health_files(cfg, day):
    """[(HH:MM, GREEN|RED), ...] of the health mails saved in logs/ for that day, earliest first."""
    folder = os.path.join(cfg["_root"], cfg["paths"].get("logs", "logs"))
    out = []
    for p in glob.glob(os.path.join(folder, f"health_{day:%Y-%m-%d}_*.html")):
        m = re.search(r"_(\d{2})(\d{2})\d{2}_(GREEN|RED)\.html$", p)
        if m:
            out.append((f"{m.group(1)}:{m.group(2)}", m.group(3)))
    return sorted(out)


def _from_state_file(cfg, path):
    """Rebuild one earlier day from state/<yyyy-mm-dd>.json (used once, to fill the history)."""
    day = datetime.strptime(os.path.basename(path)[:10], "%Y-%m-%d")
    with open(path, encoding="utf-8") as fh:
        st = json.load(fh)
    sig = st.get("sent_sig") or ""
    rec = {g: [] for g in cfg["groups"]}
    for key, r in (st.get("results") or {}).items():
        g = key.split("|", 1)[0]
        rec.setdefault(g, []).append(r)
    for g, name in re.findall(r"(?:^|\|)([A-Z]{2}): (.+?) failed to load", sig):
        if not any(_stem(r.get("name")) == _stem(name) and r.get("status") == "OK" for r in rec.get(g, [])):
            rec.setdefault(g, []).append({"name": name, "status": "FAILED", "arrived": None})
    # files that were downloaded that day but never loaded (e.g. before their handler existed) -> "Saved only"
    raw_day = os.path.join(cfg["_root"], cfg["paths"].get("raw", "data/raw"), day.strftime("%Y-%m-%d"))
    for g in cfg["groups"]:
        folder = os.path.join(raw_day, g)
        if not os.path.isdir(folder):
            continue
        for fn in sorted(os.listdir(folder)):
            if fn.startswith("~$") or not fn.lower().endswith((".xlsx", ".xlsm", ".xls")):
                continue
            if not any(_stem(r.get("name")) == _stem(fn) for r in rec.get(g, [])):
                rec.setdefault(g, []).append({"name": fn, "status": "Received", "arrived": None,
                                              "note": "Received but not loaded by the bot at that time"})
    gmail_ok = "Cannot read the mailbox" not in sig or any(rec.values())

    saved_key, saved_file = st.get("saved_key") or "", st.get("saved_file") or ""
    status = saved_key.split("|", 1)[0] if saved_key else ("RED" if sig else ("GREEN" if st.get("done") else "Incomplete"))
    m = re.search(r"_(\d{2})(\d{2})\d{2}_GREEN\.html$", saved_file)
    ready = f"{m.group(1)}:{m.group(2)}" if (status == "GREEN" and m) else None
    greens = [t for t, stt in _health_files(cfg, day) if stt == "GREEN"]
    if greens:                                    # the first GREEN mail of the day = when the data was ready
        status, ready = "GREEN", greens[0]
    if status not in ("GREEN", "RED"):
        status = "Incomplete"
    groups = [{"name": g, "checked": gmail_ok, "received": rec.get(g, [])} for g in cfg["groups"]]
    frows = file_rows(cfg, day, groups, green_day=(status == "GREEN"))
    target = cfg.get("schedule", {}).get("sla_ready_by")
    finished = st.get("done") or _day(day) < _day(datetime.now())         # an earlier day is over anyway
    met = (ready <= target) if (ready and target) else (False if (status != "GREEN" and finished) else None)
    if status == "GREEN":
        headline = "All files received and loaded"
    elif sig:
        headline = sig.split("|")[0][:200]
    else:
        headline = "Day not finished"
    return day, daily_row(cfg, day, status, headline, frows, gmail_ok, ready, met, st.get("first_start"), "Rebuilt"), frows


# --------------------------------------------------------------------------- #
# issues + freshness (rebuilt every time from the other workbooks)
# --------------------------------------------------------------------------- #
def collect_issues(db):
    parts = []
    dq = _sheet(os.path.join(db, "master_bookings.xlsx"), "DQ_Issues", [])
    if len(dq):
        parts.append(pd.DataFrame({"Source": "Bookings", "Severity": dq.get("Severity"), "Check": dq.get("Check"),
                                   "Record_Id": dq.get("Opty Id"), "Vehicle_Type": dq.get("Vehicle_Type"),
                                   "Branch": dq.get("Branch"), "Person": dq.get("Employee Full Name"),
                                   "Detail": dq.get("Detail"), "Fix_Where": "Siebel"}))
    op = _sheet(os.path.join(db, "open_pipeline.xlsx"), "Open_Issues", [])
    if len(op):
        op = op[~op["Check"].astype(str).str.startswith(FOLLOW_UP_CHECKS)]
        parts.append(pd.DataFrame({"Source": "Open pipeline", "Severity": op.get("Severity"), "Check": op.get("Check"),
                                   "Record_Id": op.get("Opty Id"), "Vehicle_Type": op.get("Vehicle_Type"),
                                   "Branch": op.get("Branch"), "Person": op.get("Employee Full Name"),
                                   "Detail": op.get("Detail"), "Fix_Where": "Siebel"}))
    si = _sheet(os.path.join(db, "ssi.xlsx"), "SSI_Issues", [])
    if len(si):
        parts.append(pd.DataFrame({"Source": "SSI", "Severity": si["Issue"].map(
                                       lambda i: "INFO" if i == "Survey not completed" else "WARN"),
                                   "Check": si["Issue"], "Record_Id": si.get("CSR#"), "Vehicle_Type": si.get("Vehicle_Type"),
                                   "Branch": None, "Person": si.get("CA").fillna(si.get("CA_ID")),
                                   "Detail": si.get("Detail"), "Fix_Where": si["Issue"].map(SSI_FIX).fillna("Check")}))
    mi = _sheet(os.path.join(db, "team_map.xlsx"), "Map_Issues", [])
    if len(mi):
        det = mi.get("Detail").astype(str)
        parts.append(pd.DataFrame({"Source": "Team list", "Severity": mi["Issue"].astype(str).map(
                                       lambda i: "WARN" if i.startswith(TEAM_WARN) else "INFO"),
                                   "Check": mi["Issue"], "Record_Id": None,
                                   "Vehicle_Type": det.str[:2].where(det.str[:2].isin(["PV", "EV"])),
                                   "Branch": None, "Person": mi.get("CA"), "Detail": mi.get("Detail"),
                                   "Fix_Where": "Team_Mapping.xlsx"}))
    if not parts:
        return pd.DataFrame(columns=ISSUE_COLS)
    out = pd.concat(parts, ignore_index=True)
    out["Severity"] = out["Severity"].fillna("WARN")
    out["Severity_Order"] = out["Severity"].map({"WARN": 1, "INFO": 2}).fillna(3).astype(int)
    for c in ("Record_Id", "Person", "Branch", "Vehicle_Type", "Detail", "Check"):
        out[c] = out[c].where(out[c].notna(), None).map(lambda v: None if v is None else str(v))
    return out.reindex(columns=ISSUE_COLS).sort_values(["Severity_Order", "Source", "Check"]).reset_index(drop=True)


def _maxdate(s):
    s = pd.to_datetime(s, errors="coerce").dropna()
    return s.max().normalize() if len(s) else None


def collect_freshness(db, day):
    rows = []
    mb = os.path.join(db, "master_bookings.xlsx")
    b = _sheet(mb, "Bookings", [])
    ll = _sheet(mb, "Load_Log", [])
    if len(ll) and "Status" in ll:
        ll = ll[ll["Status"].astype(str).str.upper() != "BACKFILL"]
    rows.append(("Bookings", 1, "master_bookings.xlsx", len(b), _maxdate(b.get("YF Date", pd.Series(dtype=object))),
                 _maxdate(ll.get("Load_Date", pd.Series(dtype=object)))))
    on = _sheet(os.path.join(db, "open_pipeline.xlsx"), "Open_Now", [])
    snap = _maxdate(on.get("Snapshot_Date", pd.Series(dtype=object)))
    rows.append(("Open pipeline", 2, "open_pipeline.xlsx", len(on), snap, snap))
    sv = _sheet(os.path.join(db, "ssi.xlsx"), "SSI_Surveys", [])
    done = sv[sv["Is_Done"] == 1] if "Is_Done" in sv else sv
    rows.append(("SSI surveys", 3, "ssi.xlsx", len(done), _maxdate(done.get("Assessment Date", pd.Series(dtype=object))),
                 _maxdate(sv.get("Load_Date", pd.Series(dtype=object)))))
    tm = _sheet(os.path.join(db, "team_map.xlsx"), "Team_Map", [])
    in_list = tm[tm["TL"] != "Not in team list"] if "TL" in tm else tm
    rows.append(("Team list", 4, "team_map.xlsx", len(in_list), None,
                 _maxdate(tm.get("Loaded_On", pd.Series(dtype=object)))))
    d0 = _day(day)
    out = [{"Source": s, "Source_Order": o, "Workbook": w, "Records": int(n), "Latest_Data": latest,
            "Last_Loaded": last, "Days_Since_Load": (d0 - last).days if last is not None else None, "Checked_On": d0}
           for s, o, w, n, latest, last in rows]
    return pd.DataFrame(out, columns=FRESH_COLS)


# --------------------------------------------------------------------------- #
# read / write
# --------------------------------------------------------------------------- #
def _read(path):
    daily = _sheet(path, "Daily_Runs", DAILY_COLS).reindex(columns=DAILY_COLS)
    files = _sheet(path, "File_Loads", FILE_COLS).reindex(columns=FILE_COLS)
    for df in (daily, files):
        if len(df):
            df["Run_Date"] = pd.to_datetime(df["Run_Date"]).dt.normalize()
    return daily, files


def _upsert(daily, files, drow, frows):
    d = drow["Run_Date"]
    daily = daily[daily["Run_Date"] != d] if len(daily) else daily
    files = files[files["Run_Date"] != d] if len(files) else files
    daily = pd.concat([x for x in (daily, pd.DataFrame([drow], columns=DAILY_COLS)) if len(x)], ignore_index=True)
    fnew = pd.DataFrame(frows, columns=FILE_COLS)
    files = pd.concat([x for x in (files, fnew) if len(x)], ignore_index=True) if len(fnew) or len(files) else fnew
    return daily, files


def _seed(cfg, daily, files, skip):
    """(Re)build every day that the bot did not record live, from the state/ folder."""
    have = set(daily.loc[daily["Recorded"] == "Live", "Run_Date"]) if len(daily) else set()
    n = 0
    for p in sorted(glob.glob(os.path.join(cfg["_root"], cfg["paths"]["state"], "????-??-??.json"))):
        try:
            d = _day(datetime.strptime(os.path.basename(p)[:10], "%Y-%m-%d"))
            if d in have or d in skip:
                continue
            _, drow, frows = _from_state_file(cfg, p)
            daily, files = _upsert(daily, files, drow, frows)
            n += 1
        except Exception:
            continue
    return daily, files, n


def _write(path, sheets):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = temp_name(path)
    with pd.ExcelWriter(tmp, engine="openpyxl", datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as w:
        for name, df in sheets:
            df.to_excel(w, sheet_name=name, index=False)
            ws = w.sheets[name]
            ws.freeze_panes = "A2"
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 10), 60)
    verify_xlsx(tmp)
    try:
        os.replace(tmp, path)
    except PermissionError:
        os.remove(tmp)
        raise PermissionError(f"{os.path.basename(path)} is open in Excel - close it; it is written again at the next check")


def _keep_first_ready(cfg, daily, drow):
    """A rerun later the same day must not move the ready time; a rerun on a LATER day means the SLA was missed."""
    if drow["Status"] != "GREEN":
        return
    prev = daily[daily["Run_Date"] == drow["Run_Date"]] if len(daily) else daily
    if len(prev):
        p = prev.iloc[-1]
        pt = p["Ready_Time"] if isinstance(p["Ready_Time"], str) else None
        if p["Status"] == "GREEN" and pt and re.fullmatch(r"\d{1,2}:\d{2}", pt):
            if not drow["Ready_Time"] or pt < drow["Ready_Time"]:
                drow.update(Ready_Time=pt, Ready_Hour=_hour(pt), SLA_Met=p["SLA_Met"])
            return
    if drow["Run_Date"] < _day(datetime.now()):
        now = datetime.now()
        drow.update(Ready_Time=f"{now:%d %b %H:%M}", Ready_Hour=None, SLA_Met="No",
                    Headline=f"{drow['Headline']} (loaded late, on {now:%d %b})")


def update(cfg, day=None, report=None, state=None, log=None):
    """Record one day (when a report is given), fill missing history, refresh issues + freshness, save."""
    db = os.path.join(cfg["_root"], cfg["paths"]["database"])
    path = os.path.join(db, LOG_FILE)
    daily, files = _read(path)
    skip = set()
    if report is not None:
        drow, frows = _from_report(cfg, day, report, state or {})
        _keep_first_ready(cfg, daily, drow)
        daily, files = _upsert(daily, files, drow, frows)
        skip.add(drow["Run_Date"])
    daily, files, seeded = _seed(cfg, daily, files, skip)
    daily = daily.sort_values("Run_Date").reset_index(drop=True)
    files = files.sort_values(["Run_Date", "File_Order"]).reset_index(drop=True) if len(files) else files
    issues = collect_issues(db)
    fresh = collect_freshness(db, day or datetime.now())
    _write(path, [("Daily_Runs", daily.reindex(columns=DAILY_COLS)), ("File_Loads", files.reindex(columns=FILE_COLS)),
                  ("Data_Issues", issues), ("Data_Freshness", fresh)])
    if log:
        log.info("run_log.xlsx updated: %d day(s)%s, %d data issue(s)", len(daily),
                 f" ({seeded} rebuilt from state/)" if seeded else "", len(issues))
    return {"days": len(daily), "seeded": seeded, "issues": len(issues), "path": path}


def record_day(cfg, day, report, state, log=None):
    return update(cfg, day, report, state, log)


def main():
    cfg_path = os.environ.get("BOOKINGS_CONFIG") or os.path.join(ROOT, "config", "config.json")
    if len(sys.argv) > 1 and sys.argv[1].endswith(".json"):
        cfg_path = sys.argv[1]
    with open(cfg_path, encoding="utf-8") as fh:
        cfg = json.load(fh)
    cfg["_root"] = ROOT
    try:
        res = update(cfg, datetime.now())
    except PermissionError as exc:
        print("NOT saved:", exc)
        return 1
    print(f"Saved {res['path']}")
    print(f"  Daily_Runs: {res['days']} day(s)" + (f"  ({res['seeded']} rebuilt from the state folder)" if res["seeded"] else ""))
    print(f"  Data_Issues: {res['issues']} open issue(s)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
