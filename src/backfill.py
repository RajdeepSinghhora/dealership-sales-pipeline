r"""
backfill.py  -  load OLD months into the master database, safely
=================================================================
Use this for files from previous months (the complete "Bookings - Current Month" file of
each month). It is SAFE to run on a live database:

  * Bookings (Opty Id) already in the database are NEVER changed, so newer daily data is
    never overwritten by an older snapshot.
  * Within one backfill run, if the same Opty Id is in two old months, the NEWER month wins.
  * Files that have no branch column (some old EV files only have 'HO ev Bookings ...') get
    their branch from the advisor's history in the other files (same report group).
  * The daily 'missing from latest file' flags are not touched.
  * Each file is written to Load_Log with Status = BACKFILL (the SLA page can ignore it).

Put the files in data/backfill/ (sub-folders allowed). File names must start with pv or ev
and contain 'Bookings Current Month', e.g. 'pv_Bookings-_Current_Month_complete_april.xlsx'.

Run (from the project folder, with master_bookings.xlsx CLOSED):
    python src/backfill.py --dry-run          # shows what WOULD happen, changes nothing
    python src/backfill.py                    # does it
    python src/backfill.py --folder D:\old    # another folder
"""
import argparse
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import run_lock  # noqa: E402
from handlers import bookings as bk  # noqa: E402

ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
NAME_KEY = "bookingscurrentmonth"


def vehicle_type_of(filename):
    m = re.match(r"^(pv|ev)(?![a-z0-9])", filename, re.I)
    return m.group(1).upper() if m else None


def find_files(folder):
    found = []
    for root, _, files in os.walk(folder):
        for f in files:
            if f.startswith("~$") or not f.lower().endswith((".xlsx", ".xlsm")):
                continue
            if NAME_KEY not in re.sub(r"[^a-z0-9]", "", f.lower()):
                continue
            vt = vehicle_type_of(f)
            if vt:
                found.append((os.path.join(root, f), vt))
    return found


def _norm_div(v):
    return re.sub(r"\s*-\s*", "-", str(v)).strip() if isinstance(v, str) and v.strip() else None


def _branch(div):
    parts = div.split("-") if div else []
    return parts[2].strip() if len(parts) > 2 and parts[2].strip() else None


def _emp(v):
    return re.sub(r"\s+", " ", str(v)).strip().upper() if isinstance(v, str) and v.strip() else None


def build_branch_knowledge(raws, master):
    """Learn (advisor, report group) -> branch, and report group -> branches, from all data that has both."""
    group_branches = defaultdict(Counter)
    emp_group_branch = defaultdict(Counter)
    divisions = Counter()
    for raw in raws:
        if "_Report_Group" not in raw or raw["Division"].isna().all():
            continue
        for emp, grp, div in zip(raw["Employee Full Name"], raw["_Report_Group"], raw["Division"]):
            div = _norm_div(div)
            br, e = _branch(div), _emp(emp)
            if not (br and grp):
                continue
            divisions[div] += 1
            group_branches[grp][br] += 1
            if e:
                emp_group_branch[(e, grp)][br] += 1
    # advisors already in the database (no report group there -> use the group of their branch)
    branch_group = {br: g for g, c in group_branches.items() for br in c}
    for emp, div in zip(master.get("Employee Full Name", []), master.get("Division", [])):
        div = _norm_div(div)
        br, e = _branch(div), _emp(emp)
        if br and e and br in branch_group:
            emp_group_branch[(e, branch_group[br])][br] += 1
        if div:
            divisions[div] += 1
    return group_branches, emp_group_branch, divisions


def recover_division(raw, knowledge):
    """Fill an empty Division from the advisor's history. Returns (raw, n_by_history, n_by_group, n_unresolved)."""
    group_branches, emp_group_branch, divisions = knowledge
    raw = raw.copy()
    by_hist = by_group = unresolved = 0
    suffix = Counter(d.split("-")[-1] for d in divisions if d.count("-") >= 3).most_common(1)
    suffix = suffix[0][0] if suffix else None
    out = []
    for emp, grp, code in zip(raw["Employee Full Name"], raw.get("_Report_Group", [None] * len(raw)), raw["Dealer Code"]):
        br = None
        hist = emp_group_branch.get((_emp(emp), grp))
        if hist:
            br = hist.most_common(1)[0][0]
            by_hist += 1
        elif grp in group_branches and len(group_branches[grp]) == 1:
            br = next(iter(group_branches[grp]))
            by_group += 1
        if br is None:
            unresolved += 1
            out.append(None)
            continue
        code = re.sub(r"\.0+$", "", str(code)).strip() if pd.notna(code) else ""
        known = [d for d in divisions if _branch(d) == br and d.startswith(code + "-")]
        out.append(Counter({d: divisions[d] for d in known}).most_common(1)[0][0] if known
                   else f"{code}-Sales-{br}-{suffix}" if code and suffix else None)
    raw["Division"] = out
    return raw, by_hist, by_group, unresolved


def merge_into(master, history, cleaned, load_date, vehicle_type, protected, batch_ids):
    """Add new Opty Ids; update ones added earlier in THIS run (newer month wins); never touch 'protected'."""
    cleaned = cleaned[cleaned["Opty Id"].notna()].drop_duplicates("Opty Id", keep="last")
    skip = cleaned["Opty Id"].isin(protected)
    upd = cleaned["Opty Id"].isin(batch_ids) & ~skip
    new = cleaned[~skip & ~upd].copy()
    upd_rows = cleaned[upd].copy()

    hist_rows = []
    if len(upd_rows):
        prev = master.set_index("Opty Id")["Sales Stage"].to_dict()
        first = master.set_index("Opty Id")["First_Seen_Date"].to_dict()
        upd_rows["First_Seen_Date"] = upd_rows["Opty Id"].map(first)
        for oid, stg, no in zip(upd_rows["Opty Id"], upd_rows["Sales Stage"], upd_rows["Stage_No"]):
            if str(prev.get(oid)) != str(stg):
                hist_rows.append({"Opty Id": oid, "Vehicle_Type": vehicle_type, "Load_Date": pd.Timestamp(load_date),
                                  "Previous_Stage": prev.get(oid), "Sales Stage": stg, "Stage_No": no})
        master = master[~master["Opty Id"].isin(upd_rows["Opty Id"])]
    for oid, stg, no in zip(new["Opty Id"], new["Sales Stage"], new["Stage_No"]):
        hist_rows.append({"Opty Id": oid, "Vehicle_Type": vehicle_type, "Load_Date": pd.Timestamp(load_date),
                          "Previous_Stage": None, "Sales Stage": stg, "Stage_No": no})
    if hist_rows:
        history = pd.concat([history, pd.DataFrame(hist_rows)], ignore_index=True)
    frames = [f for f in (master, upd_rows, new) if len(f)]
    if frames:
        master = pd.concat(frames, ignore_index=True)
    batch_ids.update(new["Opty Id"])
    return master, history, len(new), len(upd_rows), int(skip.sum())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--folder", default=os.path.join(ROOT, "data", "backfill"))
    ap.add_argument("--master", default=os.path.join(ROOT, "data", "database", bk.MASTER_FILE))
    ap.add_argument("--dry-run", action="store_true", help="show what would happen, change nothing")
    a = ap.parse_args()
    lock = run_lock.RunLock(os.path.join(ROOT, "state"), "backfill")
    try:
        lock.acquire()
    except run_lock.AlreadyRunning as held:
        print(f"NOT started: the bot is running ({held}). Close its window, then run backfill again.")
        return 3

    files = find_files(a.folder)
    if not files:
        print(f"No pv/ev 'Bookings Current Month' files found under {a.folder}")
        return 1

    plan = []
    for path, vt in files:
        try:
            raw, _ = bk.read_bookings_file(path, allow_missing_division=True)
        except bk.SchemaError as exc:
            print(f"SKIP {os.path.relpath(path, a.folder)}: {exc}")
            continue
        last = pd.to_datetime(raw["YF Date"], errors="coerce").max()
        if pd.isna(last):
            print(f"SKIP {os.path.relpath(path, a.folder)}: no YF Date in the file")
            continue
        plan.append([last.normalize(), 0 if vt == "PV" else 1, path, vt, raw])
    plan.sort(key=lambda x: (x[0], x[1]))          # oldest month first, PV before EV

    master, history, loadlog, _ = bk.load_master(a.master)
    protected = set(master["Opty Id"].dropna())
    latest = master["YF Date"].max() if len(master) else None
    print(f"Database: {len(master)} bookings"
          + (f", latest YF date {latest:%d %b %Y}" if latest is not None and pd.notna(latest) else ""))

    knowledge = build_branch_knowledge([p[4] for p in plan], master)
    print("-" * 104)
    print(f"{'':3}{'file':<52}{'month':<10}{'new':>6}{'updated':>9}{'in DB':>7}   branch")
    batch_ids, total_new, total_upd = set(), 0, 0
    for item in plan:
        load_date, _, path, vt, raw = item
        note = ""
        if raw["Division"].isna().all():
            raw, n_h, n_g, n_u = recover_division(raw, knowledge)
            note = f"recovered: {n_h} from advisor history, {n_g} from report group" + (f", {n_u} NOT found" if n_u else "")
        cleaned = bk.clean_bookings(raw, vt, load_date, path)
        months = ", ".join(sorted(cleaned["Booking_Month_Label"].dropna().unique()))
        master, history, n_new, n_upd, n_skip = merge_into(master, history, cleaned, load_date, vt, protected, batch_ids)
        total_new += n_new
        total_upd += n_upd
        rel = os.path.relpath(path, a.folder)
        print(f"{vt:<3}{rel[:51]:<52}{months:<10}{n_new:>6}{n_upd:>9}{n_skip:>7}   {note or 'from file'}")
        loadlog = pd.concat([loadlog, pd.DataFrame([{
            "Load_Date": load_date, "Run_Time": datetime.now().strftime("%H:%M:%S"),
            "Vehicle_Type": vt, "Source_File": os.path.basename(path),
            "Rows_Read": len(raw), "New": n_new, "Updated": n_upd, "Unchanged": n_skip,
            "Missing_From_Latest": 0, "Rows_Skipped": int(raw["Opty Id"].isna().sum()),
            "Extra_Columns_Ignored": 0, "Status": "BACKFILL",
            "Message": f"backfill {rel}: {n_new} added, {n_upd} updated from a newer old month, "
                       f"{n_skip} already in database (left unchanged)" + (f"; branch {note}" if note else "")}])],
            ignore_index=True)

    print("-" * 104)
    print(f"Total: {total_new} new bookings, {total_upd} updated within this run. Database after backfill: {len(master)} bookings.")
    if total_new == 0 and total_upd == 0:
        print("Nothing new to add - database left unchanged.")
        return 0

    # same-mobile checks are recalculated over the whole database (as in the daily run)
    mob = master["Mobile No"].fillna("")
    cnt = master.groupby(mob)["Opty Id"].transform("nunique")
    master["Repeat_Mobile_Count"] = cnt.where(mob != "", 1).astype(int)
    master = bk._classify_repeats(master.reset_index(drop=True))
    dq = bk.dq_issues(master)
    print("Data-quality checks after backfill:")
    for (sev, chk), n in dq.groupby(["Severity", "Check"]).size().items():
        print(f"   {sev:<5} {chk:<58} {n:>5}")

    if a.dry_run:
        print("DRY RUN - nothing was saved. Run again without --dry-run to apply.")
        return 0
    try:
        bk.save_master(a.master, master, history, loadlog, dq)
    except bk.MasterLockedError as exc:
        print(f"NOT SAVED: {exc}")
        return 2
    print(f"Saved {a.master}  (previous version kept in data/database/backup/)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
