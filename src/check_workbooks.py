r"""
check_workbooks.py  -  can every database workbook be read?  (and put back a good copy if not)
=============================================================================================
    python src\check_workbooks.py              check the workbooks in data\database (changes nothing)
    python src\check_workbooks.py --restore    for every damaged workbook:
                                                 - move it to backup\<name>_DAMAGED_<date_time>.xlsx
                                                 - copy the newest good backup back in its place
                                                 - tell you which day(s) to load again
Stop the bot first (close its window): --restore refuses to run while the bot is running.
"""
import argparse
import glob
import json
import os
import shutil
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import run_lock  # noqa: E402
from handlers import DamagedWorkbookError, verify_xlsx  # noqa: E402

ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
WORKBOOKS = {   # name: what happens when there is no good backup
    "master_bookings.xlsx": "load the history again with backfill.py, then reload the days",
    "open_pipeline.xlsx": "rebuilt from the next 1OPEN files (the trend and exit history is lost)",
    "ssi.xlsx": "rebuilt completely from the next SSI files (they are year to date)",
    "team_map.xlsx": "rebuilt by the next run (the team history starts again)",
    "run_log.xlsx": "rebuilt now from the state and logs folders",
}


def status(path):
    if not os.path.exists(path):
        return "missing", ""
    try:
        verify_xlsx(path)
        return "OK", ""
    except DamagedWorkbookError as exc:
        return "DAMAGED", str(exc)


def good_backups(db, name):
    """Readable backups of one workbook, newest first: [(yyyy-mm-dd, path)]."""
    stem = name[:-5]
    out = []
    for p in glob.glob(os.path.join(db, "backup", f"{stem}_????-??-??.xlsx")):
        if status(p)[0] == "OK":
            out.append((os.path.basename(p)[len(stem) + 1:-5], p))
    return sorted(out, reverse=True)


def main():
    ap = argparse.ArgumentParser(description="Check (and restore) the database workbooks")
    ap.add_argument("--config", default=os.path.join(ROOT, "config", "config.json"))
    ap.add_argument("--restore", action="store_true", help="replace damaged workbooks by their newest good backup")
    a = ap.parse_args()
    cfg = {}
    if os.path.exists(a.config):
        with open(a.config, encoding="utf-8") as fh:
            cfg = json.load(fh)
    paths = cfg.get("paths", {})
    db = os.path.join(ROOT, paths.get("database", os.path.join("data", "database")))
    state = os.path.join(ROOT, paths.get("state", "state"))

    print(f"Workbooks in {db}")
    damaged = []
    for name in WORKBOOKS:
        st, why = status(os.path.join(db, name))
        b = good_backups(db, name)
        print(f"  {st:<8} {name:<22} newest good backup: {b[0][0] if b else 'none':<11} {why}")
        if st == "DAMAGED":
            damaged.append(name)
    leftovers = glob.glob(os.path.join(db, "*.tmp.xlsx")) + glob.glob(os.path.join(db, "*._tmp.xlsx"))
    if not damaged:
        print("All workbooks can be read." + (f" ({len(leftovers)} leftover temp file(s) - removed by --restore.)"
                                               if leftovers else ""))
        if not (a.restore and leftovers):
            return 0
    if not a.restore:
        print(f"\n{len(damaged)} damaged. Close the bot window, Excel and Power BI, then run:\n"
              f"    python src\\check_workbooks.py --restore")
        return 1

    lock = run_lock.RunLock(state, "check_workbooks --restore")
    try:
        lock.acquire()
    except run_lock.AlreadyRunning as held:
        print(f"\nNOT restored: the bot is running ({held}). Close its window first, then run this again.")
        return 2
    try:
        _restore(db, state, cfg, damaged, leftovers)
    except PermissionError as exc:
        print(f"\nNOT finished: a workbook is open in another program ({exc.filename}). "
              f"Close Excel / Power BI / VS Code, then run --restore again.")
        return 2
    finally:
        lock.release()
    return 0


def _restore(db, state, cfg, damaged, leftovers):
    for p in leftovers:
        os.remove(p)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    reload_from = None
    print()
    for name in damaged:
        path = os.path.join(db, name)
        keep = os.path.join(db, "backup", f"{name[:-5]}_DAMAGED_{stamp}.xlsx")
        os.makedirs(os.path.dirname(keep), exist_ok=True)
        shutil.move(path, keep)
        b = good_backups(db, name)
        if b and name != "run_log.xlsx":
            shutil.copy2(b[0][1], path)
            print(f"  {name}: damaged file kept as backup\\{os.path.basename(keep)}; "
                  f"restored the copy of {b[0][0]} (taken before that day's first load)")
            if name != "ssi.xlsx":
                reload_from = min(reload_from or b[0][0], b[0][0])
        else:
            print(f"  {name}: damaged file kept as backup\\{os.path.basename(keep)}; no good backup - "
                  f"{WORKBOOKS[name]}")
    if "run_log.xlsx" in damaged:
        import run_log
        cfg["_root"] = ROOT
        cfg.setdefault("paths", {"database": db, "state": state, "logs": "logs", "raw": "data/raw"})
        res = run_log.update(cfg, datetime.now())
        print(f"  run_log.xlsx rebuilt: {res['days']} day(s)")

    if reload_from:
        days = sorted({os.path.basename(p)[:10] for p in glob.glob(os.path.join(state, "????-??-??.json"))
                       if os.path.basename(p)[:10] >= reload_from} | {reload_from})
        print("\nNow load these day(s) again, in this order (each one takes about a minute):")
        for i, d in enumerate(days):
            last = i == len(days) - 1
            print(f"    python src\\mail_bot.py --date {d} --once --final --reset-state" + ("" if last else " --no-email"))
    print("\nThen check again:  python src\\check_workbooks.py")


if __name__ == "__main__":
    sys.exit(main())
