r"""
run_demo.py  -  the whole pipeline on fictional data, start to finish, in about two minutes
===========================================================================================
1. generates "Demo Motors" (generate_demo_data.py): six months of CRM history, two weeks of daily reports
2. loads the older months with src/backfill.py
3. replays every report day with the real bot (src/mail_bot.py in folder mode): the files of that
   morning are put in the inbox folder, and the bot runs at the simulated time (_clock.py) - including
   a late start, a file that arrives after the cut-off, a broken file that is sent again and a day
   without SSI files
4. prints what the bot did each day and what is in the five workbooks Power BI reads

    python demo/run_demo.py                                   last 14 days up to today (about 3 minutes)
    python demo/run_demo.py --full                            history from April last year, 1.5x volume,
                                                              31 days of mornings replayed and six months of
                                                              daily-run history (about 20-30 minutes)
    python demo/run_demo.py --end-date 2026-10-12 --days 14   a fixed period (same numbers every time)
    python demo/run_demo.py --workdir C:\temp\demo            write somewhere else (default: this folder)

Output (in the work folder): data/database/*.xlsx, logs/health_*.html (the mails that would be sent),
logs/bot_*.log, state/*.json. Point Power BI at data/database (see docs/power-bi-guide.md).
"""
import argparse
import glob
import json
import os
import shutil
import subprocess
import sys
from datetime import date, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
sys.path.insert(0, HERE)
import generate_demo_data as gen  # noqa: E402

MARKER = os.path.join("data", ".demo_data")
OUTPUT_DIRS = [os.path.join("data", d) for d in ("raw", "database", "test_inbox", "backfill", "demo_source", "reference")] \
    + ["logs", "state"]


def make_config(work):
    """config/config.demo.json with every path made absolute inside the work folder."""
    with open(os.path.join(ROOT, "config", "config.demo.json"), encoding="utf-8") as fh:
        cfg = json.load(fh)
    cfg["paths"] = {k: os.path.join(work, v) for k, v in cfg["paths"].items()}
    cfg["team_map"]["source_file"] = os.path.join(work, cfg["team_map"]["source_file"])
    cfg["ssi"]["id_list"] = os.path.join(work, cfg["ssi"]["id_list"])
    path = os.path.join(work, "data", "demo_source", "config.json")
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, indent=2)
    return path, cfg


def prepare(work, force):
    marker = os.path.join(work, MARKER)
    db = os.path.join(work, "data", "database")
    real = [f for f in glob.glob(os.path.join(db, "*.xlsx"))]
    if real and not os.path.exists(marker) and not force:
        sys.exit(f"STOP: {db} already holds workbooks that were not made by this demo.\n"
                 f"The demo deletes data/, logs/ and state/ in the work folder. Use --workdir <empty folder>, "
                 f"or --force if you are sure.")
    for d in OUTPUT_DIRS:
        shutil.rmtree(os.path.join(work, d), ignore_errors=True)
    os.makedirs(os.path.join(work, "data"), exist_ok=True)
    with open(marker, "w", encoding="utf-8") as fh:
        fh.write("Made by demo/run_demo.py - everything in data/, logs/ and state/ here is fictional demo data.\n")


def run(cmd, log_path):
    with open(log_path, "a", encoding="utf-8") as log:
        log.write("\n$ " + " ".join(cmd) + "\n")
        log.flush()
        res = subprocess.run(cmd, stdout=log, stderr=subprocess.STDOUT, cwd=ROOT,
                             env=dict(os.environ, PYTHONIOENCODING="utf-8"))
    return res.returncode


def outcome(cfg, day, before):
    """Status of the health mail written by this step (GREEN / RED) or 'waiting' when none was written."""
    new = sorted(set(glob.glob(os.path.join(cfg["paths"]["logs"], f"health_{day}_*.html"))) - before)
    if not new:
        return "waiting", None
    name = os.path.basename(new[-1])
    return name.rsplit("_", 1)[1].replace(".html", ""), name


def main():
    ap = argparse.ArgumentParser(description="Run the pipeline on fictional demo data")
    ap.add_argument("--workdir", default=ROOT, help="folder for data/, logs/, state/ (default: the repository)")
    ap.add_argument("--end-date", help="last report day YYYY-MM-DD (default: today)")
    ap.add_argument("--days", type=int, default=14, help="calendar days of daily reports (default 14)")
    ap.add_argument("--seed", type=int, default=7)
    ap.add_argument("--force", action="store_true", help="allow overwriting a data/ folder not made by the demo")
    ap.add_argument("--full", action="store_true",
                    help="history from 1 April of last financial year, 1.5x volume, 31 days of daily reports "
                         "and six months of daily-run (SLA) history")
    a = ap.parse_args()

    work = os.path.abspath(a.workdir)
    end = datetime.strptime(a.end_date, "%Y-%m-%d").date() if a.end_date else date.today()
    prepare(work, a.force)
    t0 = datetime.now()

    print("1/4  Generating Demo Motors ...")
    meta = gen.generate(work, end, a.days, a.seed, full=a.full)
    cfg_path, cfg = make_config(work)
    os.makedirs(cfg["paths"]["logs"], exist_ok=True)
    log_path = os.path.join(cfg["paths"]["logs"], "demo_run.log")

    print("2/4  Loading the older months (src/backfill.py) ...")
    rc = run([sys.executable, os.path.join(ROOT, "src", "backfill.py"), "--folder", os.path.join(work, "data", "backfill"),
              "--master", os.path.join(cfg["paths"]["database"], "master_bookings.xlsx")], log_path)
    if rc != 0:
        sys.exit(f"backfill failed - see {log_path}")
    for item in meta.get("ssi_backfill", []):                   # last financial year's final SSI files
        rc = run([sys.executable, os.path.join(HERE, "_load_file.py"), "ssi", item["group"],
                  os.path.join(work, item["path"]), item["date"], cfg_path], log_path)
        if rc != 0:
            sys.exit(f"loading {item['path']} failed - see {log_path}")
    hist = meta.get("sla_history") or {}
    if hist:                                                   # six months of earlier mornings (SLA history)
        os.makedirs(cfg["paths"]["state"], exist_ok=True)
        for p in glob.glob(os.path.join(work, hist["folder"], "*.json")):
            shutil.copy2(p, cfg["paths"]["state"])
        rc = run([sys.executable, os.path.join(ROOT, "src", "run_log.py"), cfg_path], log_path)
        if rc != 0:
            sys.exit(f"rebuilding the earlier days failed - see {log_path}")
        print(f"     {hist['days']} earlier report days since {datetime.strptime(hist['from'], '%Y-%m-%d'):%d %b %Y} "
              f"rebuilt from the bot's state files (src/run_log.py)")

    print("3/4  Replaying the report days with the bot (src/mail_bot.py) ...")
    inbox = cfg["paths"]["test_inbox"]
    src = os.path.join(work, "data", "demo_source")
    for d in meta["run_days"]:
        day = d["date"]
        state_file = os.path.join(cfg["paths"]["state"], f"{day}.json")
        os.makedirs(cfg["paths"]["state"], exist_ok=True)
        with open(state_file, "w", encoding="utf-8") as fh:        # what Task Scheduler would have started
            json.dump({"done": False, "sent_sig": "", "results": {}, "first_start": d["bot_started"]}, fh)
        shown = []
        for step in d["steps"]:
            shutil.rmtree(inbox, ignore_errors=True)
            os.makedirs(inbox)
            for rel in step["files"]:
                shutil.copy2(os.path.join(src, rel), os.path.join(inbox, os.path.basename(rel)))
            before = set(glob.glob(os.path.join(cfg["paths"]["logs"], f"health_{day}_*.html")))
            rc = run([sys.executable, os.path.join(HERE, "_clock.py"), f"{day} {step['at']}", "--config", cfg_path,
                      "--date", day, "--once", "--no-email"], log_path)
            if rc != 0:
                sys.exit(f"the bot stopped with an error on {day} {step['at']} - see {log_path}")
            status, _ = outcome(cfg, day, before)
            shown.append(f"{step['at']} {status}" + (f" ({step['note']})" if step.get("note") else ""))
        print(f"     {datetime.strptime(day, '%Y-%m-%d'):%a %d %b}  " + "  ->  ".join(shown))
    shutil.rmtree(inbox, ignore_errors=True)
    os.makedirs(inbox)

    print("4/4  Result")
    summarize(cfg["paths"]["database"])
    print(f"\nDone in {(datetime.now() - t0).seconds} s. Workbooks: {cfg['paths']['database']}")
    print(f"Health mails (open in a browser): {cfg['paths']['logs']}{os.sep}health_*.html")
    return 0


def summarize(db):
    import pandas as pd
    for wb in ("master_bookings.xlsx", "open_pipeline.xlsx", "ssi.xlsx", "team_map.xlsx", "run_log.xlsx"):
        path = os.path.join(db, wb)
        if not os.path.exists(path):
            print(f"     {wb:<22} MISSING")
            continue
        with pd.ExcelFile(path) as x:
            parts = [f"{s} {len(x.parse(s)):,}" for s in x.sheet_names]
        print(f"     {wb:<22} " + " | ".join(parts))
    runs = pd.read_excel(os.path.join(db, "run_log.xlsx"), "Daily_Runs")
    met = runs["SLA_Met"].value_counts()
    print(f"     Daily runs: {len(runs)} days, {int((runs['Status'] == 'GREEN').sum())} GREEN, "
          f"data ready by {runs['SLA_Target'].iloc[-1]} on {int(met.get('Yes', 0))} of "
          f"{int(met.get('Yes', 0) + met.get('No', 0))} days")


if __name__ == "__main__":
    sys.exit(main())
