"""
mail_bot.py  -  the daily mail bot
==================================
* From the start time it checks the mailbox every N minutes.
* It looks ONLY at mails sent by the addresses listed in config "senders"
  (the subject line is ignored - it changes every day).
* Every Excel attachment is saved to data/raw/<date>/<PV|EV|_other>/ .
  PV or EV is decided from the FILE NAME ("PV current month.xlsx" -> PV).
* When every required file of every group has arrived, each file is passed to its handler
  (src/handlers/<name>.py), PV first, then EV. Then the HTML health mail is sent and the bot stops.
* If files are still missing at the cutoff time, a RED mail lists exactly what is missing.

Safety nets (added 10 Oct 2026):
* No internet (e.g. laptop just woke up, Wi-Fi not connected yet): the bot waits for the
  internet (up to "network_wait_minutes", default 20) before reading Gmail.
* If Gmail still cannot be reached, or the health mail cannot be sent, the day is NOT closed:
  the bot tries again every "retry_minutes" (default 10) until "give_up_time" (default 20:00).
* While it runs, the bot asks Windows not to go to sleep (it cannot wake a sleeping laptop).
* If the bot starts late (laptop asleep/off at the start time) the health mail says so.
* If the bot itself crashes, a RED mail with the error is sent (once per different error).

Data Operations page (added 12 Oct 2026):
* After every check the bot updates data/database/run_log.xlsx (src/run_log.py): one row per day
  (GREEN/RED, ready time, SLA), one row per file (arrived, status, rows), all open data issues and
  how fresh each data source is. Power BI reads it for the Data Operations page.

Run:   python src/mail_bot.py                       (daily, started by Task Scheduler)
Test:  python src/mail_bot.py --once --no-email     (one check, no mail sent, HTML saved in logs/)
       python src/mail_bot.py --once --final        (pretend the cutoff has arrived)
"""
import argparse
import email
import hashlib
import imaplib
import json
import logging
import os
import re
import shutil
import socket
import sys
import time
import traceback
from datetime import datetime, timedelta
from email import policy
from email.utils import parseaddr, parsedate_to_datetime

import health_email as he
import run_lock
import run_log
import team_map
from handlers import FileLockedError, load_handler

ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))


# --------------------------------------------------------------------------- #
# Small helpers
# --------------------------------------------------------------------------- #
def norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def stem_key(name):
    """'PV current month (1).xlsx' -> 'pvcurrentmonth'  (extension and trailing digits removed)."""
    return re.sub(r"\d+$", "", norm(os.path.splitext(str(name))[0]))


def P(cfg, key, *parts):
    return os.path.join(cfg["_root"], cfg["paths"][key], *parts)


def hhmm(s):
    h, m = s.split(":")
    return int(h), int(m)


def group_of(cfg, filename):
    """PV / EV from the start of the file name ('PV current month.xlsx', 'pv_report.xlsx')."""
    for g, gc in cfg["groups"].items():
        prefix = gc.get("file_prefix", g)
        if re.match(rf"^{re.escape(prefix)}(?![A-Za-z0-9])", filename, re.I):
            return g
    return None


def rule_for(gc, filename):
    key = stem_key(filename)
    for r in gc.get("files", []):
        if stem_key(r["name"]) == key:
            return r
    return None


def _safe_name(name):
    return re.sub(r'[<>:"/\\|?*]', "_", os.path.basename(name))


def minutes_setting(cfg, key, default, section=None):
    src = cfg.get(section, {}) if section else cfg
    return float(src.get(key, default)) * 60


def keep_awake(on, log=None):
    """Windows only: ask Windows not to go to sleep while the bot runs (on=False gives it back).
    It cannot wake a laptop that is already asleep - that needs the power settings."""
    if os.name != "nt":
        return False
    try:
        import ctypes
        ES_CONTINUOUS, ES_SYSTEM_REQUIRED = 0x80000000, 0x00000001
        flags = ES_CONTINUOUS | (ES_SYSTEM_REQUIRED if on else 0)
        fn = ctypes.windll.kernel32.SetThreadExecutionState
        fn.argtypes, fn.restype = [ctypes.c_uint], ctypes.c_uint
        return fn(flags) != 0
    except Exception as exc:  # never let this stop the bot
        if log:
            log.warning("Could not ask Windows to stay awake: %s", exc)
        return False


def wait_for_network(host, port, max_wait_s, log, step_s=60):
    """Wait until host:port can be reached (e.g. Wi-Fi reconnecting after sleep). Returns True/False."""
    t0, first = time.time(), None
    while True:
        try:
            socket.create_connection((host, port), timeout=15).close()
            if first is not None:
                log.info("Internet is back - %s reachable after %d s", host, time.time() - t0)
            return True
        except OSError as exc:
            waited = time.time() - t0
            if first is None:
                first = exc
                if max_wait_s > 0:
                    log.warning("Cannot reach %s (%s: %s) - no internet yet? Checking every %d s for up to %d min",
                                host, type(exc).__name__, exc, step_s, max_wait_s // 60)
            if waited + step_s > max_wait_s:
                log.error("Still cannot reach %s after %d min (%s: %s)", host, waited // 60, type(exc).__name__, exc)
                return False
            time.sleep(step_s)


def explain_error(err):
    """Turn a technical mailbox/SMTP error into a plain sentence (the technical text is kept after it)."""
    e = str(err)
    low = e.lower()
    if "getaddrinfo" in low or "gaierror" in low or "name or service not known" in low or "nodename" in low:
        why = ("No internet connection - the laptop could not find Gmail's server. Usually the laptop had just "
               "woken up and Wi-Fi was not connected yet, or the internet was down.")
    elif "timed out" in low or "timeout" in low:
        why = "Gmail did not answer in time - slow or unstable internet."
    elif "authenticationfailed" in low or "invalid credentials" in low or "username and password not accepted" in low:
        why = "Gmail refused the login - the app password is wrong or was removed. Create a new app password."
    elif "is not set" in low:
        why = "The Gmail app password is not saved on this computer (environment variable missing)."
    elif "refused" in low or "unreachable" in low or "reset" in low:
        why = "The connection to Gmail was blocked or dropped - check the internet / firewall."
    else:
        return e
    return f"{why} [{e}]"


# --------------------------------------------------------------------------- #
# Fetching
# --------------------------------------------------------------------------- #
def fetch_imap(cfg, day, log):
    ic = cfg["imap"]
    pw = os.environ.get(ic.get("password_env", "GMAIL_APP_PASSWORD"), "")
    if not pw:
        raise RuntimeError(f"environment variable {ic.get('password_env', 'GMAIL_APP_PASSWORD')} is not set")
    senders = [s.strip().lower() for s in cfg["senders"]]
    exts = tuple(cfg.get("extensions", [".xlsx", ".xlsm"]))
    since = (day - timedelta(days=int(ic.get("lookback_days", 1)))).strftime("%d-%b-%Y")

    M = imaplib.IMAP4_SSL(ic.get("host", "imap.gmail.com"))
    parsed = []
    try:
        M.login(ic["user"], pw)
        M.select(ic.get("folder", "INBOX"), readonly=True)
        seen = set()
        for s in senders:                                   # one search per sender, results merged
            typ, data = M.search(None, "SINCE", since, "FROM", f'"{s}"')
            if typ != "OK":
                raise RuntimeError(f"mailbox search failed for sender {s}")
            for num in data[0].split():
                if num in seen:
                    continue
                seen.add(num)
                typ, md = M.fetch(num, "(BODY.PEEK[])")     # PEEK = does not mark the mail as read
                if typ != "OK":
                    continue
                msg = email.message_from_bytes(md[0][1], policy=policy.default)
                frm = parseaddr(str(msg["From"]))[1].lower()
                if frm not in senders:
                    continue
                arrived = parsedate_to_datetime(msg["Date"]).astimezone().replace(tzinfo=None)
                if not (day <= arrived < day + timedelta(days=1)):   # only mails received on that day
                    continue
                parsed.append((arrived, frm, msg))
    finally:
        try:
            M.logout()
        except Exception:
            pass

    parsed.sort(key=lambda x: x[0])                          # oldest first -> newest copy wins
    found = {}
    for arrived, frm, msg in parsed:
        for part in msg.iter_attachments():
            name = part.get_filename()
            if not name or not name.lower().endswith(exts):
                continue
            name = _safe_name(name)
            g = group_of(cfg, name)
            folder = os.path.join(P(cfg, "raw", day.strftime("%Y-%m-%d")), g or "_other")
            os.makedirs(folder, exist_ok=True)
            payload = part.get_payload(decode=True)
            path = os.path.join(folder, name)
            with open(path, "wb") as fh:
                fh.write(payload)
            found[(g, name.lower())] = {"name": name, "path": path, "arrived": arrived, "sender": frm,
                                        "group": g, "sha1": hashlib.sha1(payload).hexdigest()}
    return list(found.values())


def fetch_folder(cfg, day, log):
    """Test / manual mode: drop the Excel files into data/test_inbox/ instead of mailing them."""
    src = P(cfg, "test_inbox")
    exts = tuple(cfg.get("extensions", [".xlsx", ".xlsm"]))
    out = []
    if not os.path.isdir(src):
        return out
    for f in sorted(os.listdir(src)):
        if not f.lower().endswith(exts) or f.startswith("~$"):
            continue
        g = group_of(cfg, f)
        folder = os.path.join(P(cfg, "raw", day.strftime("%Y-%m-%d")), g or "_other")
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, f)
        shutil.copy2(os.path.join(src, f), path)
        with open(path, "rb") as fh:
            sha = hashlib.sha1(fh.read()).hexdigest()
        out.append({"name": f, "path": path, "arrived": datetime.fromtimestamp(os.path.getmtime(path)),
                    "sender": "(test folder)", "group": g, "sha1": sha})
    return out


def collect(cfg, day, log):
    imap_mode = cfg.get("source_mode", "imap") == "imap"
    fetch = fetch_imap if imap_mode else fetch_folder
    step = int(cfg.get("network_check_seconds", 60))
    files, err = [], None
    if imap_mode:
        ic = cfg["imap"]
        wait_for_network(ic.get("host", "imap.gmail.com"), int(ic.get("port", 993)),
                         minutes_setting(cfg, "network_wait_minutes", 20), log, step)
    for attempt in (1, 2, 3):
        try:
            files, err = fetch(cfg, day, log), None
            break
        except (OSError, imaplib.IMAP4.abort) as exc:      # network hiccup -> try again shortly
            err = f"{type(exc).__name__}: {exc}"
            log.warning("Mailbox read attempt %d/3 failed: %s", attempt, err)
            if attempt < 3:
                time.sleep(min(step, 30))
        except Exception as exc:                            # login refused, bad config ... -> no point retrying
            err = f"{type(exc).__name__}: {exc}"
            break
    if err:
        log.error("Mailbox read failed: %s", err)
    groups = {}
    for g, gc in cfg["groups"].items():
        gfiles = [f for f in files if f["group"] == g]
        required = [r["name"] for r in gc.get("files", []) if r.get("required", True)]
        missing = [n for n in required if not any(stem_key(f["name"]) == stem_key(n) for f in gfiles)]
        groups[g] = {"files": gfiles, "expected": required, "missing": missing}
    others = [f for f in files if f["group"] is None]
    return groups, others, err


# --------------------------------------------------------------------------- #
# Processing
# --------------------------------------------------------------------------- #
def process_one(cfg, g, f, day, state, log):
    key = f"{g}|{f['name']}|{f['sha1']}"
    if key in state["results"]:
        return state["results"][key]
    row = {"name": f["name"], "arrived": f["arrived"].strftime("%H:%M"), "sender": f["sender"],
           "status": "OK", "rows": "-", "new": "-", "updated": "-", "unchanged": "-", "note": ""}
    rule = rule_for(cfg["groups"][g], f["name"])
    if rule is None:
        row.update(status="Received", note="Saved - no processing rule yet")
        return row          # not cached: if the config is fixed later, the file is picked up next run
    attempts = int(cfg.get("lock_retries", 3))
    for attempt in range(1, attempts + 1):
        try:
            handler = load_handler(rule["handler"])
            res = handler.process(f["path"], g, day, P(cfg, "database"))
            row.update(status="OK", rows=res["rows_read"], new=res["new"], updated=res["updated"],
                       unchanged=res["unchanged"], note="; ".join(res.get("notes", [])))
            state["results"][key] = row
            log.info("%s %s loaded: %s new, %s updated, %s unchanged", g, f["name"],
                     res["new"], res["updated"], res["unchanged"])
            return row
        except FileLockedError as exc:
            log.warning("attempt %s/%s: %s", attempt, attempts, exc)
            row.update(status="FAILED", note=str(exc))
            if attempt < attempts:
                time.sleep(int(cfg.get("lock_wait_seconds", 20)))
        except Exception as exc:
            log.exception("processing failed for %s", f["name"])
            row.update(status="FAILED", note=f"{type(exc).__name__}: {exc}")
            break
    return row


def build_summary(cfg):
    names = []
    for gc in cfg["groups"].values():
        for r in gc.get("files", []):
            if r["handler"] not in names:
                names.append(r["handler"])
    for n in names:
        try:
            mod = load_handler(n)
            if hasattr(mod, "summary"):
                m, d = mod.summary(P(cfg, "database"))
                if m:
                    return m, d
        except Exception:
            pass
    return None, None


def send_with_retry(cfg, subject, html, log, attempts=3):
    """Send the health mail; wait for the internet first and retry a few times."""
    sc = cfg["smtp"]
    step = int(cfg.get("network_check_seconds", 60))
    wait_for_network(sc.get("host", "smtp.gmail.com"), int(sc.get("port", 465)),
                     min(minutes_setting(cfg, "network_wait_minutes", 20), 300), log, step)
    ok, msg = False, ""
    for i in range(1, attempts + 1):
        ok, msg = he.send_email(sc, subject, html)
        if ok or "not set" in msg or "Authentication" in msg or i == attempts:
            break
        log.warning("Sending the health mail failed (attempt %d/%d): %s", i, attempts, msg)
        time.sleep(min(step, 30))
    return ok, msg


def deliver(cfg, day, report, state, log, send, sig, final):
    """Save the HTML in logs/ and mail it. Returns True when the mail went out (or mailing is off)."""
    now = datetime.now()
    html, subject = he.build_html(report), he.build_subject(report)
    os.makedirs(P(cfg, "logs"), exist_ok=True)
    key = f"{report['status']}|{sig}"
    if state.get("saved_key") == key and state.get("saved_file"):   # same problem again -> overwrite, no new file
        fn = state["saved_file"]
    else:
        fn = P(cfg, "logs", f"health_{day:%Y-%m-%d}_{now:%H%M%S}_{report['status']}.html")
    with open(fn, "w", encoding="utf-8") as fh:
        fh.write(html)
    state["saved_key"], state["saved_file"] = key, fn
    log.info("Health mail saved: %s", fn)
    ok = True
    if send and cfg.get("send_email", True):
        ok, msg = send_with_retry(cfg, subject, html, log)
        log.log(logging.INFO if ok else logging.ERROR, "Health mail %s: %s",
                "sent" if ok else "NOT sent", explain_error(msg) if not ok else msg)
    if ok:
        state["sent_sig"] = sig
        if final:
            state["final_sent"] = True
    return ok


def refresh_team_map(cfg, day, state, log):
    """Once a day: rebuild team_map.xlsx (CA -> TL -> SM) from data/reference/Team_Mapping.xlsx."""
    src = os.path.join(cfg["_root"], cfg.get("team_map", {}).get("source_file", team_map.SOURCE_FILE))
    if not os.path.exists(src):
        return None                                   # no team list yet -> feature off, no noise
    res = state.get("team_map")
    if not (res and res.get("ok")):
        res = team_map.refresh(cfg["_root"], cfg, P(cfg, "database"), today=day, log=log)
        state["team_map"] = res
    return res


def team_map_notes(res):
    if not res:
        return []
    if not res.get("ok"):
        return [f"Team list (Team_Mapping.xlsx) was NOT updated today: {res.get('note')}"]
    if res.get("unmapped"):
        names = ", ".join(res["unmapped"][:12]) + (" ..." if len(res["unmapped"]) > 12 else "")
        return [f"{len(res['unmapped'])} CA(s) have records in Siebel but are not in Team_Mapping.xlsx, "
                f"so they show as 'Not in team list': {names}. Add them if they still work here."]
    return []


def run_cycle(cfg, day, final, state, log, send=True, notes=None):
    now = datetime.now()
    gdata, others, err = collect(cfg, day, log)
    missing_any = any(d["missing"] for d in gdata.values())
    state["last_problem"] = bool(err)

    if not final and missing_any and not err:
        log.info("Still waiting for: %s | received so far: %s",
                 {g: d["missing"] for g, d in gdata.items() if d["missing"]},
                 {g: [f["name"] for f in d["files"]] for g, d in gdata.items()})
        return False

    errors = [f"Cannot read the mailbox - {explain_error(err)}"] if err else []
    groups_report, n_loaded = [], 0
    for g, d in gdata.items():
        order = {stem_key(r["name"]): i for i, r in enumerate(cfg["groups"][g].get("files", []))}
        files = sorted(d["files"], key=lambda x: (order.get(stem_key(x["name"]), 99), x["arrived"]))  # oldest first, newest wins
        received = [process_one(cfg, g, f, day, state, log) for f in files]
        for r in received:
            if r["status"] == "FAILED":
                errors.append(f"{g}: {r['name']} failed to load - {r['note']}")
            elif r["status"] == "OK":
                n_loaded += 1
        if not err:
            for m in d["missing"]:
                errors.append(f"{g}: '{m}.xlsx' was not received from the listed senders"
                              + (f" by {cfg['schedule']['cutoff_time']}" if final else ""))
        groups_report.append({"name": g, "expected": d["expected"], "checked": not err,
                              "missing": [] if err else d["missing"], "received": received})

    tm_res = refresh_team_map(cfg, day, state, log) if not err else None
    master, dq = build_summary(cfg)
    target = cfg["schedule"].get("sla_ready_by")
    ready = now.strftime("%H:%M")
    status = "GREEN" if not errors else "RED"
    met = None
    if target:
        met = (ready <= target) if status == "GREEN" else (False if (final or ready > target) else None)

    if status == "GREEN":
        headline = (f"All files received and loaded ({n_loaded} file{'s' if n_loaded != 1 else ''})"
                    f" \u2013 {master['total'] if master else 0} bookings in database")
    else:
        n_missing = 0 if err else sum(len(d["missing"]) for d in gdata.values())
        n_failed = sum(1 for g in groups_report for r in g["received"] if r["status"] == "FAILED")
        bits = []
        if n_missing:
            bits.append(f"{n_missing} file(s) missing")
        if n_failed:
            bits.append(f"{n_failed} file(s) failed to load")
        if err:
            bits.append("Gmail could not be reached \u2013 data NOT loaded")
        headline = " \u00b7 ".join(bits) or "Problem detected"

    notes = list(notes or []) + team_map_notes(tm_res)
    if err and cfg.get("_give_up"):
        retry = cfg["schedule"].get("retry_minutes", 10)
        notes = list(notes or []) + [f"The bot keeps trying every {retry} minutes until {cfg['_give_up']}. "
                                     f"You will get a new mail as soon as the data is loaded."]

    report = {"status": status, "headline": headline, "run_date": day.strftime("%a %d %b %Y"),
              "run_time": now.strftime("%H:%M"), "groups": groups_report, "errors": errors,
              "notes": list(notes or []),
              "others": [{"name": o["name"], "sender": o["sender"], "arrived": o["arrived"].strftime("%H:%M")}
                         for o in others],
              "master": master, "dq": dq,
              "master_title": "Master database (unchanged \u2013 nothing was loaded)" if err else None,
              "sla": {"target": target, "ready_at": ready if status == "GREEN" else None,
                      "met": False if (err and now.strftime("%H:%M") > target) else (None if err else met)}
              if target else None}

    try:                                   # Data Operations page (Power BI) - never allowed to stop the bot
        run_log.record_day(cfg, day, report, state, log)
    except Exception as exc:
        log.warning("run_log.xlsx (Data Operations page) not updated: %s", exc)

    sig = "|".join(errors)
    mail_ok = True
    if status == "GREEN" or sig != state.get("sent_sig") or (final and not state.get("final_sent")):
        mail_ok = deliver(cfg, day, report, state, log, send, sig, final)

    if err or not mail_ok:
        state["last_problem"] = True
        return False                       # not finished: try again later (main loop decides until when)
    if status == "GREEN" or final:
        state["done"] = True
        return True
    return False


# --------------------------------------------------------------------------- #
# Main loop
# --------------------------------------------------------------------------- #
def crash_report(cfg, day, state, exc, log, send):
    """The bot itself hit an unexpected error: send a RED mail once per different error."""
    try:
        where = traceback.extract_tb(exc.__traceback__)[-1] if exc.__traceback__ else None
        text = f"{type(exc).__name__}: {exc}"
        sig = f"CRASH|{text}"
        if state.get("sent_sig") == sig:
            return
        now = datetime.now()
        errors = [f"The bot stopped with an unexpected error: {text}"]
        if where:
            errors.append(f"Where: {os.path.basename(where.filename)}, line {where.lineno} ({where.name})")
        report = {"status": "RED", "headline": "The bot hit an error – data NOT loaded",
                  "run_date": day.strftime("%a %d %b %Y"), "run_time": now.strftime("%H:%M"),
                  "groups": [], "others": [], "errors": errors, "master": None, "dq": None, "sla": None,
                  "notes": [f"Full details are in logs\\bot_{day:%Y-%m-%d}.log - send that file for a fix. "
                            f"The bot tries again at the next check."]}
        deliver(cfg, day, report, state, log, send, sig, final=False)
    except Exception:
        log.exception("Could not even build the error mail")


def load_state(path):
    if os.path.exists(path):
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    return {"done": False, "sent_sig": "", "results": {}}


def save_state(path, state):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(state, fh, indent=2, default=lambda o: o.item() if hasattr(o, "item") else str(o))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", default=os.path.join(ROOT, "config", "config.json"))
    ap.add_argument("--date", help="YYYY-MM-DD (default: today)")
    ap.add_argument("--once", action="store_true", help="run a single cycle and exit")
    ap.add_argument("--final", action="store_true", help="behave as if the cutoff time was reached")
    ap.add_argument("--no-email", action="store_true", help="do not send mail (HTML is still saved to logs/)")
    ap.add_argument("--reset-state", action="store_true", help="forget today's progress and start again")
    a = ap.parse_args()

    with open(a.config, encoding="utf-8") as fh:
        cfg = json.load(fh)
    cfg["_root"] = ROOT
    os.environ["BOOKINGS_CONFIG"] = os.path.abspath(a.config)      # handlers read their settings from the same file
    day = (datetime.strptime(a.date, "%Y-%m-%d") if a.date
           else datetime.now().replace(hour=0, minute=0, second=0, microsecond=0))

    os.makedirs(P(cfg, "logs"), exist_ok=True)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
                        handlers=[logging.FileHandler(P(cfg, "logs", f"bot_{day:%Y-%m-%d}.log"), encoding="utf-8"),
                                  logging.StreamHandler(sys.stdout)])
    log = logging.getLogger("bot")
    lock = run_lock.RunLock(P(cfg, "state"), f"mail_bot for {day:%d %b %Y}")
    try:
        lock.acquire()
    except run_lock.AlreadyRunning as held:
        log.warning("Another copy of the bot is already running (%s). This copy stops without changing anything, "
                    "so the two cannot damage each other's files. To run by hand, first close the other bot window "
                    "(if you cannot see it: Task Manager > Details > python.exe).", held)
        return

    sp = P(cfg, "state", f"{day:%Y-%m-%d}.json")
    if a.reset_state and os.path.exists(sp):
        os.remove(sp)
    state = load_state(sp)
    if state["done"] and not a.once:
        log.info("Already completed for %s - nothing to do.", day.date())
        return

    sch = cfg["schedule"]
    sh, sm = hhmm(sch["start_time"])
    ch, cm = hhmm(sch["cutoff_time"])
    gh, gm = hhmm(sch.get("give_up_time", "20:00"))
    start_dt, cutoff_dt = day.replace(hour=sh, minute=sm), day.replace(hour=ch, minute=cm)
    give_up_dt = max(day.replace(hour=gh, minute=gm), cutoff_dt)
    cfg["_give_up"] = None if a.once else give_up_dt.strftime("%H:%M")     # shown in the RED mail
    poll = minutes_setting(cfg, "poll_minutes", 30, "schedule")
    retry = minutes_setting(cfg, "retry_minutes", 10, "schedule")

    started = datetime.now()
    log.info("Bot started at %s (start %s, cutoff %s, keeps retrying problems until %s)%s",
             started.strftime("%H:%M"), sch["start_time"], sch["cutoff_time"],
             give_up_dt.strftime("%H:%M"), " [manual run]" if a.once else "")
    notes = []
    if not a.once and not a.date:
        first = state.setdefault("first_start", started.strftime("%H:%M"))
        fh_, fm_ = hhmm(first)
        late = (day.replace(hour=fh_, minute=fm_) - start_dt).total_seconds() / 60
        if late > int(sch.get("late_start_warn_minutes", 20)):
            msg = (f"The bot only started at {first} – {int(late // 60)} h {int(late % 60)} min after the "
                   f"scheduled {sch['start_time']}. The laptop was most likely asleep or switched off at "
                   f"{sch['start_time']}; Windows ran the task when the laptop woke up.")
            log.warning(msg)
            notes.append(msg)
    if cfg.get("keep_awake", True) and keep_awake(True, log):
        log.info("Asked Windows to stay awake while the bot runs")

    try:
        while True:
            now = datetime.now()
            if not a.once and now < start_dt:
                wait = (start_dt - now).total_seconds()
                log.info("Waiting until %s (%d min)", start_dt.strftime("%H:%M"), wait // 60)
                time.sleep(wait)
                continue
            final = a.final or now >= cutoff_dt
            try:
                done = run_cycle(cfg, day, final, state, log, send=not a.no_email, notes=notes)
            except Exception as exc:
                log.exception("Unexpected error in cycle")
                state["last_problem"] = True
                crash_report(cfg, day, state, exc, log, send=not a.no_email)
                done = False
            save_state(sp, state)
            now = datetime.now()
            if done or a.once:
                break
            if final and now >= give_up_dt:
                log.error("Giving up for today at %s - the day is NOT complete. Run manually: "
                          "python src\\mail_bot.py --once --final", now.strftime("%H:%M"))
                break
            problem = bool(state.get("last_problem"))
            nxt = now + timedelta(seconds=retry if (final or problem) else poll)   # problem -> retry sooner
            nxt = min(nxt, give_up_dt if final else cutoff_dt)
            if final or problem:
                log.info("Next try at %s", nxt.strftime("%H:%M"))
            time.sleep(max((nxt - datetime.now()).total_seconds(), 1))
    finally:
        keep_awake(False)
        save_state(sp, state)
        log.info("Bot finished for %s (done=%s)", day.date(), state["done"])


if __name__ == "__main__":
    main()
