"""Only one copy of the bot at a time, verified writes, daily backups and the restore tool.
Regression for a real incident: a manual run and the scheduled run started in the same minute, both wrote
master_bookings.xlsx through the same temp file, and the workbook was damaged (BadZipFile)."""
import json
import os
import subprocess
import sys
from datetime import datetime

import pandas as pd
import pytest

import handlers
import run_lock
from conftest import ROOT

PROBE = ("import sys\nsys.path.insert(0, sys.argv[1])\nimport run_lock\n"
         "try:\n    run_lock.RunLock(sys.argv[2], 'probe').acquire()\n    print('STARTED')\n"
         "except run_lock.AlreadyRunning as e:\n    print('STOPPED', e)")


def probe(state):
    out = subprocess.run([sys.executable, "-c", PROBE, os.path.join(ROOT, "src"), str(state)],
                         capture_output=True, text=True)
    return out.stdout.strip()


def test_second_copy_is_stopped_while_the_first_runs(tmp_path):
    lock = run_lock.RunLock(tmp_path, "mail_bot for 12 Oct 2026").acquire()
    try:
        answer = probe(tmp_path)
        assert answer.startswith("STOPPED") and "mail_bot for 12 Oct 2026" in answer
    finally:
        lock.release()
    assert probe(tmp_path) == "STARTED"          # free again once the first copy has finished


def test_bot_stops_at_once_when_another_copy_runs(tmp_path):
    with open(os.path.join(ROOT, "config", "config.demo.json"), encoding="utf-8") as fh:
        cfg = json.load(fh)
    cfg["paths"] = {k: str(tmp_path / v) for k, v in cfg["paths"].items()}
    cfg_path = tmp_path / "config.json"
    cfg_path.write_text(json.dumps(cfg), encoding="utf-8")
    lock = run_lock.RunLock(cfg["paths"]["state"], "scheduled run").acquire()
    try:
        res = subprocess.run([sys.executable, os.path.join(ROOT, "src", "mail_bot.py"), "--config", str(cfg_path),
                              "--date", "2026-10-12", "--once", "--no-email"], capture_output=True, text=True,
                             timeout=120)
    finally:
        lock.release()
    assert "Another copy of the bot is already running (scheduled run" in res.stdout
    assert not (tmp_path / "state" / "2026-10-12.json").exists()          # it changed nothing


def write_book(path, rows=3):
    pd.DataFrame({"a": range(rows)}).to_excel(path, index=False)


def damage(path):
    import zipfile
    off = zipfile.ZipFile(path).getinfo("xl/workbook.xml").header_offset
    with open(path, "r+b") as fh:
        fh.seek(off)
        fh.write(b"\x00\x00\x00\x00")


def test_temp_files_are_per_process_and_checked(tmp_path):
    p = tmp_path / "master_bookings.xlsx"
    assert str(os.getpid()) in handlers.temp_name(str(p))
    write_book(p)
    handlers.verify_xlsx(str(p))
    damage(p)
    with pytest.raises(handlers.DamagedWorkbookError):
        handlers.verify_xlsx(str(p))


def test_daily_backup_keeps_good_copies_only(tmp_path):
    p = tmp_path / "open_pipeline.xlsx"
    write_book(p)
    b = handlers.daily_backup(str(p), today=datetime(2026, 10, 12))
    assert os.path.basename(b) == "open_pipeline_2026-10-12.xlsx"
    old = tmp_path / "backup" / "open_pipeline_2026-08-01.xlsx"
    write_book(old)
    handlers.daily_backup(str(p), today=datetime(2026, 10, 13))
    assert not old.exists()                                                       # older than 30 days -> removed
    damage(p)
    assert handlers.daily_backup(str(p), today=datetime(2026, 10, 14)) is None     # damaged file is never backed up
    assert not (tmp_path / "backup" / "open_pipeline_2026-10-14.xlsx").exists()


def test_restore_puts_back_the_newest_good_backup(tmp_path):
    db = tmp_path / "data" / "database"
    (db / "backup").mkdir(parents=True)
    (tmp_path / "state").mkdir()
    (tmp_path / "state" / "2026-10-13.json").write_text("{}")
    write_book(db / "master_bookings.xlsx", rows=5)
    write_book(db / "backup" / "master_bookings_2026-10-13.xlsx", rows=4)
    write_book(db / "backup" / "master_bookings_2026-10-12.xlsx", rows=3)
    damage(db / "backup" / "master_bookings_2026-10-13.xlsx")          # newest backup is damaged too -> skipped
    damage(db / "master_bookings.xlsx")
    cfg = tmp_path / "config.json"
    cfg.write_text(json.dumps({"paths": {"database": str(db), "state": str(tmp_path / "state")}}))
    script = os.path.join(ROOT, "src", "check_workbooks.py")
    out = subprocess.run([sys.executable, script, "--config", str(cfg), "--restore"], capture_output=True, text=True)
    assert "restored the copy of 2026-10-12" in out.stdout
    assert "--date 2026-10-12" in out.stdout and "--date 2026-10-13" in out.stdout   # days to load again, in order
    assert len(pd.read_excel(db / "master_bookings.xlsx")) == 3
    assert list((db / "backup").glob("master_bookings_DAMAGED_*.xlsx"))
    again = subprocess.run([sys.executable, script, "--config", str(cfg)], capture_output=True, text=True)
    assert "All workbooks can be read." in again.stdout
