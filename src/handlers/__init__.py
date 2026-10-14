"""
handlers/  -  one Python file per Excel file type.

A handler is any module in this folder (name not starting with "_") that defines:

    process(path, group, load_date, db_dir) -> dict
        path       full path of the downloaded Excel file
        group      "PV" or "EV" (taken from the file name)
        load_date  the day being processed (datetime)
        db_dir     folder where output workbooks must be written (data/database)
        returns    {"rows_read": int, "new": int, "updated": int, "unchanged": int, "notes": [str]}

    summary(db_dir) -> (dict | None, dict | None)        # OPTIONAL, used only for the health mail

To add a new file type:  copy _template.py -> <name>.py, fill it in, then add one line in
config/config.json under the right group ("files": [{"name": "...", "handler": "<name>"}]).
"""
import importlib


class FileLockedError(Exception):
    """Raise this when the output workbook is open in Excel and cannot be saved."""


def load_handler(name):
    try:
        return importlib.import_module(f"handlers.{name}")
    except ModuleNotFoundError as exc:
        raise RuntimeError(f"Handler '{name}' not found - create src/handlers/{name}.py ({exc})") from exc


# --------------------------------------------------------------------------- #
# Safe writing (added 13 Oct 2026, after two copies of the bot ran at the same time
# and damaged master_bookings.xlsx)
#   temp_name()     every program writes its own temp file (process id in the name)
#   verify_xlsx()   the temp file is read back completely BEFORE it replaces the workbook
#   daily_backup()  one good copy of each workbook per day in data/database/backup (kept 30 days)
# --------------------------------------------------------------------------- #
import glob
import os
import shutil
import zipfile
from datetime import datetime, timedelta

BACKUP_DAYS = 30


class DamagedWorkbookError(FileLockedError):
    """A workbook that was just written cannot be read back. The old file is kept and the bot tries again."""


def temp_name(path):
    """'master_bookings.xlsx' -> 'master_bookings.<process id>.tmp.xlsx' (two programs never share a temp file)."""
    root, ext = os.path.splitext(path)
    return f"{root}.{os.getpid()}.tmp{ext}"


def verify_xlsx(path):
    """Raise DamagedWorkbookError unless every part of the .xlsx (a zip file) can be read back."""
    try:
        with zipfile.ZipFile(path) as z:
            bad = z.testzip()
            if bad is None and "xl/workbook.xml" not in z.namelist():
                bad = "no workbook inside"
    except (zipfile.BadZipFile, OSError, EOFError) as exc:
        bad = f"{type(exc).__name__}: {exc}"
    if bad:
        raise DamagedWorkbookError(f"{os.path.basename(path)} is damaged ({bad})")


def remove_quietly(path):
    try:
        os.remove(path)
    except OSError:
        pass


def daily_backup(path, keep_days=BACKUP_DAYS, today=None):
    """Before the first change of the day copy the workbook to backup/<name>_<yyyy-mm-dd>.xlsx.
    A damaged workbook is never copied (the older good copies stay). Copies older than keep_days are deleted."""
    if not os.path.exists(path):
        return None
    today = today or datetime.now()
    folder = os.path.join(os.path.dirname(path), "backup")
    stem = os.path.splitext(os.path.basename(path))[0]
    dest = os.path.join(folder, f"{stem}_{today:%Y-%m-%d}.xlsx")
    if not os.path.exists(dest):
        try:
            verify_xlsx(path)
        except DamagedWorkbookError:
            return None
        os.makedirs(folder, exist_ok=True)
        tmp = temp_name(dest)
        shutil.copy2(path, tmp)
        os.replace(tmp, dest)
    oldest = f"{today - timedelta(days=keep_days):%Y-%m-%d}"
    for old in glob.glob(os.path.join(folder, f"{stem}_????-??-??.xlsx")):
        if os.path.basename(old)[len(stem) + 1:-5] < oldest:
            remove_quietly(old)
    return dest
