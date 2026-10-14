r"""
run_lock.py  -  only ONE copy of the bot may work on the database at a time
============================================================================
Added 13 Oct 2026. That evening a manual run and the Task Scheduler run started at the same minute.
Both wrote master_bookings.xlsx at the same time and the file was damaged (BadZipFile).

The bot (and backfill.py) now take a lock on state/bot.lock when they start. A second copy that finds
the lock taken writes one line ("another copy is already running since ...") and stops without touching
anything. Windows releases the lock by itself when the first copy ends or crashes - nothing to clean up.
"""
import os
from datetime import datetime

LOCK_FILE = "bot.lock"
LOCK_BYTE = 4096          # the lock sits on this byte; the text in front of it says who holds the lock


class AlreadyRunning(Exception):
    """Another copy holds the lock. str(exc) says which one and since when."""


class RunLock:
    def __init__(self, folder, who):
        self.path = os.path.join(folder, LOCK_FILE)
        self.who = who
        self.fd = None

    def acquire(self):
        os.makedirs(os.path.dirname(os.path.abspath(self.path)), exist_ok=True)
        fd = os.open(self.path, os.O_RDWR | os.O_CREAT, 0o644)
        try:
            _lock(fd)
        except OSError:
            os.close(fd)
            raise AlreadyRunning(self.holder()) from None
        os.lseek(fd, 0, os.SEEK_SET)
        os.write(fd, f"{self.who}, process {os.getpid()}, started {datetime.now():%d %b %Y %H:%M}".ljust(200).encode())
        self.fd = fd
        return self

    def release(self):
        if self.fd is not None:
            try:
                _unlock(self.fd)
            except OSError:
                pass
            os.close(self.fd)
            self.fd = None

    def holder(self):
        try:
            with open(self.path, "rb") as fh:
                text = fh.read(200).decode("utf-8", "replace").strip()
            return text or "another copy"
        except OSError:
            return "another copy"

    def __enter__(self):
        return self.acquire()

    def __exit__(self, *exc):
        self.release()


def _lock(fd):
    if os.name == "nt":
        import msvcrt
        os.lseek(fd, LOCK_BYTE, os.SEEK_SET)
        msvcrt.locking(fd, msvcrt.LK_NBLCK, 1)
    else:
        import fcntl
        fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)


def _unlock(fd):
    if os.name == "nt":
        import msvcrt
        os.lseek(fd, LOCK_BYTE, os.SEEK_SET)
        msvcrt.locking(fd, msvcrt.LK_UNLCK, 1)
    else:
        import fcntl
        fcntl.flock(fd, fcntl.LOCK_UN)
