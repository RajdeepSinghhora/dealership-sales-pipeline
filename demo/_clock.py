r"""
_clock.py  -  run the real bot (src/mail_bot.py) as if the clock showed another date and time
==============================================================================================
Used by run_demo.py to replay two weeks of mornings in a minute. The bot code is NOT changed:
this wrapper only makes datetime.now() (and the log time stamps) return the simulated time.

    python demo/_clock.py "2026-10-05 10:30" --config <file> --date 2026-10-05 --once --no-email
"""
import datetime as _dt
import logging
import os
import sys
import time

SRC = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src"))
sys.path.insert(0, SRC)


def main():
    target = _dt.datetime.strptime(sys.argv[1], "%Y-%m-%d %H:%M")
    offset = target - _dt.datetime.now()

    class SimulatedDateTime(_dt.datetime):
        @classmethod
        def now(cls, tz=None):
            return _dt.datetime.now(tz) + offset

    logging.Formatter.converter = staticmethod(lambda secs: time.localtime(secs + offset.total_seconds()))

    import handlers
    import mail_bot
    import run_log
    import team_map
    from handlers import bookings, open_pipeline, ssi
    for module in (handlers, mail_bot, run_log, team_map, bookings, open_pipeline, ssi):
        if hasattr(module, "datetime"):
            module.datetime = SimulatedDateTime

    sys.argv = ["mail_bot.py"] + sys.argv[2:]
    mail_bot.main()


if __name__ == "__main__":
    main()
