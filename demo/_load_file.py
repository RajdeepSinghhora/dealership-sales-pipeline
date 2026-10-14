r"""
_load_file.py  -  load one older report file through its real handler (used by run_demo.py --full)
==================================================================================================
The daily SSI report only covers the current financial year, so last year's surveys come from that
year's final file. This loads it exactly as the bot would, dated on the day it belongs to.

    python demo/_load_file.py ssi PV "<file>" 2026-03-31 <config.json>
"""
import datetime as _dt
import json
import os
import sys

SRC = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src"))
sys.path.insert(0, SRC)


def main():
    handler, group, path, day, config = sys.argv[1:6]
    when = _dt.datetime.strptime(day, "%Y-%m-%d").replace(hour=20)
    offset = when - _dt.datetime.now()

    class SimulatedDateTime(_dt.datetime):
        @classmethod
        def now(cls, tz=None):
            return _dt.datetime.now(tz) + offset

    os.environ["BOOKINGS_CONFIG"] = os.path.abspath(config)
    with open(config, encoding="utf-8") as fh:
        cfg = json.load(fh)
    import handlers
    module = handlers.load_handler(handler)
    for m in (handlers, module):
        if hasattr(m, "datetime"):
            m.datetime = SimulatedDateTime
    res = module.process(path, group, when.replace(hour=0), cfg["paths"]["database"])
    print(f"{group} {os.path.basename(path)}: {res['rows_read']} rows, {res['new']} new - " + "; ".join(res.get("notes", [])))


if __name__ == "__main__":
    main()
