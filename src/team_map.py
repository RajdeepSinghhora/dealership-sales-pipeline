r"""
team_map.py  -  CA -> TL -> SM mapping for Power BI
====================================================
Source (kept by MIS): data/reference/Team_Mapping.xlsx, sheet "Team"
    CA NAME | TL NAME | SM NAME | Team Type | Split login (PV/EV) | Note
    * Team Type: EV for the EV team, PV for the others. It is the CA's TEAM,
      not what was sold - Power BI compares it with the booking's Vehicle_Type to show cross-sales
      (PV team selling EV and the other way round). The CA keeps the credit either way.
    * Split login is optional: leave it blank normally. Fill PV or EV only when one CA has two
      CRM logins in two different teams (e.g. one CA with an EV login under TL A and a PV login under TL B).
    * When teams change, just edit this file. The bot notices the change the next morning
      and keeps the old team in Team_History with an end date.

The bot runs this once a day and writes data/database/team_map.xlsx:
  Team_Map      one row per CA name + vehicle type seen in Siebel (plus CAs only in the list).
                CA_Key = "<Employee Full Name>|<PV or EV>" is the join key to Bookings and Open_Now
                (Power BI: CA_Key column on both). CAs in Siebel that are not in the list get
                TL / SM = "Not in team list", so their numbers never disappear.
  Team_History  who was in which team from which date (Valid_From / Valid_To), starts the first day.
  Map_Issues    duplicates merged, spelling differences, conflicts, Siebel CAs not in the list.

Cleaning rules:
  * rows with an empty TL are skipped; names are made unique ignoring case, dots and spaces,
    and shown in Proper Case
  * spelling differences are matched word by word, never loosely: every word must be close
    (KAVYA IYER / KAVYA IYYER = same person, ROHAN MEHTA / ROHAN MEHRA = two people),
    or the shorter name must sit inside the longer one (ARJUN RAO = ARJUN KUMAR RAO).
    One-word names (e.g. RAVI, SONU) are never matched by spelling.

Run by hand:  python src/team_map.py
"""
import argparse
import difflib
import json
import os
import re
import sys
from datetime import datetime

import pandas as pd
from handlers import daily_backup, temp_name, verify_xlsx

ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
OUTPUT_FILE = "team_map.xlsx"
SOURCE_FILE = os.path.join("data", "reference", "Team_Mapping.xlsx")
BAD = {"", "#N/A", "N/A", "NA", "#REF!", "#VALUE!", "#NAME?", "0", "-", "NAN", "NONE"}
SIMILAR = 0.85                     # whole-name similarity needed to treat two names as the same person
WORD_SIMILAR = 0.8                 # ... and every word must be at least this similar
PLACEHOLDER = re.compile(r"^(unspecified|seller not found)", re.I)
SELLER_NOT_FOUND = "SELLER NOT FOUND"       # SSI surveys whose seller is in no list (same text as in handlers/ssi.py)
NOT_IN_LIST = "Not in team list"
VTS = ("PV", "EV")


# --------------------------------------------------------------------------- #
def clean_text(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def key_of(name):
    """'  rOHan  mehta ' -> 'ROHAN MEHTA' (case, dots and extra spaces ignored)."""
    return re.sub(r"\s+", " ", re.sub(r"[^A-Za-z ]", " ", str(name))).strip().upper()


def proper(name):
    s = clean_text(name)
    return " ".join(w[:1].upper() + w[1:].lower() if w.isalpha() else w.title() for w in s.split(" "))


def similar(a, b):
    return difflib.SequenceMatcher(None, a, b).ratio()


def same_person(a, b):
    """Strict name match on keys. Returns a score (0 = different people)."""
    if a == b:
        return 1.0
    ta, tb = a.split(), b.split()
    if len(ta) == len(tb) and len(ta) >= 2:
        if all(similar(x, y) >= WORD_SIMILAR for x, y in zip(ta, tb)):
            s = similar(a, b)
            return s if s >= SIMILAR else 0.0
        return 0.0
    short, long_ = (ta, tb) if len(ta) < len(tb) else (tb, ta)
    if len(short) >= 2 and all(w in long_ for w in short):
        return 0.9
    return 0.0


# --------------------------------------------------------------------------- #
def read_source(path):
    """Team_Mapping.xlsx (sheet 'Team' or the first sheet) or a .csv -> list of row dicts."""
    if path.lower().endswith(".csv"):
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
    else:
        with pd.ExcelFile(path) as x:
            sheet = "Team" if "Team" in x.sheet_names else x.sheet_names[0]
            df = x.parse(sheet, dtype=str, keep_default_na=False)
    cols = {re.sub(r"[^a-z]", "", str(c).lower()): c for c in df.columns}

    def pick(*names):
        for n in names:
            if n in cols:
                return cols[n]
        return None

    c_ca, c_tl, c_sm = pick("caname", "ca"), pick("tlname", "tl"), pick("smname", "sm")
    c_vt = pick("splitloginpvev", "splitlogin", "loginpvev", "vehicletype", "vt", "pvev")
    c_tt = pick("teamtype", "team")
    c_note = pick("note", "notes", "remark", "remarks")
    if not (c_ca and c_tl and c_sm):
        raise ValueError(f"{os.path.basename(path)} needs the columns CA NAME, TL NAME, SM NAME "
                         f"(found: {', '.join(map(str, df.columns))})")
    out = []
    for i, r in df.iterrows():
        vt = clean_text(r[c_vt]).upper() if c_vt else ""
        tt = clean_text(r[c_tt]).upper() if c_tt else ""
        out.append({"Row": i + 2, "CA_raw": clean_text(r[c_ca]), "TL": clean_text(r[c_tl]), "SM": clean_text(r[c_sm]),
                    "VT": vt if vt in VTS else "", "TT": tt if tt in VTS else "",
                    "Note": clean_text(r[c_note]) if c_note else ""})
    return out


def build(rows, siebel_pairs, today):
    """rows from read_source; siebel_pairs = {(Employee Full Name, 'PV'|'EV')}. Returns (team_map, issues)."""
    issues = []
    usable = []
    for r in rows:
        if not r["CA_raw"]:
            continue
        if r["TL"].upper() in BAD:
            issues.append(("Skipped - no TL", r["CA_raw"], f"row {r['Row']}: TL is empty"))
            continue
        if r["SM"].upper() in BAD:
            r["SM"] = ""
        if not r.get("TT"):
            issues.append(("Team Type missing", proper(r["CA_raw"]), f"row {r['Row']}: choose PV or EV in Team Type"))
        usable.append(r)

    # 1) one entry per name (ignoring case / dots / spaces) and vehicle type ('' = both)
    persons = {}
    for r in usable:
        k = (key_of(r["CA_raw"]), r["VT"])
        if k in persons:
            p = persons[k]
            p["variants"].append(r["CA_raw"])
            if (r["TL"], r["SM"]) != (p["TL"], p["SM"]):
                issues.append(("Conflict - same CA, different team", proper(r["CA_raw"]),
                               f"rows {p['Row']} and {r['Row']}: {p['TL']} / {p['SM']} vs {r['TL']} / {r['SM']}"
                               f" - using row {p['Row']}. Put PV / EV in 'Vehicle Type' if both are right."))
        else:
            persons[k] = dict(r, Key=k[0], variants=[r["CA_raw"]], keys={k[0]})
    for p in persons.values():
        if len(set(p["variants"])) > 1:
            issues.append(("Duplicate merged (case)", proper(p["CA_raw"]), " | ".join(p["variants"])))

    # 2) spelling variants inside the list (same TL + SM + vehicle type, strict match) -> one entry
    merged = {}
    for k, p in persons.items():
        twin = next((m for m in merged.values() if (m["TL"], m["SM"], m["VT"]) == (p["TL"], p["SM"], p["VT"])
                     and same_person(p["Key"], m["Key"])), None)
        if twin:
            twin["keys"].add(p["Key"])
            twin["variants"] += p["variants"]
            continue                                   # e.g. Kavya Iyer / Kavya Iyyer: both spellings, same team
        merged[k] = p
    entries = list(merged.values())

    def find(sk, vt):
        """Best list entry for a Siebel name + vehicle type: exact name first, then strict spelling."""
        exact = [e for e in entries if sk in e["keys"]]
        pool, how = exact, "Same as Siebel"
        if not pool:
            scored = [(max(same_person(sk, k) for k in e["keys"]), e) for e in entries]
            best = max((s for s, _ in scored), default=0)
            pool = [e for s, e in scored if s > 0 and s == best]
            how = "Spelling differs from Siebel"
        for want in (vt, ""):                          # an entry for this vehicle type wins over 'both'
            hits = [e for e in pool if e["VT"] == want]
            if hits:
                return hits[0], how
        return None, None

    # 3) every Siebel name + vehicle type -> TL / SM
    out, used = [], set()
    for name, vt in sorted({(clean_text(n), v) for n, v in siebel_pairs if clean_text(n)}):
        sk = key_of(name)
        base = {"CA_Key": f"{name}|{vt}", "Siebel_Name": name, "Vehicle_Type": vt}
        if PLACEHOLDER.match(sk):
            label = "Seller not found (SSI)" if sk.startswith("SELLER") else "No advisor (Unspecified)"
            out.append(dict(base, CA=label, TL=NOT_IN_LIST, SM=NOT_IN_LIST, Team_Type=NOT_IN_LIST,
                            Match="Placeholder in Siebel", List_Row=None))
            continue
        hit, how = find(sk, vt)
        if hit is None:
            out.append(dict(base, CA=proper(name), TL=NOT_IN_LIST, SM=NOT_IN_LIST, Team_Type=NOT_IN_LIST,
                            Match="In Siebel, not in team list", List_Row=None))
            issues.append(("In Siebel, not in team list", proper(name), f"{vt} - add to Team_Mapping.xlsx if this CA still works"))
            continue
        used.add(id(hit))
        if how != "Same as Siebel":
            issues.append((how, proper(name), f"Siebel '{name}' ({vt}) = list '{hit['CA_raw']}' ({hit['TL']})"))
        out.append(dict(base, CA=proper(hit["CA_raw"]), TL=proper(hit["TL"]), SM=proper(hit["SM"]) or NOT_IN_LIST,
                        Team_Type=hit["TT"] or "Not set", Match=how, List_Row=hit["Row"]))

    # 4) people in the list with no Siebel records yet (new joiners) - kept so they show in slicers
    for e in entries:
        if id(e) not in used:
            vt = e["VT"] or "PV"
            out.append({"CA_Key": f"{key_of(e['CA_raw'])}|{vt}", "Siebel_Name": key_of(e["CA_raw"]), "Vehicle_Type": vt,
                        "CA": proper(e["CA_raw"]), "TL": proper(e["TL"]), "SM": proper(e["SM"]) or NOT_IN_LIST,
                        "Team_Type": e["TT"] or "Not set", "Match": "In list, no Siebel records yet", "List_Row": e["Row"]})

    tm = pd.DataFrame(out)
    tm = tm[~tm["CA_Key"].str.upper().duplicated()].sort_values(["SM", "TL", "CA", "Vehicle_Type"]).reset_index(drop=True)
    tm["Loaded_On"] = pd.Timestamp(today)
    iss = pd.DataFrame(issues, columns=["Issue", "CA", "Detail"]).drop_duplicates()
    iss.insert(0, "Loaded_On", pd.Timestamp(today))
    return tm, iss


def update_history(history, tm, today):
    """One open row (Valid_To empty) per CA_Key; a team change closes it and opens a new one."""
    today = pd.Timestamp(today).normalize()
    cols = ["CA_Key", "CA", "Vehicle_Type", "TL", "SM", "Team_Type", "Valid_From", "Valid_To"]
    h = history.copy() if history is not None and len(history) else pd.DataFrame(columns=cols)
    h = h.reindex(columns=cols)
    for c in ("Valid_From", "Valid_To"):
        h[c] = pd.to_datetime(h[c], errors="coerce")
    current = {str(r["CA_Key"]).upper(): i for i, r in h[h["Valid_To"].isna()].iterrows()}
    new_rows = []
    for _, r in tm.iterrows():
        k = r["CA_Key"].upper()
        if k in current:
            i = current[k]
            if (h.at[i, "TL"], h.at[i, "SM"]) == (r["TL"], r["SM"]):
                continue
            h.at[i, "Valid_To"] = today - pd.Timedelta(days=1)
        row = {c: r.get(c) for c in cols[:6]}
        row.update(Valid_From=today, Valid_To=pd.NaT)
        new_rows.append(row)
    if new_rows:
        h = pd.concat([h, pd.DataFrame(new_rows)], ignore_index=True)
    return h[cols]


# --------------------------------------------------------------------------- #
def siebel_pairs_from_db(db_dir):
    """(CA name, PV/EV) from bookings, open pipeline and SSI (SSI: the actual seller of each survey)."""
    pairs = set()
    for fname, sheet, col in (("master_bookings.xlsx", None, "Employee Full Name"),
                              ("open_pipeline.xlsx", "Open_Now", "Employee Full Name"),
                              ("ssi.xlsx", "SSI_Surveys", "CA")):
        p = os.path.join(db_dir, fname)
        if not os.path.exists(p):
            continue
        with pd.ExcelFile(p) as x:
            sh = sheet or x.sheet_names[0]
            if sh not in x.sheet_names:
                continue
            df = x.parse(sh, usecols=lambda c: c in (col, "Vehicle_Type"))
        if {col, "Vehicle_Type"} <= set(df.columns):
            pairs.update((n, v) for n, v in zip(df[col], df["Vehicle_Type"])
                         if isinstance(n, str) and n.strip() and v in VTS)
            if fname == "ssi.xlsx":                 # surveys with no seller found still need a team row
                pairs.update((SELLER_NOT_FOUND, v) for n, v in zip(df[col], df["Vehicle_Type"])
                             if not (isinstance(n, str) and n.strip()) and v in VTS)
    return pairs


def refresh(cfg_root, cfg, db_dir, today=None, log=None):
    """Called once a day by the bot. Never raises: returns a dict for the health mail."""
    today = pd.Timestamp(today or datetime.now()).normalize()
    src = os.path.join(cfg_root, cfg.get("team_map", {}).get("source_file", SOURCE_FILE))
    if not os.path.exists(src):
        return {"ok": False, "note": f"team list not found: {src}"}
    try:
        tm, iss = build(read_source(src), siebel_pairs_from_db(db_dir), today)
        path = os.path.join(db_dir, OUTPUT_FILE)
        history = None
        if os.path.exists(path):
            with pd.ExcelFile(path) as x:
                if "Team_History" in x.sheet_names:
                    history = x.parse("Team_History")
        hist = update_history(history, tm[tm["TL"] != NOT_IN_LIST], today)
        tmp = temp_name(path)
        daily_backup(path)
        with pd.ExcelWriter(tmp, engine="openpyxl") as w:
            tm.to_excel(w, sheet_name="Team_Map", index=False)
            hist.to_excel(w, sheet_name="Team_History", index=False)
            iss.to_excel(w, sheet_name="Map_Issues", index=False)
        verify_xlsx(tmp)
        os.replace(tmp, path)
    except PermissionError:
        msg = f"{OUTPUT_FILE} or Team_Mapping.xlsx is open in Excel - close it; team list not updated today"
        if log:
            log.warning("Team map: %s", msg)
        return {"ok": False, "note": msg}
    except Exception as exc:
        if log:
            log.exception("Team map failed")
        return {"ok": False, "note": f"{type(exc).__name__}: {exc}"}

    sieb = tm[~tm["Match"].isin(["In list, no Siebel records yet", "Placeholder in Siebel"])]
    unmapped = sorted(set(tm.loc[tm["Match"] == "In Siebel, not in team list", "CA"]))
    res = {"ok": True, "cas": int(tm.loc[tm["TL"] != NOT_IN_LIST, "CA"].nunique()),
           "tls": int(tm.loc[tm["TL"] != NOT_IN_LIST, "TL"].nunique()),
           "sms": int(tm.loc[tm["SM"] != NOT_IN_LIST, "SM"].nunique()),
           "siebel_mapped": int((sieb["TL"] != NOT_IN_LIST).sum()), "siebel_total": int(len(sieb)),
           "unmapped": unmapped, "conflicts": int(iss["Issue"].str.startswith("Conflict").sum())}
    if log:
        log.info("Team map: %d CAs, %d TLs, %d SMs; %d of %d Siebel CA logins mapped; not in list: %s",
                 res["cas"], res["tls"], res["sms"], res["siebel_mapped"], res["siebel_total"], ", ".join(unmapped) or "none")
    return res


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", default=os.path.join(ROOT, "config", "config.json"))
    a = ap.parse_args()
    with open(a.config, encoding="utf-8") as fh:
        cfg = json.load(fh)
    res = refresh(ROOT, cfg, os.path.join(ROOT, cfg["paths"]["database"]))
    for k, v in res.items():
        print(f"{k:14}: {v}")
    return 0 if res.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
