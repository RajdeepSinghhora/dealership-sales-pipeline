r"""
ssi.py  -  OEM SSI (Sales Satisfaction Index) survey files
============================================================
Files (daily mail):  "pv SSI Report (YTD)"  and  "EV - SSI Report (YTD)"
Layout:  row 1 title, row 3 "Attribute Score" / "Attribute Value" markers, row 4 headers, data from row 5.
         Columns A..AC = survey details (CSR#, Month, Chassis No, CA Name (= login ID), Total Score,
         Box Category ...). From AD onwards every question appears twice: its score and its answer.

Output: data/database/ssi.xlsx
  SSI_Surveys   one row per survey (CSR#), kept across the year (YTD file is upserted, never trimmed)
  SSI_Answers   one row per survey x question: answer, points, rating (0-10), "No" flag, penalty
  SSI_Issues    surveys whose seller could not be found, shared IDs, surveys not completed yet

Who gets the credit (CA -> TL -> SM):
  1. Actual seller by chassis from the GM sales register (published CSV, config ssi.actual_ca_csv_url)
  2. if the chassis is not there: the CA ID list (data/reference/CA_ID_List.xlsx)
  3. if the ID is not there either: the Siebel owner of the login (from open_pipeline.xlsx)
  CA_Source says which rule was used for every survey.

Score bands (config ssi.low_score / ssi.target_score, default 930 / 980):
  >= target -> "Target met",  >= low -> "Below target",  < low -> "Low score"
"""
import difflib
import json
import os
import re
import urllib.request
from datetime import datetime

import pandas as pd

from . import FileLockedError
from .open_pipeline import _save

OUTPUT_FILE = "ssi.xlsx"
CHASSIS_CACHE = os.path.join("data", "reference", "CHASSIS_CA_last_good.csv")
ID_LIST = os.path.join("data", "reference", "CA_ID_List.xlsx")
NOT_FOUND = "SELLER NOT FOUND"     # team key of a survey whose seller is in no list (team_map.py shows it as such)
JUNK_NAME = re.compile(r"^(cancel+ed|cancel|td|direct|na|n/?a|none|nan|)$", re.I)
SURVEY_COLS = ["CSR#", "Vehicle_Type", "Survey_Month", "Assessment Date", "Assessment Status", "Is_Done", "Branch",
               "Division", "PPL", "PL", "Chassis No", "Registration No", "Sale Date", "Delivered to Customer Date",
               "Days_Sale_To_Delivery", "Days_Delivery_To_Call", "Vehicle Usage", "Customer", "Mobile_Masked",
               "CA_ID", "CA", "CA_Key", "CA_Source", "Total Score", "Box Category", "Score_Band", "Score_Band_Order",
               "Lowest_Question", "Lowest_Rating", "No_Answers", "Voice Of Customer (New)", "Load_Date"]
ANSWER_COLS = ["CSR#", "Vehicle_Type", "Survey_Month", "Q_No", "Question", "Q_Type", "Answer", "Points",
               "Rating", "Is_No", "Has_Penalty"]


# --------------------------------------------------------------------------- #
def _root(db_dir):
    return os.path.abspath(os.path.join(db_dir, "..", ".."))


def _settings(db_dir):
    """The 'ssi' block of the config the bot was started with (BOOKINGS_CONFIG), else config/config.json."""
    path = os.environ.get("BOOKINGS_CONFIG") or os.path.join(_root(db_dir), "config", "config.json")
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh).get("ssi", {})
    except Exception:
        return {}


def _key(name):
    return re.sub(r"\s+", " ", re.sub(r"[^A-Za-z ]", " ", str(name))).strip().upper()


def _text(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def _branch(division):
    parts = _text(division).split("-")
    return parts[2].strip() if len(parts) > 2 and parts[2].strip() else None


def _month(v, fallback):
    m = re.match(r"^\s*(\d{4})\s*/\s*(\d{1,2})\s*$", _text(v))
    if m:
        return pd.Timestamp(int(m.group(1)), int(m.group(2)), 1)
    fb = pd.to_datetime(fallback, errors="coerce")
    return pd.Timestamp(fb.year, fb.month, 1) if pd.notna(fb) else pd.NaT


# --------------------------------------------------------------------------- #
def read_ssi(path):
    """-> (surveys DataFrame with the original columns, list of question dicts)"""
    raw = pd.read_excel(path, header=None, dtype=object)
    hdr = next((i for i in range(min(12, len(raw))) if "CSR#" in [_text(v) for v in raw.iloc[i]]), None)
    if hdr is None:
        raise ValueError("this is not an SSI report - no 'CSR#' header found in the first rows")
    heads = [_text(v) for v in raw.iloc[hdr]]
    kinds = [_text(v).lower() for v in raw.iloc[hdr - 1]] if hdr > 0 else [""] * len(heads)
    try:
        last_base = heads.index("Box Category")
    except ValueError:
        raise ValueError("SSI report has no 'Box Category' column")
    data = raw.iloc[hdr + 1:].reset_index(drop=True)
    data = data[data[heads.index("CSR#")].notna()]
    base = data.iloc[:, :last_base + 1].copy()
    base.columns = heads[:last_base + 1]

    # Each question has an "Attribute Score" column (points) and an "Attribute Value" column (answer).
    # They are matched by the question TEXT, not by position: since Oct 2026 the OEM's YTD report holds the
    # old and the new questionnaire side by side and the score / value columns are interleaved
    # (Score A, Score B, Value A, Value B). A question that is in both questionnaires appears twice:
    # the 1st score column goes with the 1st value column, the 2nd with the 2nd.
    groups, names = {}, {}
    for c in range(last_base + 1, len(heads)):
        q = heads[c].replace("?s ", "'s ")
        if not q:
            continue
        key = _same_question(_key(q), groups)
        g = groups.setdefault(key, {"score": [], "value": []})
        g["score" if "score" in kinds[c] else "value"].append(c)       # no marker -> treated as the answer
        if len(q) > len(names.get(key, "")):
            names[key] = q                                              # 'Proces' / 'Process' -> the longer one
    questions = []
    for key, g in groups.items():
        for k in range(max(len(g["score"]), len(g["value"]))):
            questions.append({"q": names[key],
                              "score": g["score"][k] if k < len(g["score"]) else None,
                              "value": g["value"][k] if k < len(g["value"]) else None})
    return base, data, questions


def _same_question(key, groups, min_ratio=0.97):
    """Same question with a typo in one questionnaire ('...Delivery Proces' vs '...Delivery Process')."""
    if key in groups:
        return key
    for other in groups:
        if abs(len(other) - len(key)) <= 2 and difflib.SequenceMatcher(None, key, other).ratio() >= min_ratio:
            return other
    return key


# Yes/No questions where "No" is NOT a failure (information only) - never counted as a "No" answer
INFO_QUESTION = re.compile(r"anything pending|first four wheeler|first free service|personal or commercial", re.I)


def _q_type(answers):
    vals = {a for a in answers if a}
    if not vals:
        return None
    if vals <= {"YES", "NO"}:
        return "Yes / No"
    if all(re.fullmatch(r"\W*\d{1,2}", v) for v in vals):
        return "Rating 0-10"
    if any("MIN" in v for v in vals):
        return "Time band"
    return "Information"


def build_answers(data, questions, surveys):
    rows = []
    for no, q in enumerate(questions, start=1):
        vals = [_text(v) for v in (data[q["value"]] if q["value"] is not None else [None] * len(data))]
        pts = list(data[q["score"]]) if q["score"] is not None else [None] * len(data)
        clean = [v if v.lower() != "none" else "" for v in vals]
        qt = _q_type([v.upper() for v in clean])
        if qt is None:
            continue                                   # question not asked in this file yet (all blank)
        if qt == "Yes / No" and INFO_QUESTION.search(q["q"]):
            qt = "Information"                         # e.g. "Anything pending ...?" - "No" is the good answer
        for csr, vt, month, ans, p in zip(surveys["CSR#"], surveys["Vehicle_Type"], surveys["Survey_Month"], clean, pts):
            if not ans and (p is None or pd.isna(p)):
                continue
            p = pd.to_numeric(p, errors="coerce")
            rating = None
            if qt == "Rating 0-10":
                m = re.search(r"\d{1,2}", ans)
                rating = int(m.group()) if m else None
            rows.append({"CSR#": csr, "Vehicle_Type": vt, "Survey_Month": month, "Q_No": no, "Question": q["q"],
                         "Q_Type": qt, "Answer": ans or None, "Points": p, "Rating": rating,
                         "Is_No": 1 if ans.upper() == "NO" and qt == "Yes / No" else 0,
                         "Has_Penalty": 1 if pd.notna(p) and p < 0 else 0})
    return pd.DataFrame(rows, columns=ANSWER_COLS)


# --------------------------------------------------------------------------- #
def _chassis_map(db_dir, log_notes):
    """Chassis No -> Actual CA, from the published GM register link (cached once a day)."""
    root, s = _root(db_dir), _settings(db_dir)
    cache = os.path.join(root, s.get("chassis_cache", CHASSIS_CACHE))
    url = _text(s.get("actual_ca_csv_url"))
    fresh = os.path.exists(cache) and datetime.fromtimestamp(os.path.getmtime(cache)).date() == datetime.now().date()
    if url and not fresh:
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (dealership bookings bot)"})
            with urllib.request.urlopen(req, timeout=60) as resp:
                text = resp.read().decode("utf-8-sig", errors="replace")
            if "<html" in text[:500].lower():
                raise ValueError("the link returned a web page, not CSV - publish the tab as CSV")
            os.makedirs(os.path.dirname(cache), exist_ok=True)
            with open(cache, "w", encoding="utf-8", newline="") as fh:
                fh.write(text)
        except Exception as exc:
            log_notes.append(f"actual-seller link not reachable ({type(exc).__name__}) - used the last good copy")
    elif not url:
        log_notes.append("actual-seller link (ssi.actual_ca_csv_url) not set - "
                         + (f"used the saved copy {os.path.basename(cache)}" if os.path.exists(cache)
                            else "credited by CA ID list"))
    if not os.path.exists(cache):
        return {}
    df = pd.read_csv(cache, dtype=str, keep_default_na=False)
    cols = {re.sub(r"[^a-z]", "", c.lower()): c for c in df.columns}
    c_ch, c_ca = cols.get("chassisno") or df.columns[0], cols.get("actualca") or df.columns[1]
    out = {}
    for ch, ca in zip(df[c_ch], df[c_ca]):
        ch, ca = _text(ch).upper(), _text(ca)
        if ch and ca and not JUNK_NAME.match(ca):
            out[ch] = ca
    return out


def _id_list(db_dir):
    p = os.path.join(_root(db_dir), _settings(db_dir).get("id_list", ID_LIST))
    if not os.path.exists(p):
        return {}
    df = pd.read_excel(p, dtype=str).fillna("")
    cols = {re.sub(r"[^a-z]", "", c.lower()): c for c in df.columns}
    c_id = cols.get("caid") or cols.get("id") or df.columns[0]
    c_nm = cols.get("caname") or cols.get("name") or df.columns[1]
    out = {}
    for i, n in zip(df[c_id], df[c_nm]):
        i, n = _text(i).upper(), _text(n)
        if i and n and not JUNK_NAME.match(n):
            out.setdefault(i, [])
            if _key(n) not in [_key(x) for x in out[i]]:
                out[i].append(n)
    return out


def _login_owners(db_dir):
    p = os.path.join(db_dir, "open_pipeline.xlsx")
    if not os.path.exists(p):
        return {}
    with pd.ExcelFile(p) as x:
        if "Open_Now" not in x.sheet_names:
            return {}
        df = x.parse("Open_Now", usecols=lambda c: c in ("Employee Login", "Employee Full Name"), dtype=str)
    if not {"Employee Login", "Employee Full Name"} <= set(df.columns):
        return {}
    df = df.dropna()
    return df.groupby(df["Employee Login"].str.strip().str.upper())["Employee Full Name"].agg(
        lambda s: s.value_counts().index[0]).to_dict()


def credit(surveys, db_dir, notes):
    chassis, ids, owners = _chassis_map(db_dir, notes), _id_list(db_dir), _login_owners(db_dir)
    ca, src = [], []
    for ch, cid in zip(surveys["Chassis No"], surveys["CA_ID"]):
        ch, cid = _text(ch).upper(), _text(cid).upper()
        if ch in chassis:
            ca.append(chassis[ch]); src.append("Actual seller (GM register)")
        elif cid in ids:
            names = ids[cid]
            if len(names) == 1:
                ca.append(names[0]); src.append("CA ID list")
            else:
                own = owners.get(cid)
                pick = next((n for n in names if own and _key(n) == _key(own)), names[0])
                ca.append(pick); src.append("CA ID list (shared ID - check)")
        elif cid in owners:
            ca.append(owners[cid]); src.append("Siebel login owner")
        else:
            ca.append(None); src.append("Not found")
    names = [_key(c) if c else None for c in ca]            # a plain list: a pandas column would turn None into NaN
    surveys["CA"] = names
    surveys["CA_Source"] = src
    surveys["CA_Key"] = [f"{c or NOT_FOUND}|{v}" for c, v in zip(names, surveys["Vehicle_Type"])]
    return surveys


def clean_surveys(base, group, load_date, db_dir, notes):
    s = _settings(db_dir)
    low, target = float(s.get("low_score", 930)), float(s.get("target_score", 980))
    d = pd.DataFrame()
    d["CSR#"] = base["CSR#"].map(_text)
    d["Vehicle_Type"] = group
    d["Survey_Month"] = [_month(m, a) for m, a in zip(base.get("Month"), base.get("Assessment Date"))]
    for c in ("Assessment Date", "Sale Date", "Delivered to Customer Date"):
        d[c] = pd.to_datetime(base.get(c), errors="coerce")
    d["Assessment Status"] = base.get("Assessment Status").map(_text)
    d["Total Score"] = pd.to_numeric(base.get("Total Score"), errors="coerce")
    d["Is_Done"] = ((d["Assessment Status"].str.lower() == "done") & d["Total Score"].notna()).astype(int)
    d["Division"] = base.get("Division").map(_text)
    d["Branch"] = d["Division"].map(_branch)
    for c in ("PPL", "PL", "Chassis No", "Registration No", "Vehicle Usage", "Box Category", "Voice Of Customer (New)"):
        d[c] = base[c].map(lambda v: _text(v) or None) if c in base else None
    d["Days_Sale_To_Delivery"] = pd.to_numeric(base.get("Number of Days taken for Delivered to Customer"), errors="coerce")
    d["Days_Delivery_To_Call"] = pd.to_numeric(base.get("Number of days taken to be contacted"), errors="coerce")
    d["Customer"] = [(_text(a) + " " + _text(b)).strip() or None for a, b in zip(base.get("First Name"), base.get("Last Name"))]
    mob = base.get("Phone No-Cell").map(lambda v: re.sub(r"\D", "", _text(v)))
    d["Mobile_Masked"] = mob.map(lambda v: "XXXXXX" + v[-4:] if len(v) >= 4 else None)
    d["CA_ID"] = base.get("CA Name").map(_text).str.upper()
    band = []
    for done, sc in zip(d["Is_Done"], d["Total Score"]):
        if not done:
            band.append(("Not completed", 9))
        elif sc >= target:
            band.append((f"Target met ({target:.0f}+)", 1))
        elif sc >= low:
            band.append((f"Below target ({low:.0f}-{target - 1:.0f})", 2))
        else:
            band.append((f"Low score (under {low:.0f})", 3))
    d["Score_Band"], d["Score_Band_Order"] = [b[0] for b in band], [b[1] for b in band]
    d["Load_Date"] = pd.Timestamp(load_date)
    return credit(d, db_dir, notes)


def _worst(answers):
    """Lowest rated question and number of 'No' answers per survey (for the CXM call list)."""
    r = answers[answers["Q_Type"] == "Rating 0-10"].dropna(subset=["Rating"])
    low = r[r["Rating"] < 10].sort_values(["CSR#", "Rating", "Q_No"]).drop_duplicates("CSR#")   # all 10s -> nothing weak
    nos = answers.groupby("CSR#")["Is_No"].sum()
    return low.set_index("CSR#")[["Question", "Rating"]], nos


def _issues(surveys):
    out = []
    for _, r in surveys.iterrows():
        if r["CA_Source"] == "Not found":
            out.append(("Seller not found", r, f"ID {r['CA_ID']} is not in the GM register, the ID list or Siebel - add it to CA_ID_List.xlsx"))
        elif r["CA_Source"].startswith("CA ID list (shared"):
            out.append(("Shared ID - seller guessed", r, f"ID {r['CA_ID']} has several CAs in CA_ID_List.xlsx; chassis not in the GM register"))
        if not r["Is_Done"]:
            out.append(("Survey not completed", r, f"status '{r['Assessment Status']}' - not counted in the score"))
    return pd.DataFrame([{"Issue": i, "CSR#": r["CSR#"], "Vehicle_Type": r["Vehicle_Type"], "Survey_Month": r["Survey_Month"],
                          "CA_ID": r["CA_ID"], "CA": r["CA"], "Chassis No": r["Chassis No"], "Detail": t}
                         for i, r, t in out],
                        columns=["Issue", "CSR#", "Vehicle_Type", "Survey_Month", "CA_ID", "CA", "Chassis No", "Detail"])


def _load(path):
    if not os.path.exists(path):
        return pd.DataFrame(columns=SURVEY_COLS), pd.DataFrame(columns=ANSWER_COLS)
    with pd.ExcelFile(path) as x:
        s = x.parse("SSI_Surveys") if "SSI_Surveys" in x.sheet_names else pd.DataFrame(columns=SURVEY_COLS)
        a = x.parse("SSI_Answers") if "SSI_Answers" in x.sheet_names else pd.DataFrame(columns=ANSWER_COLS)
    return s, a


# --------------------------------------------------------------------------- #
def process(path, group, load_date, db_dir):
    notes = []
    base, data, questions = read_ssi(path)
    orphans = [q["q"] for q in questions if q["score"] is not None and q["value"] is None]
    if orphans:                                        # the OEM changed the layout again -> say so in the health mail
        notes.append(f"LAYOUT CHECK: {len(orphans)} question(s) have a points column but no answer column "
                     f"(e.g. '{orphans[0][:50]}') - penalty points may be missing, send this file to MIS")
    surveys = clean_surveys(base, group, load_date, db_dir, notes)
    surveys["_i"] = range(len(surveys))
    answers = build_answers(data.reset_index(drop=True), questions, surveys.reset_index(drop=True))
    low, nos = _worst(answers)
    surveys["Lowest_Question"] = surveys["CSR#"].map(low["Question"]) if len(low) else None
    surveys["Lowest_Rating"] = surveys["CSR#"].map(low["Rating"]) if len(low) else None
    surveys["No_Answers"] = surveys["CSR#"].map(nos).fillna(0).astype(int) if len(nos) else 0
    surveys = surveys.drop(columns="_i").reindex(columns=SURVEY_COLS)

    path_out = os.path.join(db_dir, OUTPUT_FILE)
    old_s, old_a = _load(path_out)
    old_ids = set(old_s["CSR#"].astype(str))
    new_ids = set(surveys["CSR#"]) - old_ids
    keep_s = old_s[~old_s["CSR#"].astype(str).isin(surveys["CSR#"])]
    keep_a = old_a[~old_a["CSR#"].astype(str).isin(surveys["CSR#"])]
    all_s = pd.concat([keep_s, surveys], ignore_index=True).sort_values(["Survey_Month", "Vehicle_Type", "CSR#"])
    all_a = pd.concat([keep_a, answers], ignore_index=True)
    _save(path_out, [("SSI_Surveys", all_s), ("SSI_Answers", all_a), ("SSI_Issues", _issues(all_s))])

    done = surveys[surveys["Is_Done"] == 1]
    by = surveys["CA_Source"].value_counts().to_dict()
    notes.append(f"{len(done)} completed surveys, average {done['Total Score'].mean():.0f}" if len(done) else "no completed surveys")
    notes.append("credited: " + ", ".join(f"{v} {k.lower()}" for k, v in by.items()))
    return {"rows_read": len(surveys), "new": len(new_ids), "updated": len(surveys) - len(new_ids),
            "unchanged": 0, "notes": notes}
