from openpyxl import Workbook

from handlers import ssi

BASE = ["Assessment Date", "Assessment Status", "CSR#", "Region", "Dealer Code", "Dealer Name", "Division", "PPL", "PL",
        "Date", "Month", "Chassis No", "Registration No", "Sale Date", "Delivered to Customer Date",
        "Number of Days taken for Delivered to Customer", "Number of days taken to be contacted", "Vehicle Usage",
        "Account Name", "First Name", "Last Name", "Phone No-Cell", "Phone No-Off", "Phone No-Res", "CA Name",
        "Registration Type", "Voice Of Customer (New)", "Total Score", "Box Category"]


def write(path, markers, questions, rows):
    wb = Workbook()
    ws = wb.active
    ws.append(["EV - SSI Report (YTD)"])
    ws.append([])
    ws.append([None] * len(BASE) + markers)
    ws.append(BASE + questions)
    for r in rows:
        ws.append(r)
    wb.save(path)


def base_row(csr):
    r = [None] * len(BASE)
    r[BASE.index("CSR#")], r[BASE.index("Assessment Status")], r[BASE.index("Total Score")] = csr, "Done", 900
    return r


def test_score_and_answer_columns_are_paired_by_question_text(tmp_path):
    """Two questionnaires side by side, columns interleaved: Score A, Score B, Value A, Value B."""
    a, b = "Time taken to complete the delivery process", "Was the vehicle delivered on time?"
    p = tmp_path / "ssi.xlsx"
    write(p, ["Attribute Score", "Attribute Score", "Attribute Value", "Attribute Value"], [a, b, a, b],
          [base_row("CSR-1") + [-100, 0, "More than 60 mins", "Yes"]])
    _base, data, questions = ssi.read_ssi(str(p))
    pairs = {q["q"]: (q["score"], q["value"]) for q in questions}
    assert pairs[a] == (len(BASE), len(BASE) + 2)
    assert pairs[b] == (len(BASE) + 1, len(BASE) + 3)


def test_typo_in_one_questionnaire_is_the_same_question():
    groups = {"TIME TO COMPLETE DELIVERY PROCESS": {}}
    assert ssi._same_question("TIME TO COMPLETE DELIVERY PROCES", groups) == "TIME TO COMPLETE DELIVERY PROCESS"
    assert ssi._same_question("TIME TO COMPLETE THE HANDOVER", groups) == "TIME TO COMPLETE THE HANDOVER"


def test_question_types():
    assert ssi._q_type(["YES", "NO"]) == "Yes / No"
    assert ssi._q_type(["10", "9", "`8"]) == "Rating 0-10"
    assert ssi._q_type(["UP TO 30 MINS", "45 TO 60 MINS"]) == "Time band"
    assert ssi._q_type(["", ""]) is None


def test_anything_pending_no_is_not_a_failure(tmp_path):
    q1, q2 = "Is anything pending from the dealership after the delivery?", "Was the vehicle delivered on time?"
    p = tmp_path / "ssi.xlsx"
    write(p, ["Attribute Score", "Attribute Value"] * 2, [q1, q1, q2, q2],
          [base_row("CSR-1") + [0, "No", -60, "No"], base_row("CSR-2") + [0, "Yes", 0, "Yes"]])
    base, data, questions = ssi.read_ssi(str(p))
    surveys = base.assign(Vehicle_Type="EV", Survey_Month=None)
    ans = ssi.build_answers(data.reset_index(drop=True), questions, surveys.reset_index(drop=True))
    info = ans[ans["Question"] == q1]
    assert set(info["Q_Type"]) == {"Information"} and info["Is_No"].sum() == 0
    yn = ans[ans["Question"] == q2]
    assert yn["Is_No"].sum() == 1 and yn["Has_Penalty"].sum() == 1


def test_survey_with_no_seller_found_gets_a_real_team_key(tmp_path):
    """Regression: the key used to become the text 'nan|PV', which matched no team (a (Blank) row in Power BI)."""
    import pandas as pd
    s = pd.DataFrame({"Chassis No": ["DMX1", "DMX2"], "CA_ID": ["XX9_3999001", "YY9_5999001"],
                      "Vehicle_Type": ["PV", "EV"]})
    out = ssi.credit(s, str(tmp_path / "data" / "database"), [])
    assert list(out["CA_Source"]) == ["Not found", "Not found"]
    assert list(out["CA_Key"]) == ["SELLER NOT FOUND|PV", "SELLER NOT FOUND|EV"]
