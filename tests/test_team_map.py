import team_map as tm


def test_strict_name_matching():
    assert tm.same_person("KAVYA IYER", "KAVYA IYYER") > 0          # spelling slip, same person
    assert tm.same_person("ROHAN MEHTA", "ROHAN VERMA") == 0         # different people
    assert tm.same_person("ARJUN RAO", "ARJUN KUMAR RAO") > 0        # middle name left out
    assert tm.same_person("RAVI", "RAVI KUMAR") == 0                 # one-word names never matched
    assert tm.key_of("  rOHan.  mehta ") == "ROHAN MEHTA"


def row(i, ca, tl, sm, tt="PV", vt=""):
    return {"Row": i, "CA_raw": ca, "TL": tl, "SM": sm, "VT": vt, "TT": tt, "Note": ""}


def test_build_maps_crm_names_to_teams():
    rows = [row(2, "Kavya Iyer", "Gaurav Sethi", "Rajesh Iyengar"),
            row(3, "Harsh Vora", "Farhan Qureshi", "Imran Sheikh", "EV", "EV"),
            row(4, "Harsh Vora", "Gaurav Sethi", "Rajesh Iyengar", "EV", "PV"),
            row(5, "Sanjay Puri", "", "", "PV"),
            row(6, "Kabir Anand", "Anjali Mathur", "Rajesh Iyengar")]
    crm = {("KAVYA IYYER", "PV"), ("HARSH VORA", "EV"), ("HARSH VORA", "PV"), ("LAKSHMI NARAYAN", "PV"),
           ("UNSPECIFIED", "PV"), ("SELLER NOT FOUND", "PV")}
    m, issues = tm.build(rows, crm, "2026-10-12")
    t = m.set_index("CA_Key")
    assert t.loc["KAVYA IYYER|PV", "TL"] == "Gaurav Sethi"
    assert t.loc["KAVYA IYYER|PV", "Match"] == "Spelling differs from Siebel"
    assert t.loc["HARSH VORA|EV", "TL"] == "Farhan Qureshi"            # split login: each login its own team
    assert t.loc["HARSH VORA|PV", "TL"] == "Gaurav Sethi"
    assert t.loc["LAKSHMI NARAYAN|PV", "TL"] == tm.NOT_IN_LIST
    assert t.loc["UNSPECIFIED|PV", "CA"] == "No advisor (Unspecified)"
    assert t.loc["SELLER NOT FOUND|PV", "CA"] == "Seller not found (SSI)"
    assert "In list, no Siebel records yet" in set(m["Match"])            # new joiner kept for the slicers
    assert {"Skipped - no TL", "In Siebel, not in team list"} <= set(issues["Issue"])
