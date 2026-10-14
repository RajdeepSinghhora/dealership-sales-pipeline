"""
health_email.py
---------------
Builds the HTML health-status mail (green = all good, red = problem) and sends it.
All styling is inline, because Gmail strips <style> blocks and the colours would vanish.
"""
import os
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from html import escape as esc

GREEN, RED, AMBER, GREY = "#1e8e3e", "#d93025", "#f29900", "#5f6368"
LIGHT = {"GREEN": "#e6f4ea", "RED": "#fce8e6", "AMBER": "#fef7e0"}
FONT = "font-family:Arial,Helvetica,sans-serif;"


def _chip(text, color):
    return (f'<span style="display:inline-block;padding:2px 9px;border-radius:10px;'
            f'background:{color};color:#ffffff;font-size:12px;font-weight:bold;">{esc(str(text))}</span>')


def _th(text, align="left"):
    return (f'<th align="{align}" style="padding:8px 10px;background:#1f3864;color:#ffffff;'
            f'font-size:12px;{FONT}">{esc(str(text))}</th>')


def _td(text, color="#202124", bold=False, align="left", bg=None, raw=False):
    style = f"padding:7px 10px;border-bottom:1px solid #e0e0e0;font-size:13px;color:{color};{FONT}"
    if bold:
        style += "font-weight:bold;"
    if bg:
        style += f"background:{bg};"
    return f'<td align="{align}" style="{style}">{text if raw else esc(str(text))}</td>'


def _section(title):
    return (f'<tr><td style="padding:18px 20px 6px 20px;{FONT}font-size:15px;font-weight:bold;'
            f'color:#1f3864;">{esc(title)}</td></tr>')


def build_subject(report):
    icon = {"GREEN": "\u2705", "RED": "\U0001F534", "AMBER": "\U0001F7E0"}[report["status"]]
    word = {"GREEN": "ALL OK", "RED": "ACTION NEEDED", "AMBER": "WAITING"}[report["status"]]
    return f"{icon} Bookings pipeline {word} - {report['run_date']}"


def build_html(report):
    """
    report keys:
      status: GREEN | RED | AMBER      headline: str      run_date, run_time: str
      groups: [{name, expected:[...], received:[{name, arrived, status, rows, new, updated,
                unchanged, note}], missing:[...]}]
      others: [{name, sender, arrived}]   (optional)
      errors: [str]
      master: {total, pv, ev, latest_yf, stage_counts:{...}}  (optional)
      dq: {warn, info, by_check:[{Severity,Check,n}], repeated_by_advisor:{name:n}} (optional)
      sla: {target, ready_at, met: bool|None} (optional)
      notes: [str] (optional, amber "Good to know" box)
      groups[i].checked: False when the mailbox could not be read (files shown as "not checked")
      master_title: str (optional heading for the database section)
    """
    st = report["status"]
    color = {"GREEN": GREEN, "RED": RED, "AMBER": AMBER}[st]
    rows = []

    # --- banner ---------------------------------------------------------------
    rows.append(
        f'<tr><td style="background:{color};padding:22px 20px;{FONT}">'
        f'<div style="font-size:12px;color:#ffffff;letter-spacing:1px;">BOOKINGS DATA PIPELINE &middot; '
        f'{esc(report["run_date"])} &middot; {esc(report["run_time"])}</div>'
        f'<div style="font-size:24px;font-weight:bold;color:#ffffff;margin-top:6px;">'
        f'{"&#10004; " if st == "GREEN" else "&#9888; "}{esc(report["headline"])}</div></td></tr>')

    # --- errors box (red mails) -----------------------------------------------
    if report.get("errors"):
        items = "".join(f'<li style="margin:4px 0;">{esc(e)}</li>' for e in report["errors"])
        rows.append(
            f'<tr><td style="padding:14px 20px 0 20px;"><div style="background:{LIGHT["RED"]};'
            f'border-left:6px solid {RED};padding:12px 16px;{FONT}font-size:13px;color:#7a1810;">'
            f'<b>What went wrong</b><ul style="margin:8px 0 0 18px;padding:0;">{items}</ul></div></td></tr>')

    # --- notes box (late start, retry plan, ...) -------------------------------
    if report.get("notes"):
        items = "".join(f'<li style="margin:4px 0;">{esc(n)}</li>' for n in report["notes"])
        rows.append(
            f'<tr><td style="padding:10px 20px 0 20px;"><div style="background:{LIGHT["AMBER"]};'
            f'border-left:6px solid {AMBER};padding:12px 16px;{FONT}font-size:13px;color:#5c3c00;">'
            f'<b>Good to know</b><ul style="margin:8px 0 0 18px;padding:0;">{items}</ul></div></td></tr>')

    # --- file check per group -------------------------------------------------
    for g in report["groups"]:
        if g.get("checked") is False:            # mailbox could not be read -> we do not know
            rows.append(_section(f'{g["name"]} mail  –  not checked (Gmail could not be reached)'))
            rows.append(f'<tr><td style="padding:0 20px;{FONT}font-size:13px;color:{GREY};">'
                        f'Expected: {esc(", ".join(g["expected"]))}</td></tr>')
            continue
        n_exp, n_miss = len(g["expected"]), len(g["missing"])
        ok = n_miss == 0 and all(r["status"] != "FAILED" for r in g["received"])
        rows.append(_section(f'{g["name"]} mail  \u2013  {n_exp - n_miss} of {n_exp} expected files received'))
        body = [f'<tr>{_th("File")}{_th("Arrived")}{_th("Rows", "right")}{_th("New", "right")}'
                f'{_th("Updated", "right")}{_th("Unchanged", "right")}{_th("Status")}</tr>']
        for r in g["received"]:
            failed = r["status"] == "FAILED"
            bg = LIGHT["RED"] if failed else None
            chip = _chip("FAILED", RED) if failed else (
                _chip("Loaded", GREEN) if r["status"] == "OK" else _chip(r["status"], GREY))
            note = (f'<div style="font-size:11px;font-weight:normal;color:{RED if failed else GREY};margin-top:3px;">'
                    f'{esc(r["note"])}</div>') if r.get("note") else ""
            body.append(
                "<tr>" + _td(esc(r["name"]) + note, bold=failed, bg=bg, raw=True) + _td(esc(r.get("arrived", "-")) + (f'<div style="font-size:10px;color:{GREY};">{esc(r["sender"])}</div>' if r.get("sender") else ""), bg=bg, raw=True)
                + _td(r.get("rows", "-"), align="right", bg=bg) + _td(r.get("new", "-"), align="right", bg=bg)
                + _td(r.get("updated", "-"), align="right", bg=bg) + _td(r.get("unchanged", "-"), align="right", bg=bg)
                + _td(chip, raw=True, bg=bg) + "</tr>")
        for m in g["missing"]:
            body.append(
                "<tr>" + _td(m, color=RED, bold=True, bg=LIGHT["RED"]) + _td("NOT RECEIVED", color=RED, bold=True, bg=LIGHT["RED"])
                + "".join(_td("-", align="right", bg=LIGHT["RED"]) for _ in range(4))
                + _td(_chip("MISSING", RED), raw=True, bg=LIGHT["RED"]) + "</tr>")
        border = GREEN if ok else RED
        rows.append(f'<tr><td style="padding:0 20px;"><table width="100%" cellspacing="0" cellpadding="0" '
                    f'style="border-top:3px solid {border};border-collapse:collapse;">{"".join(body)}</table></td></tr>')

    # --- other files from the same senders ------------------------------------
    others = report.get("others")
    if others:
        rows.append(_section(f"Other files received today (not PV / EV) \u2013 {len(others)}"))
        body = [f'<tr>{_th("File")}{_th("From")}{_th("Arrived")}</tr>']
        for o in others[:30]:
            body.append("<tr>" + _td(o["name"]) + _td(o.get("sender", "-")) + _td(o.get("arrived", "-")) + "</tr>")
        rows.append(f'<tr><td style="padding:0 20px;"><table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;">{"".join(body)}</table></td></tr>')

    # --- master database ------------------------------------------------------
    m = report.get("master")
    if m:
        rows.append(_section(report.get("master_title") or "Master database (after this run)"))
        cells = "".join(
            f'<td align="center" style="padding:10px;border:1px solid #e0e0e0;{FONT}">'
            f'<div style="font-size:22px;font-weight:bold;color:#1f3864;">{esc(str(v))}</div>'
            f'<div style="font-size:11px;color:{GREY};">{esc(k)}</div></td>'
            for k, v in [("Total bookings", m["total"]), ("PV bookings", m["pv"]),
                         ("EV bookings", m["ev"]), ("Latest YF date", m.get("latest_yf") or "-")])
        rows.append(f'<tr><td style="padding:0 20px;"><table width="100%" cellspacing="0" cellpadding="0"><tr>{cells}</tr></table></td></tr>')
        if m.get("stage_counts"):
            sc = " &nbsp;|&nbsp; ".join(f"<b>{esc(k)}</b>: {v}" for k, v in m["stage_counts"].items())
            rows.append(f'<tr><td style="padding:8px 20px 0 20px;{FONT}font-size:12px;color:{GREY};">Stage mix &ndash; {sc}</td></tr>')

    # --- data quality ---------------------------------------------------------
    d = report.get("dq")
    if d:
        warn_color = AMBER if d["warn"] else GREEN
        rows.append(_section("CRM data quality"))
        rows.append(f'<tr><td style="padding:0 20px;{FONT}font-size:13px;">'
                    f'{_chip(str(d["warn"]) + " warning" + ("" if d["warn"] == 1 else "s"), warn_color)} &nbsp; '
                    f'{_chip(str(d["info"]) + " info", GREY)}</td></tr>')
        if d.get("by_check"):
            body = [f'<tr>{_th("Check")}{_th("Level")}{_th("Count", "right")}</tr>']
            for r in d["by_check"]:
                c = AMBER if r["Severity"] == "WARN" else GREY
                body.append("<tr>" + _td(r["Check"]) + _td(r["Severity"], color=c, bold=True) + _td(r["n"], align="right") + "</tr>")
            rows.append(f'<tr><td style="padding:8px 20px 0 20px;"><table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;">{"".join(body)}</table></td></tr>')
        if d.get("repeated_by_advisor"):
            adv = ", ".join(f"{esc(k)} ({v})" for k, v in d["repeated_by_advisor"].items())
            rows.append(f'<tr><td style="padding:8px 20px 0 20px;{FONT}font-size:12px;color:#202124;">'
                        f'<b>Advisors to coach &ndash; customer name typed twice:</b> {adv}</td></tr>')

    # --- SLA -----------------------------------------------------------------
    s = report.get("sla")
    if s:
        met = s.get("met")
        chip = _chip("SLA MET", GREEN) if met else (_chip("SLA MISSED", RED) if met is False else _chip("SLA PENDING", AMBER))
        rows.append(_section("SLA"))
        rows.append(f'<tr><td style="padding:0 20px;{FONT}font-size:13px;">Target: data ready by <b>{esc(s["target"])}</b>'
                    f' &nbsp;|&nbsp; Ready at: <b>{esc(s.get("ready_at") or "-")}</b> &nbsp; {chip}</td></tr>')

    rows.append(f'<tr><td style="padding:22px 20px;{FONT}font-size:11px;color:{GREY};">'
                f'Automatic message from the Bookings pipeline. Refresh Power BI to see the new data.</td></tr>')

    return (f'<html><head><meta charset="utf-8"></head><body style="margin:0;padding:16px;background:#f1f3f4;">'
            f'<table align="center" width="680" cellspacing="0" cellpadding="0" '
            f'style="background:#ffffff;border:1px solid #dadce0;border-radius:6px;">{"".join(rows)}</table></body></html>')


def send_email(smtp_cfg, subject, html):
    """Send via SMTP (Gmail app password). Returns (ok, message)."""
    pw = os.environ.get(smtp_cfg.get("password_env", "GMAIL_APP_PASSWORD"), "")
    if not pw:
        return False, f"Environment variable {smtp_cfg.get('password_env', 'GMAIL_APP_PASSWORD')} is not set"
    msg = MIMEMultipart("alternative")
    msg["Subject"], msg["From"] = subject, smtp_cfg["user"]
    msg["To"] = ", ".join(smtp_cfg["to"])
    msg.attach(MIMEText("Open this mail in an HTML-capable client to see the pipeline status.", "plain", "utf-8"))
    msg.attach(MIMEText(html, "html", "utf-8"))
    try:
        with smtplib.SMTP_SSL(smtp_cfg.get("host", "smtp.gmail.com"), int(smtp_cfg.get("port", 465)),
                              context=ssl.create_default_context(), timeout=60) as s:
            s.login(smtp_cfg["user"], pw)
            s.sendmail(smtp_cfg["user"], smtp_cfg["to"], msg.as_string())
        return True, "sent"
    except Exception as exc:  # network, auth, ...
        return False, f"{type(exc).__name__}: {exc}"
