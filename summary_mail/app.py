import io
import os
from datetime import datetime, timedelta, date
from typing import List, Optional, Tuple

import numpy as np
import pandas as pd
from flask import Flask, render_template, request, jsonify
from flask_cors import CORS
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

# ---------- SMTP config ----------
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USERNAME = os.getenv("SMTP_USERNAME")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")
EMAIL_FROM = os.getenv("EMAIL_FROM", SMTP_USERNAME)
# ----------------------------------

# ---------- helper: column detection ----------
def find_col(cols: List[str], needles: List[str]) -> Optional[str]:
    for c in cols:
        lc = str(c).lower()
        for n in needles:
            if n in lc:
                return c
    return None

def normalize_schema(df: pd.DataFrame) -> pd.DataFrame:
    cols = list(df.columns)
    # Date
    date_col = find_col(cols, ["date"])
    if not date_col:
        raise ValueError("No column containing 'date' found. Please include a Date column.")
    df = df.rename(columns={date_col: "Date"})

    # Category/Subcategory/Line Item
    cat = find_col(cols, ["category", "cat"])
    sub = find_col(cols, ["sub-category", "subcategory", "sub category", "subcat"])
    line = find_col(cols, ["line item", "line_item", "lineitem", "task", "activity", "li"])

    if cat : df = df.rename(columns={cat: "Category"})
    if sub : df = df.rename(columns={sub: "Subcategory"})
    if line: df = df.rename(columns={line: "Line Item"})

    # Planned/Actual LI
    planned_li = find_col(cols, ["planned li", "planned line", "planned items", "planned count", "planned"])
    actual_li  = find_col(cols, ["actual li", "actual line", "actual items", "actual count", "actual"])

    if planned_li: df = df.rename(columns={planned_li: "Planned LI"})
    if actual_li:  df = df.rename(columns={actual_li: "Actual LI"})

    # Planned/Actual Efforts
    planned_eff = find_col(cols, ["planned effort", "planned efforts", "planned mins", "planned minutes"])
    actual_eff  = find_col(cols, ["actual effort", "actual efforts", "actual mins", "actual minutes"])

    if planned_eff: df = df.rename(columns={planned_eff: "Planned Efforts (mins)"})
    if actual_eff:  df = df.rename(columns={actual_eff: "Actual Efforts (mins)"})

    # Details
    planned_det = find_col(cols, ["planned details", "planned detail", "plan detail", "plan desc"])
    actual_det  = find_col(cols, ["actual details", "actual detail", "actual desc"])

    if planned_det: df = df.rename(columns={planned_det: "Planned Details"})
    if actual_det:  df = df.rename(columns={actual_det: "Actual Details"})

    # Ensure columns exist
    if "Category" not in df.columns: df["Category"] = ""
    if "Subcategory" not in df.columns: df["Subcategory"] = ""
    if "Line Item" not in df.columns: df["Line Item"] = ""
    if "Planned LI" not in df.columns: df["Planned LI"] = 0
    if "Actual LI" not in df.columns: df["Actual LI"] = 0
    if "Planned Efforts (mins)" not in df.columns: df["Planned Efforts (mins)"] = 0.0
    if "Actual Efforts (mins)" not in df.columns: df["Actual Efforts (mins)"] = 0.0
    if "Planned Details" not in df.columns: df["Planned Details"] = ""
    if "Actual Details" not in df.columns: df["Actual Details"] = ""

    # Coerce numeric types
    df["Planned LI"] = pd.to_numeric(df["Planned LI"], errors="coerce").fillna(0).astype(int)
    df["Actual LI"]  = pd.to_numeric(df["Actual LI"], errors="coerce").fillna(0).astype(int)
    df["Planned Efforts (mins)"] = pd.to_numeric(df["Planned Efforts (mins)"], errors="coerce").fillna(0.0)
    df["Actual Efforts (mins)"]  = pd.to_numeric(df["Actual Efforts (mins)"], errors="coerce").fillna(0.0)

    # Ensure strings
    for c in ["Category", "Subcategory", "Line Item", "Planned Details", "Actual Details"]:
        df[c] = df[c].astype(str).fillna("")

    return df

# ---------- robust date parsing ----------
def parse_any_date(val):
    """Robust parse for datetimes, Excel serials, and many string formats."""
    if pd.isna(val):
        return pd.NaT

    # Already timestamp/datetime
    if isinstance(val, (pd.Timestamp, datetime, np.datetime64)):
        return pd.to_datetime(val, errors="coerce").normalize()

    # If numeric-like (Excel serial)
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        try:
            return pd.to_datetime(val, unit="d", origin="1899-12-30", errors="coerce").normalize()
        except Exception:
            pass

    s = str(val).strip()
    if not s:
        return pd.NaT

    # Try pandas parse with month-first, then day-first
    dt = pd.to_datetime(s, errors="coerce", dayfirst=False, infer_datetime_format=True)
    if pd.isna(dt):
        dt = pd.to_datetime(s, errors="coerce", dayfirst=True, infer_datetime_format=True)
    if pd.isna(dt):
        return pd.NaT
    return pd.to_datetime(dt).normalize()

# ---------- io helper: .xlsx only ----------
def load_xlsx(file_storage) -> pd.DataFrame:
    name = (file_storage.filename or "").lower()
    if not name.endswith(".xlsx"):
        raise ValueError("Only .xlsx files accepted.")
    data = io.BytesIO(file_storage.read())
    # Read without forcing types; we'll normalize later
    return pd.read_excel(data, sheet_name=0)

# ---------- week helpers ----------
def get_week_ranges(today: date) -> Tuple[Tuple[date, date], Tuple[date, date]]:
    start_current = today - timedelta(days=today.weekday())  # Monday
    end_current = start_current + timedelta(days=6)
    start_next = end_current + timedelta(days=1)
    end_next = start_next + timedelta(days=6)
    return (start_current, end_current), (start_next, end_next)

# ---------- section builders ----------
def current_week_rows(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["LI Match"] = np.where(out["Planned LI"] == out["Actual LI"], "Match", "Mismatch")
    out["Effort Δ (mins)"] = out["Actual Efforts (mins)"] - out["Planned Efforts (mins)"]
    cols = [
        "Date_only", "Category", "Subcategory", "Line Item",
        "Planned LI", "Actual LI", "LI Match",
        "Planned Efforts (mins)", "Actual Efforts (mins)", "Effort Δ (mins)",
        "Planned Details", "Actual Details"
    ]
    for c in cols:
        if c not in out.columns:
            out[c] = "" if "Details" in c or "Category" in c else 0
    return out[cols]

def next_week_rows(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    cols = ["Date_only", "Category", "Subcategory", "Line Item", "Planned LI", "Planned Efforts (mins)", "Planned Details"]
    for c in cols:
        if c not in out.columns:
            out[c] = "" if "Details" in c or "Category" in c else 0
    return out[cols]

def deviation_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=[
            "Category","Subcategory","Line Item",
            "Planned LI","Actual LI","LI Δ",
            "Planned Efforts (mins)","Actual Efforts (mins)","Effort Δ (mins)","Effort Δ %"
        ])
    g = df.groupby(["Category","Subcategory","Line Item"], dropna=False).agg({
        "Planned LI":"sum","Actual LI":"sum",
        "Planned Efforts (mins)":"sum","Actual Efforts (mins)":"sum"
    }).reset_index()
    g["LI Δ"] = g["Actual LI"] - g["Planned LI"]
    g["Effort Δ (mins)"] = g["Actual Efforts (mins)"] - g["Planned Efforts (mins)"]
    g["Effort Δ %"] = g.apply(
        lambda r: (r["Effort Δ (mins)"] / r["Planned Efforts (mins)"] * 100.0) if r["Planned Efforts (mins)"] else 0.0,
        axis=1
    )
    g = g.round({"Effort Δ (mins)":2,"Effort Δ %":1})
    return g[[
        "Category","Subcategory","Line Item",
        "Planned LI","Actual LI","LI Δ",
        "Planned Efforts (mins)","Actual Efforts (mins)","Effort Δ (mins)","Effort Δ %"
    ]]

# ---------- HTML helpers ----------
def html_table(df: pd.DataFrame, empty_msg="No data"):
    if df.empty:
        return f"<div style='padding:10px;background:#fff8f0;border-radius:6px;color:#92400e'>{empty_msg}</div>"
    thead = "".join(f"<th style='padding:8px;text-align:left;border-bottom:1px solid #e6eefb'>{c}</th>" for c in df.columns)
    rows_html = ""
    for i, (_, r) in enumerate(df.iterrows()):
        bg = "#ffffff" if i%2==0 else "#f8fafc"
        tds = "".join(f"<td style='padding:8px;border-bottom:1px solid #f1f5f9'>{r[c]}</td>" for c in df.columns)
        rows_html += f"<tr style='background:{bg}'>{tds}</tr>"
    return f"<table style='width:100%;border-collapse:collapse;font:13px Arial,Helvetica'>{'<thead><tr>'+thead+'</tr></thead>'}<tbody>{rows_html}</tbody></table>"

def build_email_html(cur_df, nxt_df, dev_df, start_c, end_c, start_n, end_n):
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    return f"""
    <div style="font-family:Arial,Helvetica,sans-serif;color:#0f172a">
      <h2>Weekly Summary Report</h2>
      <div style="color:#6b7280">Generated: {now}</div>

      <h3>1) Current Week Statistics ({start_c} to {end_c})</h3>
      {html_table(cur_df, 'No current-week records')}

      <h3 style="margin-top:18px">2) Next Week Plan ({start_n} to {end_n})</h3>
      {html_table(nxt_df, 'No next-week plan')}

      <h3 style="margin-top:18px">3) Plan vs Actual Deviation / Interference</h3>
      {html_table(dev_df, 'No deviation data')}
    </div>
    """

# ---------- API ----------
@app.route("/upload_timesheet", methods=["POST"])
def upload_timesheet():
    try:
        receiver = request.form.get("receiver_email") or request.form.get("email")
        file = request.files.get("file")
        if not receiver or not file:
            return jsonify({"error":"receiver_email and .xlsx file required"}), 400

        #  Use helper load_xlsx()
        df_raw = load_xlsx(file)

        df = normalize_schema(df_raw)

        # parse dates robustly
        parsed = df["Date"].apply(parse_any_date)

        if parsed.isna().all():
            return jsonify({"error":"No parsable dates found in 'Date' column"}), 400

        df["Date_parsed"] = parsed
        df["Date_only"] = df["Date_parsed"].dt.date
        df = df.dropna(subset=["Date_only"]).copy()

        today = datetime.now().date()
        (start_c, end_c), (start_n, end_n) = get_week_ranges(today)

        cur = df[(df["Date_only"] >= start_c) & (df["Date_only"] <= end_c)].copy()
        nxt = df[(df["Date_only"] >= start_n) & (df["Date_only"] <= end_n)].copy()

        cur_stats = current_week_rows(cur)
        nxt_plan = next_week_rows(nxt)
        dev_sum = deviation_summary(cur)

        html_body = build_email_html(cur_stats, nxt_plan, dev_sum, start_c, end_c, start_n, end_n)

        # send email
        msg = MIMEMultipart("alternative")
        msg["Subject"] = f"Weekly Summary: {start_c} to {end_c}"
        msg["From"] = EMAIL_FROM
        msg["To"] = receiver
        msg.attach(MIMEText(html_body, "html"))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) as server:
            server.ehlo()
            if SMTP_PORT == 587:
                server.starttls()
                server.ehlo()
            server.login(SMTP_USERNAME, SMTP_PASSWORD)
            server.sendmail(EMAIL_FROM, [receiver], msg.as_string())

        return jsonify({
            "message":"Email sent.",
            "current_week_range": f"{start_c} to {end_c}",
            "next_week_range": f"{start_n} to {end_n}",
            "current_rows": int(len(cur_stats)),
            "next_rows": int(len(nxt_plan))
        }), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/")
def home():
    return render_template("index.html") 

if __name__ == "__main__":
    # pip install flask flask-cors pandas numpy openpyxl
    app.run(port=5000, debug=True)
