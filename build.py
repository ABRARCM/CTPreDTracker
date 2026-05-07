#!/usr/bin/env python3
"""Build CT Pre-D Tracker dashboard — PRE (Pre-D Tracker) + POST REVIEW (Pre-Auth Tracker)."""
import pandas as pd
import numpy as np
import json, os
from datetime import datetime, datetime as dt

DIR = os.path.dirname(__file__)
OD = os.path.expanduser("~/Library/CloudStorage/OneDrive-ChildSmilesGroup,LLC(2)/ABRA RCM - CT")
SRC_PRE = os.path.join(OD, "CT PRE-D TRACKER.xlsx")
SRC_POST = os.path.join(OD, "Scrubbing Claims Department", "PRE-AUTH TRACKER.xlsx")
TEMPLATE = os.path.join(DIR, "index.html")
OUT = os.path.join(DIR, "index.html")

def parse_d(s):
    try:
        parts = s.split("/")
        return dt(int(parts[2]), int(parts[0]), int(parts[1]))
    except:
        return dt(1900,1,1)

def safe_int(x):
    if pd.isna(x) or str(x).strip() == "":
        return ""
    try:
        return str(int(float(x)))
    except:
        return str(x)

# ── Categorise PRE-D STATUS ───────────────────────────────────────────────
def categorize(s):
    if pd.isna(s) or str(s).strip() == "":
        return "Unknown"
    s_upper = str(s).upper()
    has_approved = "APPROVED" in s_upper
    has_pending = "PENDING REVIEW" in s_upper or "PENDED" in s_upper or "PENDING" in s_upper
    has_denied = "DENIED" in s_upper or "NOT COVERED" in s_upper
    has_not_found = "NOT FOUND" in s_upper
    has_missing = "MISSING" in s_upper or ("X-RAY" in s_upper and "REQ" in s_upper)
    has_no_info = "NO INFORMATION" in s_upper
    has_no_dos = "DATE OF SERVICE" in s_upper or "NO DATE" in s_upper
    has_not_meet = "NOT MEET" in s_upper or "DO NOT MEET" in s_upper
    has_not_eligible = "NOT ELIGIBLE" in s_upper
    has_resubmitted = "RESUBMITTED" in s_upper

    if has_approved and not has_pending and not has_denied and not has_not_found:
        return "Approved"
    if has_approved and has_not_found:
        return "Partially Approved"
    if has_denied:
        return "Denied"
    if has_not_eligible:
        return "Not Eligible"
    if has_resubmitted:
        return "Resubmitted"
    if has_pending:
        return "Pending Review"
    if has_not_meet:
        return "Does Not Meet Standards"
    if has_missing:
        return "Missing Info / X-Ray Required"
    if has_no_dos:
        return "Missing Date of Service"
    if has_no_info:
        return "No Subscriber Info"
    if has_not_found:
        return "Not Found"
    return "Other"


# ═══════════════════════════════════════════════════════════════════════════
# PRE data (CT PRE-D TRACKER)
# ═══════════════════════════════════════════════════════════════════════════
SHEET_TO_OFFICE = {
    "DERBY": "Derby", "NORWALK GP": "Norwalk GP", "BRIDGEPORT": "Bridgeport",
    "DANBURY": "Danbury", "NORWALK": "Norwalk", "STAMFORD": "Stamford",
}

xls_pre = pd.ExcelFile(SRC_PRE, engine="openpyxl")
frames_pre = []
for sheet in xls_pre.sheet_names:
    if sheet not in SHEET_TO_OFFICE:
        continue
    office = SHEET_TO_OFFICE[sheet]
    tmp = pd.read_excel(xls_pre, sheet_name=sheet)
    tmp["OFFICE"] = office
    frames_pre.append(tmp)

df_pre = pd.concat(frames_pre, ignore_index=True)
df_pre = df_pre.rename(columns={
    "DATE": "date", "PATIENT ID": "patient_id", "PATIENT NAME": "patient_name",
    "SUBSCRIBER ID": "subscriber_id", "PROVIDER": "provider", "PROC": "proc_code",
    "TOOTH #": "tooth", "DOS": "dos", "PRE-D STATUS": "pred_status",
    "ACTION": "action", "OFFICE": "office",
})
keep_pre = ["date","patient_id","subscriber_id","provider","proc_code","tooth","dos","pred_status","action","office"]
df_pre = df_pre[[c for c in keep_pre if c in df_pre.columns]]

for col in ["date", "dos"]:
    df_pre[col] = pd.to_datetime(df_pre[col], errors="coerce").dt.strftime("%m/%d/%Y").replace("NaT", "")

df_pre["patient_id"] = df_pre["patient_id"].apply(safe_int)
df_pre["subscriber_id"] = df_pre["subscriber_id"].apply(safe_int)
df_pre["tooth"] = df_pre["tooth"].apply(lambda x: safe_int(x) if pd.notna(x) else "")
df_pre["category"] = df_pre["pred_status"].apply(categorize)
df_pre = df_pre.fillna("")

pre_dates = [d for d in df_pre["date"].unique().tolist() if d]
pre_dates.sort(key=parse_d, reverse=True)
pre_offices = sorted(df_pre["office"].unique().tolist())

print(f"PRE: {len(df_pre)} rows, {len(pre_offices)} offices, most recent: {pre_dates[0] if pre_dates else 'N/A'}")
print(f"PRE categories: {df_pre['category'].value_counts().to_dict()}")


# ═══════════════════════════════════════════════════════════════════════════
# POST REVIEW data (PRE-AUTH TRACKER)
# ═══════════════════════════════════════════════════════════════════════════
xls_post = pd.ExcelFile(SRC_POST, engine="openpyxl")
frames_post = []
for sheet in ["ORTHO", "PEDO&GP"]:
    tmp = pd.read_excel(xls_post, sheet_name=sheet)
    tmp["_sheet"] = sheet
    frames_post.append(tmp)

df_post = pd.concat(frames_post, ignore_index=True)
df_post = df_post.rename(columns={
    "Clinic": "office",
    "PatNum": "patient_id",
    "ProcCode": "proc_code",
    "ProcDate": "dos",
    "SubscriberID": "subscriber_id",
    "Processed Status": "pred_status",
    "Processed On": "date",
    "Success": "action",
    "Error Description": "error_desc",
    "ProcFee": "proc_fee",
    "InsPayEst": "ins_pay_est",
})

df_post["dept"] = df_post["_sheet"].map({"ORTHO": "Ortho", "PEDO&GP": "Pedo/GP"}).fillna("")
keep_post = ["office","patient_id","subscriber_id","proc_code","dos","pred_status","date","action","proc_fee","ins_pay_est","dept"]
df_post = df_post[[c for c in keep_post if c in df_post.columns]]

# Clean office names
office_map = {
    "Bridgeport": "Bridgeport", "Stamford": "Stamford", "Norwalk": "Norwalk",
    "Danbury": "Danbury", "Derby": "Derby", "Norwalk GP": "Norwalk GP",
}
df_post["office"] = df_post["office"].apply(lambda x: office_map.get(str(x).strip(), str(x).strip()) if pd.notna(x) else "")
# Filter out rows with invalid office
valid_offices = set(office_map.values())
df_post = df_post[df_post["office"].isin(valid_offices)].copy()

for col in ["date", "dos"]:
    df_post[col] = pd.to_datetime(df_post[col], errors="coerce").dt.strftime("%m/%d/%Y").replace("NaT", "")

df_post["patient_id"] = df_post["patient_id"].apply(safe_int)
df_post["subscriber_id"] = df_post["subscriber_id"].apply(safe_int)
df_post["proc_fee"] = df_post["proc_fee"].apply(lambda x: f"${x:,.2f}" if pd.notna(x) and x != "" else "")
df_post["ins_pay_est"] = df_post["ins_pay_est"].apply(lambda x: f"${x:,.2f}" if pd.notna(x) and x != "" else "")

# Provider column doesn't exist in post — leave blank
df_post["provider"] = ""
df_post["tooth"] = ""

df_post["category"] = df_post["pred_status"].apply(categorize)
# Also use the Success/action column for categorisation where pred_status is vague
def refine_category(row):
    act = str(row.get("action", "")).upper().strip()
    if act in ("APPROVED", "YES"):
        return "Approved"
    if act == "DENIED":
        return "Denied"
    if act == "PENDING":
        return "Pending Review"
    if act == "MISSING":
        return "Missing Info / X-Ray Required"
    if act == "RESUBMITTED":
        return "Resubmitted"
    if act == "HOLD":
        return "On Hold"
    if act == "PPO":
        return "PPO"
    if act in ("ORTHO", "ORTHO"):
        return "Ortho Referral"
    return row["category"]

df_post["category"] = df_post.apply(refine_category, axis=1)
df_post = df_post.fillna("")

post_dates = [d for d in df_post["date"].unique().tolist() if d]
post_dates.sort(key=parse_d, reverse=True)
post_offices = sorted(df_post["office"].unique().tolist())

print(f"POST: {len(df_post)} rows, {len(post_offices)} offices, most recent: {post_dates[0] if post_dates else 'N/A'}")
print(f"POST categories: {df_post['category'].value_counts().to_dict()}")


# ═══════════════════════════════════════════════════════════════════════════
# Inject into template
# ═══════════════════════════════════════════════════════════════════════════
timestamp = datetime.now().strftime("%B %d, %Y %I:%M %p")

# Combine all offices
all_offices = sorted(set(pre_offices + post_offices))

with open(TEMPLATE, "r") as f:
    html = f.read()

html = html.replace("__PRE_DATA_JSON__", json.dumps(df_pre.to_dict(orient="records"), default=str))
html = html.replace("__POST_DATA_JSON__", json.dumps(df_post.to_dict(orient="records"), default=str))
html = html.replace("__PRE_DATES_JSON__", json.dumps(pre_dates))
html = html.replace("__POST_DATES_JSON__", json.dumps(post_dates))
html = html.replace("__OFFICES_JSON__", json.dumps(all_offices))
html = html.replace("__TIMESTAMP__", timestamp)

with open(OUT, "w") as f:
    f.write(html)

print(f"\nDashboard written to {OUT}")
print(f"  PRE: {len(df_pre)} rows | POST: {len(df_post)} rows")
