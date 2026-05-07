#!/usr/bin/env python3
"""Push CT Pre-D Tracker data from Excel to Firebase Realtime Database.

Usage:
    python3 auto_update.py              # Update both POST REVIEW + PRE DETERMINATION
    python3 auto_update.py --pre-det    # Update PRE DETERMINATION only (Thursday)

Schedule:
    Monday 8:00 AM    — full update (both)
    Thursday 8:00 AM  — --pre-det (Pre Determination only)
"""
import pandas as pd
import numpy as np
import json, os, sys, requests
from datetime import datetime, datetime as dt

DIR = os.path.dirname(os.path.abspath(__file__))
LOG = os.path.join(DIR, "auto_update.log")
OD = os.path.expanduser("~/Library/CloudStorage/OneDrive-ChildSmilesGroup,LLC(2)/ABRA RCM - CT")
SRC_PRE = os.path.join(OD, "CT PRE-D TRACKER.xlsx")
SRC_POST = os.path.join(OD, "Scrubbing Claims Department", "PRE-AUTH TRACKER.xlsx")
FB_URL = "https://ct-pred-tracker-default-rtdb.firebaseio.com"

def log(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG, "a") as f:
        f.write(line + "\n")

def safe_int(x):
    if pd.isna(x) or str(x).strip() == "":
        return ""
    try:
        return str(int(float(x)))
    except:
        return str(x)

def parse_d(s):
    try:
        parts = s.split("/")
        return dt(int(parts[2]), int(parts[0]), int(parts[1]))
    except:
        return dt(1900, 1, 1)

def categorize(s):
    if pd.isna(s) or str(s).strip() == "":
        return "Unknown"
    su = str(s).upper()
    has_approved = "APPROVED" in su
    has_pending = "PENDING REVIEW" in su or "PENDED" in su or "PENDING" in su
    has_denied = "DENIED" in su or "NOT COVERED" in su
    has_not_found = "NOT FOUND" in su
    has_missing = "MISSING" in su or ("X-RAY" in su and "REQ" in su)
    has_no_info = "NO INFORMATION" in su
    has_no_dos = "DATE OF SERVICE" in su or "NO DATE" in su
    has_not_meet = "NOT MEET" in su or "DO NOT MEET" in su
    has_not_eligible = "NOT ELIGIBLE" in su
    has_resubmitted = "RESUBMITTED" in su

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

def fb_put(path, data):
    """PUT data to Firebase REST API."""
    url = f"{FB_URL}/{path}.json"
    r = requests.put(url, json=data, timeout=120)
    if r.status_code != 200:
        log(f"Firebase PUT failed ({r.status_code}): {path} — {r.text[:200]}")
        return False
    return True


# ═══════════════════════════════════════════════════════════════════════════
def build_post_review():
    """POST REVIEW = CT PRE-D TRACKER.xlsx (the manual tracker per office)."""
    log("Reading POST REVIEW data (CT PRE-D TRACKER.xlsx)...")
    SHEET_MAP = {
        "DERBY": "Derby", "NORWALK GP": "Norwalk GP", "BRIDGEPORT": "Bridgeport",
        "DANBURY": "Danbury", "NORWALK": "Norwalk", "STAMFORD": "Stamford",
    }
    xls = pd.ExcelFile(SRC_PRE, engine="openpyxl")
    frames = []
    for sheet in xls.sheet_names:
        if sheet not in SHEET_MAP:
            continue
        tmp = pd.read_excel(xls, sheet_name=sheet)
        tmp["OFFICE"] = SHEET_MAP[sheet]
        frames.append(tmp)

    df = pd.concat(frames, ignore_index=True)
    df = df.rename(columns={
        "DATE": "date", "PATIENT ID": "patient_id", "SUBSCRIBER ID": "subscriber_id",
        "PROVIDER": "provider", "PROC": "proc_code", "TOOTH #": "tooth",
        "DOS": "dos", "PRE-D STATUS": "pred_status", "ACTION": "action", "OFFICE": "office",
    })
    keep = ["date", "patient_id", "subscriber_id", "provider", "proc_code", "tooth", "dos", "pred_status", "action", "office"]
    df = df[[c for c in keep if c in df.columns]]

    for col in ["date", "dos"]:
        df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%m/%d/%Y").replace("NaT", "")
    df["patient_id"] = df["patient_id"].apply(safe_int)
    df["subscriber_id"] = df["subscriber_id"].apply(safe_int)
    df["tooth"] = df["tooth"].apply(lambda x: safe_int(x) if pd.notna(x) else "")
    df["category"] = df["pred_status"].apply(categorize)
    df = df.fillna("")

    dates = sorted([d for d in df["date"].unique().tolist() if d], key=parse_d, reverse=True)
    records = df.to_dict(orient="records")
    log(f"  POST REVIEW: {len(records)} rows, most recent: {dates[0] if dates else 'N/A'}")
    return records, dates


def build_pre_determination():
    """PRE DETERMINATION = PRE-AUTH TRACKER.xlsx (Ortho + Pedo&GP)."""
    log("Reading PRE DETERMINATION data (PRE-AUTH TRACKER.xlsx)...")
    xls = pd.ExcelFile(SRC_POST, engine="openpyxl")
    frames = []
    for sheet in ["ORTHO", "PEDO&GP"]:
        tmp = pd.read_excel(xls, sheet_name=sheet)
        tmp["_sheet"] = sheet
        frames.append(tmp)

    df = pd.concat(frames, ignore_index=True)
    df = df.rename(columns={
        "Clinic": "office", "PatNum": "patient_id", "ProcCode": "proc_code",
        "ProcDate": "dos", "SubscriberID": "subscriber_id",
        "Processed Status": "pred_status", "Processed On": "date",
        "Success": "action", "ProcFee": "proc_fee", "InsPayEst": "ins_pay_est",
    })
    df["dept"] = df["_sheet"].map({"ORTHO": "Ortho", "PEDO&GP": "Pedo/GP"}).fillna("")
    keep = ["office", "patient_id", "subscriber_id", "proc_code", "dos", "pred_status", "date", "action", "proc_fee", "ins_pay_est", "dept"]
    df = df[[c for c in keep if c in df.columns]]

    office_map = {"Bridgeport": "Bridgeport", "Stamford": "Stamford", "Norwalk": "Norwalk",
                  "Danbury": "Danbury", "Derby": "Derby", "Norwalk GP": "Norwalk GP"}
    df["office"] = df["office"].apply(lambda x: office_map.get(str(x).strip(), str(x).strip()) if pd.notna(x) else "")
    df = df[df["office"].isin(set(office_map.values()))].copy()

    for col in ["date", "dos"]:
        df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%m/%d/%Y").replace("NaT", "")
    df["patient_id"] = df["patient_id"].apply(safe_int)
    df["subscriber_id"] = df["subscriber_id"].apply(safe_int)
    df["proc_fee"] = df["proc_fee"].apply(lambda x: f"${x:,.2f}" if pd.notna(x) and x != "" else "")
    df["ins_pay_est"] = df["ins_pay_est"].apply(lambda x: f"${x:,.2f}" if pd.notna(x) and x != "" else "")
    df["provider"] = ""
    df["tooth"] = ""
    df["category"] = df["pred_status"].apply(categorize)

    def refine(row):
        act = str(row.get("action", "")).upper().strip()
        if act in ("APPROVED", "YES"): return "Approved"
        if act == "DENIED": return "Denied"
        if act == "PENDING": return "Pending Review"
        if act == "MISSING": return "Missing Info / X-Ray Required"
        if act == "RESUBMITTED": return "Resubmitted"
        if act == "HOLD": return "On Hold"
        if act == "PPO": return "PPO"
        if act in ("ORTHO", "ORTHO"): return "Ortho Referral"
        return row["category"]
    df["category"] = df.apply(refine, axis=1)
    df = df.fillna("")

    dates = sorted([d for d in df["date"].unique().tolist() if d], key=parse_d, reverse=True)
    records = df.to_dict(orient="records")
    log(f"  PRE DETERMINATION: {len(records)} rows, most recent: {dates[0] if dates else 'N/A'}")
    return records, dates


# ═══════════════════════════════════════════════════════════════════════════
def fb_get(path):
    """GET data from Firebase REST API."""
    url = f"{FB_URL}/{path}.json"
    r = requests.get(url, timeout=120)
    if r.status_code == 200:
        return r.json()
    log(f"Firebase GET failed ({r.status_code}): {path}")
    return None


def push_to_firebase(section, new_records, new_dates, full_refresh=False):
    """Push only the latest date's data to Firebase (incremental), or full refresh."""
    if full_refresh:
        log(f"Full refresh: pushing {len(new_records)} rows to Firebase → pred_tracker/{section}/data ...")
        ok = fb_put(f"pred_tracker/{section}/data", new_records)
        ok2 = fb_put(f"pred_tracker/{section}/dates", new_dates)
    else:
        # Get existing data from Firebase
        existing = fb_get(f"pred_tracker/{section}/data") or []
        existing_dates = fb_get(f"pred_tracker/{section}/dates") or []

        if not existing:
            log(f"No existing data — doing full push of {len(new_records)} rows.")
            ok = fb_put(f"pred_tracker/{section}/data", new_records)
            ok2 = fb_put(f"pred_tracker/{section}/dates", new_dates)
        else:
            # Find the latest date in new data
            latest_date = new_dates[0] if new_dates else None
            if not latest_date:
                log("No dates found in new data, skipping.")
                return False

            # Get only rows with the latest date from new data
            latest_rows = [r for r in new_records if r.get("date") == latest_date]
            log(f"Latest date: {latest_date} — {len(latest_rows)} new rows")

            # Remove any existing rows with this date (in case of re-run)
            existing = [r for r in existing if r and r.get("date") != latest_date]
            log(f"Existing rows (after removing {latest_date}): {len(existing)}")

            # Append new rows
            merged = existing + latest_rows
            log(f"Merged total: {len(merged)} rows — pushing to Firebase...")
            ok = fb_put(f"pred_tracker/{section}/data", merged)

            # Update dates list
            if latest_date not in existing_dates:
                merged_dates = [latest_date] + existing_dates
            else:
                merged_dates = new_dates  # use fresh sorted dates
            ok2 = fb_put(f"pred_tracker/{section}/dates", merged_dates)

    if ok:
        log(f"  Data pushed successfully.")
    if ok2:
        log(f"  Dates pushed.")
    ts = datetime.now().strftime("%m/%d/%Y %I:%M %p")
    total = len(new_records) if full_refresh else len(fb_get(f"pred_tracker/{section}/data") or [])
    fb_put(f"pred_tracker/{section}/meta", {"lastUpdated": ts, "rowCount": total})
    return ok and ok2


def main():
    pre_det_only = "--pre-det" in sys.argv
    full_refresh = "--full" in sys.argv
    mode = "Pre Determination only" if pre_det_only else "Full update (both)"
    if full_refresh:
        mode += " [FULL REFRESH]"
    log(f"\n{'='*60}")
    log(f"Auto-update started — {mode}")

    all_offices = set()

    if not pre_det_only:
        records, dates = build_post_review()
        for r in records:
            all_offices.add(r.get("office", ""))
        push_to_firebase("pre", records, dates, full_refresh=full_refresh)

    records, dates = build_pre_determination()
    for r in records:
        all_offices.add(r.get("office", ""))
    push_to_firebase("post", records, dates, full_refresh=full_refresh)

    # Update shared offices list
    offices = sorted([o for o in all_offices if o])
    fb_put("pred_tracker/offices", offices)
    log(f"Offices updated: {offices}")

    log(f"Auto-update complete.")
    log(f"{'='*60}\n")


if __name__ == "__main__":
    main()
