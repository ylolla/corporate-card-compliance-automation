"""
iExpense Compliance Automation Suite — Flask App
==========================================
Single-file app. Ports the proven notebook logic into a web tool run in the
browser: upload two files -> review -> preview emails -> send.

Works on BOTH Mac and Windows. See SETUP GUIDE (separate document) for install.

HOW TO RUN:
    pip install flask pandas openpyxl
    (Windows, to send real email:  pip install pywin32)
    python app.py
    then open http://127.0.0.1:5000 in the browser.

SENDING SAFETY:
    The app never sends silently. When you click Send, it asks you to choose
    TEST (simulate only, writes a log) or SEND FOR REAL. Real sending uses the
    installed Outlook and only works on Windows.

The fiscal-officer reference file (All Orgs Detail.xlsx) is NOT uploaded each
run — it lives in the shared folder and only changes when divisions change.
"""

import os
import sys
import platform
import traceback
from datetime import datetime

import numpy as np
import pandas as pd
from flask import Flask, request, jsonify, render_template_string

# =============================================================
# CONFIG — EDIT THIS ONE BLOCK
# =============================================================
# The shared folder both Mac and Windows reach over the Office VPN.
# Point this at the SAME shared location on each machine. Everything else
# (uploads, logs, the All Orgs reference file) is found relative to it.
#
# Examples:
#   Mac:      "/Users/demo0000/Desktop/iExpense Late Payment Submissions"
#   Windows:  r"Z:\Finance\iExpense Late Payment Submissions"
#             (use the r"..." form on Windows so backslashes work)
#
# If you leave it as "" the app uses the folder app.py sits in.
SHARED_FOLDER = ""

# ---- Everything below is derived automatically; no need to edit ----
APP_DIR = SHARED_FOLDER.strip() or os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(APP_DIR, "_uploads")
LOG_PATH = os.path.join(APP_DIR, "send_log.csv")

# Fixed reference file (fiscal officers). Must sit in APP_DIR.
ORGS_PATH = os.path.join(APP_DIR, "All Orgs Detail.xlsx")
ORGS_PATH_FALLBACK = os.path.join(os.path.dirname(os.path.abspath(__file__)), "All Orgs Detail.xlsx")

# Monthly name -> NetID directory (last-ditch recovery). Optional; must sit in APP_DIR.
EMPDASH_PATH = os.path.join(APP_DIR, "Employee Data Dashboard - Updated.csv")

SOURCE_SHEET = "Sheet1"
IS_WINDOWS = platform.system() == "Windows"

os.makedirs(UPLOAD_DIR, exist_ok=True)

# ---- Status rules (from the notebook) ----
EMAILABLE = {"Pending Manager Approval", "Pending Your Resolution", "Rejected", "Withdrawn"}
EXCLUDABLE = {"Paid", "Pending Payables Approval", "Ready for Payment", "Deleted"}
REVIEW_STATUS = "Pending System Administrator Action"
MANAGER_FAULT = {"Pending Manager Approval"}

# All six aging buckets, ordered oldest-worst. The UI lets the process owner tick which ones
# to email; DEFAULT_BUCKETS is the pre-checked set when the page loads.
BUCKET_ORDER = {
    "< 30 Days": 1, "30 - 60 Days": 2, "60 - 90 Days": 3,
    "90 - 120 Days": 4, "120 - 150 Days": 5, "> 150 Days": 6,
}
ALL_BUCKETS = ["< 30 Days", "30 - 60 Days", "60 - 90 Days",
               "90 - 120 Days", "120 - 150 Days", "> 150 Days"]
DEFAULT_BUCKETS = ["120 - 150 Days", "> 150 Days"]

# =============================================================
# EDITABLE WORDING — the owner's finalized, approved templates (per bucket)
# =============================================================
# Subject line per aging bucket (the process owner-specified). Falls back to the <30 wording.
SUBJECT_BY_BUCKET = {
    "< 30 Days": "Outstanding Corporate Card Transactions",
    "30 - 60 Days": "Outstanding Corporate Card Transactions - Action Required",
    "60 - 90 Days": "Outstanding Corporate Card Transactions over 60 days - Action Required",
    "90 - 120 Days": "Outstanding Corporate Card Transactions over 90 days - Action Required",
    "120 - 150 Days": "Outstanding Corporate Card Transactions over 120 days - Action Required",
    "> 150 Days": "Outstanding Corporate Card Transactions over 150 days - Action Required",
}
SUBJECT_DEFAULT = "Outstanding Corporate Card Transactions"
GREETING = "Hello,"

# --- Opening paragraph, per aging bucket ---
SUSPENSION_INTRO_BY_BUCKET = {
    "< 30 Days": (
        "The below corporate card transactions are outstanding and need to be settled "
        "in iExpense at your earliest convenience."
    ),
    "30 - 60 Days": (
        "The corporate card transactions listed below remain outstanding, please submit "
        "and settle these transactions in iExpense at your earliest convenience."
    ),
    "60 - 90 Days": (
        "Some or all of the corporate card transactions listed below have been outstanding "
        "for more than 60 days. Please submit and reconcile these expenses in iExpense as "
        "soon as possible to avoid potential suspension of your corporate card."
    ),
    "90 - 120 Days": (
        "Our records indicate that the corporate card transaction(s) listed below remain "
        "outstanding. Because some or all of these transactions are more than 90 days old, "
        "your card has been suspended effective immediately and will remain suspended until "
        "the transactions are fully settled in the iExpense system."
    ),
    "120 - 150 Days": (
        "Our records indicate that the corporate card transaction(s) listed below remain "
        "outstanding. Because some or all of these transactions are more than 120 days old, "
        "your card has been suspended effective immediately and will remain suspended until "
        "the transactions are fully settled in the iExpense system."
    ),
    "> 150 Days": (
        "Our records indicate that the corporate card transaction(s) listed below remain "
        "outstanding. Because some or all of these transactions are more than 150 days old, "
        "your card has been suspended effective immediately and will remain suspended until "
        "the transactions are fully settled in the iExpense system."
    ),
}
SUSPENSION_INTRO_DEFAULT = (
    "The below corporate card transactions are outstanding and need to be settled "
    "in iExpense at your earliest convenience."
)

# Added when the person has Pending Manager Approval items
PMA_INTRO = (
    "One or more of the below expense report(s) are pending manager approval. "
    "Please review and approve at your earliest convenience."
)

# Policy paragraph, per bucket. Under 60 days: standard. 60+ days: escalated.
_POLICY_STANDARD = (
    "In accordance with the Corporate Card Cardholder Agreement and the Business Expense "
    "Policy, all expenses, including corporate card charges, must be reported and settled "
    "in the iExpense system within 60 days of the expense date. Cards with transactions "
    "that remain unsettled beyond this timeframe are subject to suspension until the "
    "outstanding transactions are resolved. As a reminder, corporate card expenses should "
    "be settled within 15 business days of the expense date. The 60-day period is the "
    "maximum allowed, and settling promptly helps keep your account in good standing."
)
_POLICY_ESCALATED = (
    "In accordance with the Corporate Card Cardholder Agreement and the Business Expense "
    "Policy, all expenses, including corporate card charges, must be reported and settled "
    "in the iExpense system within 60 days of the expense date. Cards with transactions "
    "that remain unsettled beyond this timeframe are subject to suspension until the "
    "outstanding transactions are resolved. In addition to card suspension, expenses that "
    "are not settled within the required timeframe may be denied reimbursement, and any "
    "reimbursement made after the deadline may be treated as taxable income to you. If an "
    "advance was provided and remains unsettled, it may be reported to Payroll as taxable "
    "income, and future advances may be restricted."
)
POLICY_BODY_BY_BUCKET = {
    "< 30 Days": _POLICY_STANDARD,
    "30 - 60 Days": _POLICY_STANDARD,
    "60 - 90 Days": _POLICY_ESCALATED,
    "90 - 120 Days": _POLICY_ESCALATED,
    "120 - 150 Days": _POLICY_ESCALATED,
    "> 150 Days": _POLICY_ESCALATED,
}
POLICY_BODY_DEFAULT = _POLICY_STANDARD

# Closing line + signature (same on every email)
CLOSING_LINE = (
    "Please contact me directly if you have any questions or need assistance resolving "
    "the outstanding transactions."
)
# 60+ day buckets: from the process owner, full signature.
SIGNATURE = (
    "Thank you,\nthe process owner\n\n"
    "Finance Operations\n"
    "Director, Finance Operations\n"
    "000-000-0000 | 000 Example St | City, ST 00000"
)

# Under-60 day buckets (< 30, 30-60): from the Finance Center, simple signature.
CLOSING_LINE_FC = (
    "Please reach out if you have any questions or need assistance resolving "
    "the outstanding transactions."
)
SIGNATURE_FC = "Thank you,\nFinance Center"

# Sender address for the under-60 (Finance Center) emails. Requires the Outlook
# account running the tool to have Send-As permission on this mailbox; otherwise
# the send is failed with a clear message rather than sent from the wrong person.
FINANCE_CENTER_FROM = "finance.center@example.edu"

# Buckets considered "under 60 days" -> Finance Center sender + simple signature.
UNDER_60_BUCKETS = {"< 30 Days", "30 - 60 Days"}


def is_finance_center_bucket(bucket):
    """True if the person's worst bucket is under 60 days (Finance Center sends it)."""
    return str(bucket).strip() in UNDER_60_BUCKETS

# CC'd on EVERY email (the process owner keeps a copy since sent items sometimes don't land
# in her Outlook Sent folder).
CC_ALWAYS = ["process.owner@example.edu"]

# Policy document links — hyperlinked in the HTML email on these exact phrases.
CARDHOLDER_AGREEMENT_URL = "https://example.edu/policy/cardholder-agreement.pdf"
BUSINESS_EXPENSE_POLICY_URL = "https://example.edu/policy/business-expense-policy"
LINK_PHRASES = [
    ("Corporate Card Cardholder Agreement", CARDHOLDER_AGREEMENT_URL),
    ("Business Expense Policy", BUSINESS_EXPENSE_POLICY_URL),
]

# Dummy Test: real emails send, but every To/CC is forced to a safe address so no
# real cardholder is contacted. Mac sends through Apple Mail (renders HTML), to Yesh;
# Windows sends through Outlook, to the process owner (she runs it there).
DUMMY_RECIPIENT_MAC = "demo.user@example.edu"
DUMMY_RECIPIENT_WINDOWS = "process.owner@example.edu"

TABLE_COLUMNS = [
    "Report Header Id", "Card Holder Name", "Transaction Date",
    "Billed Amount", "Billed Date", "Expense Status",
    "Merchant Name1", "Merchant City", "Merchant Province State",
]

# =============================================================
# LOGIC — ported directly from the notebook, as functions
# =============================================================

def clean_id(series):
    s = series.astype(str).str.strip()
    s = s.str.replace(r"\.0$", "", regex=True)
    s = s.replace({"nan": np.nan, "None": np.nan, "": np.nan})
    return s


def norm_code(val):
    if pd.isna(val):
        return np.nan
    s = str(val).strip()
    if s == "" or s.lower() == "nan":
        return np.nan
    s = s.split("-")[0].strip()
    try:
        return str(int(s))
    except ValueError:
        return s.lstrip("0") or "0"


def build_fo_email(nid):
    """Accepts a single NetID or a '; '-joined list; returns '; '-joined
    netid@example.edu addresses. Blank/invalid NetIDs are skipped."""
    if pd.isna(nid):
        return np.nan
    parts = [p.strip() for p in str(nid).split(";")]
    emails = []
    for p in parts:
        if p and " " not in p and "@" not in p and p.lower() not in ("nan", "none") \
                and "FAKE" not in p.upper():
            emails.append(f"{p}@example.edu")
    return "; ".join(emails) if emails else np.nan


def _norm_status(s):
    if pd.isna(s):
        return None
    s = str(s).strip()
    return None if s in ("", "0", "nan", "None") else s


def _netid_value_ok(nid):
    return (pd.notna(nid) and str(nid).strip() not in ("", "nan", "None")
            and "FAKE" not in str(nid).upper())


def _has_netid(row):
    return _netid_value_ok(row.get("Employee NetID"))


def classify(row):
    """Return (final_status, disposition).

    iExpense (newer) status wins; the owner's is a fallback. The effective status
    decides EXCLUDE / REVIEW / SEND. Only 'Pending System Administrator Action'
    and 'no NetID' go to review.
    """
    t = _norm_status(row["Expense Status"])            # the process owner (old)
    x = _norm_status(row.get("iexp_Expense Status"))   # iExpense (new) -> wins
    effective = x if x is not None else t

    if effective == REVIEW_STATUS:
        return pd.Series([REVIEW_STATUS, "REVIEW"])
    if effective in EXCLUDABLE:
        return pd.Series([effective, "EXCLUDE"])

    fs = effective if effective is not None else ""
    if not _has_netid(row):
        return pd.Series(["No email found (no NetID)", "REVIEW"])
    return pd.Series([fs, "SEND"])


def load_orgs():
    """Load the fixed fiscal-officer reference file."""
    path = ORGS_PATH if os.path.exists(ORGS_PATH) else ORGS_PATH_FALLBACK
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Could not find 'All Orgs Detail.xlsx'. Place it next to app.py "
            f"(looked at: {ORGS_PATH})"
        )
    ira = pd.read_excel(path, sheet_name="IRA Orgs and Group", dtype=str)
    poc = pd.read_excel(path, sheet_name="Compliance POC", dtype=str)

    ira["Group"] = ira["Group"].astype(str).str.strip()
    ira["code_norm"] = ira["GL Org"].apply(norm_code)
    ira_lookup = (ira[["code_norm", "Group"]]
                  .dropna(subset=["code_norm"])
                  .drop_duplicates(subset="code_norm", keep="first"))

    poc["Division"] = poc["Division"].astype(str).str.strip()
    poc["Net ID"] = poc["Net ID"].astype(str).str.strip().replace(
        {"nan": np.nan, "None": np.nan, "": np.nan})
    poc["Fiscal Officer/Compliance POC"] = poc["Fiscal Officer/Compliance POC"].astype(str).str.strip()

    # A division can have MORE THAN ONE fiscal officer — collect ALL of their
    # NetIDs (deduped, order preserved), semicolon-joined, so every FO is CC'd.
    def _join_unique(series):
        seen, out = set(), []
        for v in series.dropna():
            v = str(v).strip()
            if v and v.lower() not in ("nan", "none") and v not in seen:
                seen.add(v)
                out.append(v)
        return "; ".join(out)

    poc_grp = poc.groupby("Division", as_index=False).agg(
        **{"Fiscal Officer/Compliance POC": ("Fiscal Officer/Compliance POC", _join_unique),
           "Net ID": ("Net ID", _join_unique)}
    )
    poc_lookup = poc_grp[["Division", "Fiscal Officer/Compliance POC", "Net ID"]]
    return ira_lookup, poc_lookup


def load_emp_dashboard():
    """Load the monthly Employee Data Dashboard (name -> NetID). Optional; used as
    the last-ditch NetID recovery. Returns (unique_map, ambiguous_set)."""
    path = EMPDASH_PATH if os.path.exists(EMPDASH_PATH) else None
    if path is None:
        return {}, set()
    dash = pd.read_csv(path, dtype=str, encoding="utf-8-sig")
    dash.columns = [c.strip() for c in dash.columns]
    if "Employee" not in dash.columns or "NetID" not in dash.columns:
        return {}, set()
    dash["Employee"] = dash["Employee"].astype(str).str.strip()
    nid = dash["NetID"].astype(str).str.strip()
    nid = nid.mask(nid.str.contains("FAKE", case=False, na=False))
    dash["NetID"] = nid.replace({"": np.nan, "nan": np.nan, "None": np.nan})
    m = dash.dropna(subset=["NetID"]).groupby("Employee")["NetID"].agg(lambda s: set(s.dropna()))
    unique_map = {nm: next(iter(ids)) for nm, ids in m.items() if len(ids) == 1}
    ambiguous = {nm for nm, ids in m.items() if len(ids) > 1}
    return unique_map, ambiguous


def build_cardholder_email(netid):
    """Always netid@example.edu; never the raw Employee Email. FAKE/blank -> none."""
    if pd.isna(netid):
        return np.nan
    nid = str(netid).strip()
    if nid in ("", "nan", "None") or " " in nid or "@" in nid or "FAKE" in nid.upper():
        return np.nan
    return f"{nid}@example.edu"


def summarize_person(g):
    statuses = sorted(g["final_status"].dropna().unique().tolist())
    # Approver CC on ANY emailed row that has a real approver (not just PMA)
    approver_ccs = sorted(set(g["Approver Email"].dropna()) - {""})
    # Fiscal Officer Email may be a '; '-joined list (a division can have 2+ FOs) —
    # split into individual addresses so each is CC'd.
    fo_set = set()
    for v in g["Fiscal Officer Email"].dropna():
        for addr in str(v).split(";"):
            addr = addr.strip()
            if addr:
                fo_set.add(addr)
    fo_ccs = sorted(fo_set)
    worst_idx = g["bucket_rank"].idxmax() if g["bucket_rank"].notna().any() else g.index[0]
    netid = g["Employee NetID"].iloc[0] if "Employee NetID" in g.columns else np.nan
    return pd.Series({
        "Employee NetID": netid,
        "Cardholder Email": build_cardholder_email(netid),
        "num_transactions": len(g),
        "total_billed": round(g["Billed Amount Num"].sum(), 2),
        "statuses": ", ".join(statuses),
        "worst_bucket": g.loc[worst_idx, "Bucket_clean"],
        "approver_ccs": "; ".join(approver_ccs),
        "fiscal_officer_ccs": "; ".join(fo_ccs),
        "department": g["Department Name"].iloc[0] if "Department Name" in g.columns else "",
    })


def process_files(tawnia_path, iexpense_path):
    """THE BRAIN. Trx Id pipeline the notebook proved. Returns everything the UI needs."""
    # ---- Load the owner's file ----
    tawnia = pd.read_excel(tawnia_path, sheet_name=SOURCE_SHEET, dtype=str)

    # ---- Load iExpense ----
    try:
        iexp = pd.read_csv(iexpense_path, dtype=str, encoding="utf-8-sig")
    except UnicodeDecodeError:
        iexp = pd.read_csv(iexpense_path, dtype=str, encoding="latin-1")

    if "Trx Id" not in tawnia.columns:
        raise ValueError("the owner's file has no 'Trx Id' column. Check the sheet/columns.")
    if "Trx Id" not in iexp.columns:
        raise ValueError("The iExpense file has no 'Trx Id' column. Check the export.")

    # ---- Drop blank cardholder names (hard rule) ----
    blank_ch = tawnia["Card Holder Name"].isna() | (tawnia["Card Holder Name"].astype(str).str.strip() == "")
    tawnia = tawnia[~blank_ch].copy()

    # ---- Rename datetime "days outstanding" col IF present ----
    date_col = [c for c in tawnia.columns if not isinstance(c, str)]
    if date_col:
        tawnia = tawnia.rename(columns={date_col[0]: "Days Outstanding"})

    # ---- Clean Trx Id on both sides ----
    tawnia["Trx Id"] = clean_id(tawnia["Trx Id"])
    iexp["Trx Id"] = clean_id(iexp["Trx Id"])
    iexp = iexp[iexp["Trx Id"].notna()].copy()   # drop iExpense rows with blank Trx Id

    # ---- STEP 5b: scrub FAKE netids, fill blanks within iExpense by exact name ----
    _nid = iexp["Employee NetID"].astype(str).str.strip()
    _nid = _nid.mask(_nid.str.contains("FAKE", case=False, na=False))
    iexp["Employee NetID"] = _nid.replace({"": np.nan, "nan": np.nan, "None": np.nan})
    iexp["Employee Name"] = iexp["Employee Name"].astype(str).str.strip()
    name_netids = (iexp.dropna(subset=["Employee NetID"])
                   .groupby("Employee Name")["Employee NetID"].agg(lambda s: set(s.dropna())))
    uniq_name_netid = {nm: next(iter(ids)) for nm, ids in name_netids.items() if len(ids) == 1}
    iexp["Employee NetID"] = iexp.apply(
        lambda r: r["Employee NetID"] if pd.notna(r["Employee NetID"])
        else uniq_name_netid.get(r["Employee Name"], np.nan), axis=1)

    # ---- Drop 'Unknown User' placeholders, dedup Trx Id ----
    placeholder = iexp["Employee Name"].astype(str).str.contains("Unknown", case=False, na=False)
    iexp_clean = iexp[~placeholder].copy()
    iexp_clean = iexp_clean.drop_duplicates(subset="Trx Id", keep="first")

    # ---- Join on Trx Id (the process owner is master) ----
    iexp_cols = ["Trx Id", "Report Header Id", "Employee Name", "Employee NetID", "Employee Email",
                 "Expense Current Approver Name", "Expense Current Approver NetID", "Approver Email",
                 "Expense Status", "Creation Date", "Last Update Date", "Report Submitted Date"]
    iexp_cols = [c for c in iexp_cols if c in iexp_clean.columns]
    iexp_lookup = iexp_clean[iexp_cols].copy().rename(columns={
        "Employee Name": "iexp_Employee Name",
        "Expense Status": "iexp_Expense Status",
        "Creation Date": "iexp_Creation Date",
        "Last Update Date": "iexp_Last Update Date",
        "Report Header Id": "iexp_Report Header Id",
    })
    merged = tawnia.merge(iexp_lookup, on="Trx Id", how="left", indicator=True)

    # ---- STEP 5c: last-ditch NetID from Employee Data Dashboard (name -> NetID) ----
    dash_unique, _dash_ambig = load_emp_dashboard()
    if "Employee NetID" not in merged.columns:
        merged["Employee NetID"] = np.nan
    need = ~merged["Employee NetID"].apply(_netid_value_ok)
    for idx in merged[need].index:
        nm = str(merged.at[idx, "Card Holder Name"]).strip()
        nid = dash_unique.get(nm)
        if nid:
            merged.at[idx, "Employee NetID"] = nid

    # ---- Fiscal officer two-hop ----
    ira_lookup, poc_lookup = load_orgs()
    merged["code_norm"] = merged["GL Org"].apply(norm_code)
    merged = merged.merge(ira_lookup.rename(columns={"Group": "Default Division"}),
                          on="code_norm", how="left")
    merged = merged.merge(
        poc_lookup.rename(columns={
            "Division": "Default Division",
            "Fiscal Officer/Compliance POC": "Fiscal Officer",
            "Net ID": "fo_netid",
        }),
        on="Default Division", how="left")
    merged["Fiscal Officer Email"] = merged["fo_netid"].apply(build_fo_email)

    # ---- Classify every row ----
    merged[["final_status", "disposition"]] = merged.apply(classify, axis=1)
    send_df = merged[merged["disposition"] == "SEND"].copy()
    review_df = merged[merged["disposition"] == "REVIEW"].copy()

    # ---- Bucket ranking + numerics ----
    send_df["Bucket_clean"] = send_df["Bucket"].astype(str).str.strip()
    send_df["bucket_rank"] = send_df["Bucket_clean"].map(BUCKET_ORDER)
    if "Days Outstanding" in send_df.columns:
        send_df["Days Outstanding Num"] = pd.to_numeric(send_df["Days Outstanding"], errors="coerce")
    else:
        send_df["Days Outstanding Num"] = np.nan
    send_df["Billed Amount Num"] = pd.to_numeric(send_df["Billed Amount"], errors="coerce")

    # ---- Exclude Arts & Sciences (and a few named exclusions) ----
    # Belt & suspenders: the Division-lookup path AND the GL Org L1 marker that
    # lives directly in the owner's file (SAS-SCHOOL OF ARTS AND SCIENCES). Either one
    # marks a row as A&S, so nothing slips through if the lookup doesn't resolve.
    div_as = send_df["Default Division"].astype(str).str.strip().str.upper() == "ARTS AND SCIENCES"
    if "GL Org L1" in send_df.columns:
        l1_as = send_df["GL Org L1"].astype(str).str.contains("ARTS AND SCIENCES", case=False, na=False)
    else:
        l1_as = pd.Series(False, index=send_df.index)

    # Named individual exclusions (e.g. the President — not A&S by org, but excluded).
    NAME_EXCLUSIONS = {"Placeholder, Name X"}
    name_excl = send_df["Card Holder Name"].astype(str).str.strip().isin(NAME_EXCLUSIONS)

    is_as = div_as | l1_as | name_excl
    as_rows = send_df[is_as].copy()
    send_df = send_df[~is_as].copy()

    # ---- Scrub FAKE approver CCs (drop, don't flag) ----
    send_df["Approver Email"] = send_df["Approver Email"].where(
        ~send_df["Approver Email"].astype(str).str.contains("FAKE", case=False, na=False),
        np.nan,
    ) if "Approver Email" in send_df.columns else np.nan

    emailable_df = send_df.copy()

    # ---- Per-person summary (ALL emailable people; bucket filtering happens later) ----
    people = (emailable_df.groupby("Card Holder Name", group_keys=False)
              .apply(summarize_person, include_groups=False)
              .reset_index()
              .sort_values("total_billed", ascending=False)
              .reset_index(drop=True))

    return {
        "merged": merged,
        "emailable_df": emailable_df,
        "people": people,
        "review_df": review_df,
        "as_rows": as_rows,
        "counts": {
            "total_txns": len(tawnia),
            "all_people": len(people),
            "flagged": len(review_df),
            "as_people": int(as_rows["Card Holder Name"].nunique()) if len(as_rows) else 0,
        },
    }


def select_by_buckets(state, buckets):
    """Given the processed state and a set of bucket names, return the people to
    email (worst bucket in the set) plus all their items. Lets the UI change the
    bucket selection without re-running the whole pipeline."""
    buckets = set(buckets) if buckets else set(DEFAULT_BUCKETS)
    people = state["people"]
    chosen = people[people["worst_bucket"].isin(buckets)].copy().reset_index(drop=True)
    names = set(chosen["Card Holder Name"])
    items = state["emailable_df"][state["emailable_df"]["Card Holder Name"].isin(names)].copy()
    return chosen, items


# ---- Email building (ported) ----
def format_amount(v):
    try:
        return f"{float(v):,.2f}"
    except (ValueError, TypeError):
        return str(v)


def fmt_date(v):
    # Mac/Linux strip leading zeros with %-m; Windows uses %#m. Do it manually
    # so the same code runs on both without platform-specific format codes.
    try:
        d = pd.to_datetime(v)
        return f"{d.month}/{d.day}/{d.year}"
    except Exception:
        return str(v)


def age_phrase_from_bucket(bucket):
    b = str(bucket).strip()
    return {"> 150 Days": "150 days", "120 - 150 Days": "120 days",
            "90 - 120 Days": "90 days", "60 - 90 Days": "60 days",
            "30 - 60 Days": "30 days", "< 30 Days": "30 days"}.get(b, b)


def build_rows(items):
    out = []
    for _, r in items.iterrows():
        row = {}
        for col in TABLE_COLUMNS:
            val = r.get(col, "")
            if col == "Billed Amount":
                val = format_amount(val)
            elif col in ("Transaction Date", "Billed Date"):
                val = fmt_date(val)
            elif pd.isna(val):
                val = ""
            row[col] = str(val)
        out.append(row)
    return out


def build_email_for_person(name, items, person_row):
    bucket = str(person_row["worst_bucket"]).strip()
    statuses = set(items["final_status"].dropna().unique())
    has_pma = bool(statuses & MANAGER_FAULT)

    # Above the table: greeting + bucket intro + (optional) PMA note
    intro_paras = [GREETING, SUSPENSION_INTRO_BY_BUCKET.get(bucket, SUSPENSION_INTRO_DEFAULT)]
    if has_pma:
        intro_paras.append(PMA_INTRO)

    # Below the table: bucket policy paragraph + closing + signature.
    # All emails come from the process owner with her full signature.
    policy = POLICY_BODY_BY_BUCKET.get(bucket, POLICY_BODY_DEFAULT)
    footer = policy + "\n\n" + CLOSING_LINE + "\n\n" + SIGNATURE
    from_addr = ""   # empty = send from the default account (the process owner)

    ccs = []
    if person_row.get("approver_ccs"):
        ccs += [c.strip() for c in person_row["approver_ccs"].split(";") if c.strip()]
    # Fiscal officers are only copied for 60+ day buckets. Under 60 days
    # (< 30, 30-60), the FO is NOT involved.
    if not is_finance_center_bucket(bucket) and person_row.get("fiscal_officer_ccs"):
        ccs += [c.strip() for c in person_row["fiscal_officer_ccs"].split(";") if c.strip()]
    ccs += CC_ALWAYS
    ccs = sorted(set(ccs))

    return {
        "name": name,
        "to": person_row["Cardholder Email"],
        "cc": ccs,
        "subject": SUBJECT_BY_BUCKET.get(bucket, SUBJECT_DEFAULT),
        "from_addr": from_addr,
        "intro": "\n\n".join(intro_paras),
        "columns": list(TABLE_COLUMNS),
        "rows": build_rows(items),
        "footer": footer,
        "total": format_amount(person_row["total_billed"]),
        "dept": person_row["department"],
        "bucket": person_row["worst_bucket"],
    }


def build_all_emails(state, chosen_people, chosen_items):
    emails = []
    for _, prow in chosen_people.iterrows():
        name = prow["Card Holder Name"]
        items = chosen_items[chosen_items["Card Holder Name"] == name]
        emails.append(build_email_for_person(name, items, prow))
    return emails


# =============================================================
# SENDING — stubbed on Mac. This is the ONE place to wire Outlook.
# =============================================================
def render_plain(e):
    """Plain-text fallback body (used only if HTML somehow can't be set)."""
    lines = [e["intro"], ""]
    cols = e.get("columns") or (list(e["rows"][0].keys()) if e["rows"] else [])
    lines.append(" | ".join(cols))
    for r in e["rows"]:
        lines.append(" | ".join(str(r[c]) for c in cols))
    lines += ["", e["footer"]]
    return "\n".join(lines)


def _html_escape(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def _linkify(text_html):
    """Turn the policy phrases into <a> links inside already-escaped HTML text."""
    for phrase, url in LINK_PHRASES:
        text_html = text_html.replace(
            phrase, f'<a href="{url}">{phrase}</a>'
        )
    return text_html


def render_html(e):
    """Full HTML body: paragraphs, a real bordered table, and clickable policy links."""
    def paras(block):
        out = []
        for para in block.split("\n\n"):
            para_html = _linkify(_html_escape(para)).replace("\n", "<br>")
            out.append(f'<p style="margin:0 0 12px 0">{para_html}</p>')
        return "".join(out)

    cols = e.get("columns") or (list(e["rows"][0].keys()) if e["rows"] else [])
    th = "".join(
        f'<th style="border:1px solid #cccccc;padding:6px 10px;background:#00693e;'
        f'color:#ffffff;text-align:left;font-size:12px">{_html_escape(c)}</th>'
        for c in cols
    )
    trs = []
    for r in e["rows"]:
        tds = "".join(
            f'<td style="border:1px solid #cccccc;padding:6px 10px;font-size:12px">'
            f'{_html_escape(r.get(c, ""))}</td>'
            for c in cols
        )
        trs.append(f"<tr>{tds}</tr>")
    table = (
        '<table style="border-collapse:collapse;margin:14px 0;font-family:Arial,'
        'Helvetica,sans-serif">'
        f'<thead><tr>{th}</tr></thead><tbody>{"".join(trs)}</tbody></table>'
    )

    return (
        '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;'
        'color:#222222;line-height:1.5">'
        + paras(e["intro"]) + table + paras(e["footer"]) +
        "</div>"
    )


def _recipients_for_mode(e, mode):
    """Who actually receives this email, by mode.
    TEST -> the process owner only. DUMMY -> the demo user. PRODUCTION -> real cardholder + CCs."""
    if mode == "PRODUCTION":
        return e["to"], "; ".join(e["cc"])
    if mode == "TEST":
        return DUMMY_RECIPIENT_WINDOWS, DUMMY_RECIPIENT_WINDOWS   # the process owner
    # DUMMY
    who = DUMMY_RECIPIENT_MAC if not IS_WINDOWS else DUMMY_RECIPIENT_WINDOWS
    return who, who


def send_emails(emails, mode="TEST"):
    """
    All three modes REALLY send. They differ only in WHO receives:
      TEST       -> everything to the process owner (so she can review the real emails)
      DUMMY      -> everything to the demo user or the process owner (Windows)
      PRODUCTION -> real cardholders, with fiscal officer + manager CCs

    Mac sends through Apple Mail (renders HTML reliably); Windows through Outlook.
    Every run is logged to send_log.csv and an audit workbook.
    """
    if mode not in ("TEST", "DUMMY", "PRODUCTION"):
        mode = "TEST"

    outlook = None
    com_ready = False
    # DEMO BUILD: no mail client is ever contacted. In the real app this block
    # initializes Outlook (Windows) or Apple Mail (Mac); here it's a no-op so the
    # demo runs anywhere with no pywin32 / mail setup and can't send anything.

    try:
        results = []
        for e in emails:
            to_addr, cc_addr = _recipients_for_mode(e, mode)
            status, error = "", ""
            # Under-60 emails come from the Finance Center mailbox. Apply the
            # from-address for TEST (so the process owner sees the real sender) and PRODUCTION.
            # Not for DUMMY, which is purely a formatting check to one inbox.
            want_from = e.get("from_addr") if mode in ("TEST", "PRODUCTION") else ""
            try:
                # ---- DEMO BUILD: sending is mocked. Nothing is ever emailed. ----
                _ = render_html(e)  # still render, to prove the email builds
                # (real build would send via Outlook/Apple Mail here)
                status = {"TEST": "SENT — to the process owner (test)",
                          "DUMMY": "SENT — dummy copy",
                          "PRODUCTION": "SENT"}[mode]
            except Exception as ex:
                msg = str(ex)
                if want_from and ("denied" in msg.lower() or "permission" in msg.lower()
                                  or "SentOnBehalf" in msg or "-2147" in msg):
                    msg = (f"Could not send as {want_from} — the Outlook account needs "
                           f"Send-As permission on that mailbox. ({msg})")
                status, error = "FAILED", msg

            results.append({"name": e["name"], "to": to_addr, "cc": cc_addr,
                            "result": status, "error": error})
    finally:
        if IS_WINDOWS and com_ready:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # ---- Log every run ----
    log_df = pd.DataFrame([{
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "mode": mode,
        "name": r["name"], "to": r["to"], "cc": r["cc"],
        "result": r["result"], "error": r["error"],
    } for r in results])
    header = not os.path.exists(LOG_PATH)
    log_df.to_csv(LOG_PATH, mode="a", header=header, index=False)
    return results


def _send_via_mac_mail(to_addr, cc_addr, subject, html_body, from_addr=""):
    """Send one HTML email through Apple Mail (Mail.app) via AppleScript.

    Apple Mail has a real `html content` property that renders HTML reliably
    (Outlook for Mac's `content` does not). The HTML is passed via a temp file so
    quotes, ampersands (in URLs), and newlines can't break the script.

    If from_addr is given, the message's `sender` is set to it — this only works
    if that address is a configured account (or alias) in Mail.app on this Mac.

    Requires Mail.app configured with the user's account. First send triggers a
    macOS permission prompt ('Terminal wants to control Mail') — click Allow.
    """
    import subprocess, tempfile, os as _os

    def esc(s):
        return str(s).replace("\\", "\\\\").replace('"', '\\"')

    fd, html_path = tempfile.mkstemp(suffix=".html")
    with _os.fdopen(fd, "w", encoding="utf-8") as fh:
        fh.write(html_body)

    # Build recipient lines (cc may hold several "; "-separated addresses)
    cc_list = [c.strip() for c in str(cc_addr).split(";") if c.strip()]
    cc_lines = "\n".join(
        f'        make new cc recipient with properties {{address:"{esc(c)}"}}'
        for c in cc_list
    )
    sender_line = f'        set sender to "{esc(from_addr)}"\n' if from_addr else ""

    # Set `html content` directly in the message's creation properties — this is
    # the most reliable Apple Mail pattern (setting it afterward can race and send
    # an empty body). Recipients are added after, then send.
    script = f'''
set theFile to POSIX file "{esc(html_path)}"
set fileRef to open for access theFile
set htmlText to (read fileRef as «class utf8»)
close access fileRef
tell application "Mail"
    set newMsg to make new outgoing message with properties {{subject:"{esc(subject)}", content:htmlText, visible:true}}
    tell newMsg
        set html content to htmlText
{sender_line}        make new to recipient with properties {{address:"{esc(to_addr)}"}}
{cc_lines}
    end tell
    delay 0.4
    send newMsg
end tell
'''
    try:
        proc = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    finally:
        try:
            _os.remove(html_path)
        except OSError:
            pass
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or "osascript failed (is Mail.app set up with your account?)")


# =============================================================
# FLASK APP
# =============================================================
app = Flask(__name__)


# Make sure NaN/Infinity never leak into responses as invalid JSON tokens.
# (pandas produces NaN for blank cells; plain json renders that as literal `NaN`,
# which browsers can't parse.) This provider emits null instead.
try:
    from flask.json.provider import DefaultJSONProvider

    class _SafeJSON(DefaultJSONProvider):
        def dumps(self, obj, **kwargs):
            kwargs.setdefault("allow_nan", False)   # raise instead of writing NaN
            import math

            def _scrub(o):
                if isinstance(o, float) and (math.isnan(o) or math.isinf(o)):
                    return None
                if isinstance(o, dict):
                    return {k: _scrub(v) for k, v in o.items()}
                if isinstance(o, (list, tuple)):
                    return [_scrub(v) for v in o]
                return o
            return super().dumps(_scrub(obj), **kwargs)

    app.json = _SafeJSON(app)
except Exception:
    pass

STATE = {}  # holds the processed result between requests (single-user tool)


@app.route("/")
def index():
    return render_template_string(PAGE)


@app.route("/process", methods=["POST"])
def process():
    try:
        f1 = request.files.get("outstanding")
        f2 = request.files.get("iexpense")
        if not f1 or not f2:
            return jsonify({"ok": False, "error": "Please upload both files."}), 400

        # Which buckets did the process owner tick? (comma-separated). Default if none.
        buckets_raw = request.form.get("buckets", "")
        buckets = [b for b in buckets_raw.split("||") if b] or list(DEFAULT_BUCKETS)

        p1 = os.path.join(UPLOAD_DIR, "outstanding.xlsx")
        p2 = os.path.join(UPLOAD_DIR, "iexpense.csv")
        f1.save(p1)
        f2.save(p2)

        state = process_files(p1, p2)
        STATE["data"] = state
        payload = _apply_buckets(buckets)
        return jsonify({"ok": True, **payload})
    except Exception as ex:
        return jsonify({"ok": False, "error": str(ex),
                        "trace": traceback.format_exc()}), 500


def write_audit_workbook(state, chosen_people, chosen_items, mode, buckets):
    """Write a timestamped per-run audit workbook (never overwrites). 4 tabs:
    Emailed, Excluded (w/ reason), Flagged (w/ reason), All Transactions.
    Test runs get a _TEST suffix so they can't be mistaken for a real send."""
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    suffix = "_TEST" if mode == "TEST" else ("_DUMMY" if mode == "DUMMY" else "")
    fname = f"iExpense_Analysis_{ts}{suffix}.xlsx"
    path = os.path.join(APP_DIR, fname)

    merged = state["merged"]

    # --- Emailed tab: the people actually getting notices this run ---
    emailed_cols = {
        "Card Holder Name": "Cardholder", "Cardholder Email": "Cardholder Email",
        "department": "Department", "num_transactions": "# Items",
        "total_billed": "Total $", "worst_bucket": "Oldest Bucket",
        "statuses": "Statuses", "approver_ccs": "Manager CC",
        "fiscal_officer_ccs": "Fiscal Officer CC",
    }
    emailed = chosen_people[[c for c in emailed_cols if c in chosen_people.columns]].rename(columns=emailed_cols)

    # --- Excluded tab: everything dropped, with a reason ---
    excl = merged[merged["disposition"] == "EXCLUDE"].copy()
    excl_out = pd.DataFrame({
        "Trx Id": excl.get("Trx Id"),
        "Cardholder": excl.get("Card Holder Name"),
        "Department": excl.get("Department Name"),
        "Bucket": excl.get("Bucket"),
        "the process owner Status": excl.get("Expense Status"),
        "iExpense Status": excl.get("iexp_Expense Status"),
        "Reason (final status)": excl.get("final_status"),
        "Billed Amount": excl.get("Billed Amount"),
    })
    # Arts & Sciences are excluded separately (never entered EXCLUDE disposition)
    as_rows = state.get("as_rows")
    if as_rows is not None and len(as_rows):
        as_out = pd.DataFrame({
            "Trx Id": as_rows.get("Trx Id"),
            "Cardholder": as_rows.get("Card Holder Name"),
            "Department": as_rows.get("Department Name"),
            "Bucket": as_rows.get("Bucket"),
            "the process owner Status": as_rows.get("Expense Status"),
            "iExpense Status": as_rows.get("iexp_Expense Status"),
            "Reason (final status)": "Excluded — Arts & Sciences",
            "Billed Amount": as_rows.get("Billed Amount"),
        })
        excl_out = pd.concat([excl_out, as_out], ignore_index=True)

    # --- Flagged tab: review pile, with reason ---
    rev = state["review_df"].copy()
    flagged_out = pd.DataFrame({
        "Trx Id": rev.get("Trx Id"),
        "Cardholder": rev.get("Card Holder Name"),
        "Department": rev.get("Department Name"),
        "Bucket": rev.get("Bucket"),
        "the process owner Status": rev.get("Expense Status"),
        "iExpense Status": rev.get("iexp_Expense Status"),
        "Reason (final status)": rev.get("final_status"),
        "Billed Amount": rev.get("Billed Amount"),
    })

    # --- All Transactions tab: full merged set with disposition ---
    all_cols = ["Trx Id", "Card Holder Name", "Department Name", "Report Header Id",
                "Bucket", "Expense Status", "iexp_Expense Status", "Employee NetID",
                "final_status", "disposition", "Billed Amount"]
    all_out = merged[[c for c in all_cols if c in merged.columns]].copy()

    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        emailed.to_excel(xw, sheet_name="Emailed", index=False)
        excl_out.to_excel(xw, sheet_name="Excluded", index=False)
        flagged_out.to_excel(xw, sheet_name="Flagged for Review", index=False)
        all_out.to_excel(xw, sheet_name="All Transactions", index=False)
    return fname


def _apply_buckets(buckets):
    """Filter the already-processed data to the chosen buckets and rebuild emails.
    Returns the counts + review rows for the UI."""
    state = STATE["data"]
    chosen_people, chosen_items = select_by_buckets(state, buckets)
    STATE["emails"] = build_all_emails(state, chosen_people, chosen_items)
    STATE["buckets"] = buckets

    def _clean(v):
        """NaN/None -> '' so the JSON is always valid (NaN is not legal JSON)."""
        if v is None:
            return ""
        try:
            if pd.isna(v):
                return ""
        except (TypeError, ValueError):
            pass
        return v

    review = [{
        "name": _clean(r["Card Holder Name"]), "dept": _clean(r["department"]),
        "items": int(r["num_transactions"]),
        "amount": format_amount(r["total_billed"]),
        "bucket": _clean(r["worst_bucket"]), "statuses": _clean(r["statuses"]),
    } for _, r in chosen_people.iterrows()]

    counts = dict(state["counts"])
    counts["selected_people"] = len(chosen_people)
    counts["selected_dollars"] = round(chosen_people["total_billed"].sum(), 2)
    counts["buckets"] = buckets

    # Flagged-for-review transactions, with reason — shown directly in the app so
    # nobody has to open the Excel to see who was held back and why.
    rev = state["review_df"]
    flagged = [{
        "name": _clean(r.get("Card Holder Name", "")),
        "dept": _clean(r.get("Department Name", "")),
        "bucket": str(_clean(r.get("Bucket", ""))).strip(),
        "trx": _clean(r.get("Trx Id", "")),
        "tawnia_status": _clean(r.get("Expense Status", "")),
        "iexp_status": _clean(r.get("iexp_Expense Status", "")),
        "reason": _clean(r.get("final_status", "")),
        "amount": format_amount(pd.to_numeric(r.get("Billed Amount"), errors="coerce")),
    } for _, r in rev.iterrows()]

    return {"counts": counts, "review": review, "flagged": flagged}


@app.route("/rebucket", methods=["POST"])
def rebucket():
    """Re-filter to a new bucket selection without re-uploading (instant)."""
    if "data" not in STATE:
        return jsonify({"ok": False, "error": "Upload and process files first."}), 400
    buckets = (request.json or {}).get("buckets") or list(DEFAULT_BUCKETS)
    payload = _apply_buckets(buckets)
    return jsonify({"ok": True, **payload})


@app.route("/emails")
def get_emails():
    return jsonify({"ok": True, "emails": STATE.get("emails", [])})


@app.route("/send", methods=["POST"])
def send():
    emails = STATE.get("emails", [])
    if not emails:
        return jsonify({"ok": False, "error": "Nothing to send — process files first."}), 400
    mode = (request.json or {}).get("mode", "TEST")
    if mode not in ("TEST", "DUMMY", "PRODUCTION"):
        mode = "TEST"
    try:
        results = send_emails(emails, mode=mode)
    except Exception as ex:
        return jsonify({"ok": False, "error": str(ex)}), 500

    # ---- Write a timestamped audit workbook for THIS run (never overwrites) ----
    audit_file = None
    try:
        state = STATE["data"]
        chosen_people, chosen_items = select_by_buckets(state, STATE.get("buckets"))
        audit_file = write_audit_workbook(state, chosen_people, chosen_items, mode, STATE.get("buckets"))
    except Exception as ex:
        audit_file = f"(audit workbook failed: {ex})"

    sent_ok = sum(1 for r in results if r["result"].startswith("SENT"))
    failed = sum(1 for r in results if r["result"] == "FAILED")
    return jsonify({"ok": True, "mode": mode, "count": len(results),
                    "sent_ok": sent_ok, "failed": failed,
                    "results": results, "log": os.path.basename(LOG_PATH),
                    "audit_file": audit_file})


@app.route("/platform")
def platform_info():
    return jsonify({"is_windows": IS_WINDOWS, "system": platform.system()})


@app.route("/preview_browser", methods=["POST"])
def preview_browser():
    """Write the FIRST selected email's exact HTML to a file and open it in the
    default browser. This is a can't-fail way to see the real formatting — no
    Mail/Outlook involved, just a file on disk."""
    emails = STATE.get("emails", [])
    if not emails:
        return jsonify({"ok": False, "error": "Nothing to preview — process files first."}), 400
    e = emails[0]
    html = (
        "<!DOCTYPE html><html><head><meta charset='utf-8'>"
        "<title>Email preview — " + _html_escape(e["name"]) + "</title></head>"
        "<body style='margin:24px;background:#f4f5f3'>"
        "<div style='max-width:820px;margin:0 auto;background:#fff;padding:26px 30px;"
        "border:1px solid #d9e0db;border-radius:10px'>"
        "<div style='font-family:Arial;font-size:12px;color:#555;"
        "border-bottom:1px solid #eee;padding-bottom:10px;margin-bottom:16px'>"
        "<b>FROM:</b> " + _html_escape(e.get("from_addr") or "Finance Operations (default)") + "<br>"
        "<b>TO:</b> " + _html_escape(e["to"]) + "<br>"
        "<b>CC:</b> " + _html_escape(", ".join(e["cc"])) + "<br>"
        "<b>SUBJECT:</b> " + _html_escape(e["subject"]) +
        "</div>" + render_html(e) + "</div></body></html>"
    )
    out_path = os.path.join(APP_DIR, "email_preview.html")
    with open(out_path, "w", encoding="utf-8") as fh:
        fh.write(html)

    # Open in default browser (cross-platform)
    opened = True
    try:
        if IS_WINDOWS:
            os.startfile(out_path)  # noqa
        elif platform.system() == "Darwin":
            import subprocess
            subprocess.run(["open", out_path])
        else:
            import webbrowser
            webbrowser.open("file://" + out_path)
    except Exception:
        opened = False
    return jsonify({"ok": True, "file": os.path.basename(out_path),
                    "path": out_path, "opened": opened, "name": e["name"]})


# =============================================================
# PAGE (dashboard) — served inline so it's one file
# =============================================================
PAGE = r"""
<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>iExpense Compliance Automation Suite</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600&family=Newsreader:wght@400;500&family=IBM+Plex+Mono:wght@400;500&display=swap');
  :root{--green:#00693e;--green-deep:#0a4d31;--green-soft:#e7f1ec;--ink:#1a2420;--ink-soft:#4a5853;--paper:#fbfaf6;--card:#fff;--line:#d9e0db;--amber:#b8741a;--amber-soft:#fbf0e0;--red:#9e2b25;--shadow:0 1px 2px rgba(10,40,28,.06),0 8px 28px rgba(10,40,28,.07)}
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Newsreader',Georgia,serif;background:var(--paper);color:var(--ink);line-height:1.55}
  .wrap{max-width:1080px;margin:0 auto;padding:0 28px 80px}
  header.mast{border-bottom:2px solid var(--green);padding:30px 0 22px}
  .mast-row{display:flex;align-items:baseline;justify-content:space-between;gap:20px;flex-wrap:wrap}
  .mast h1{font-family:'Fraunces',serif;font-weight:600;font-size:30px;color:var(--green-deep)}
  .mast .sub{font-family:'IBM Plex Mono',monospace;font-size:11px;text-transform:uppercase;letter-spacing:.18em;color:var(--ink-soft)}
  .flag{font-family:'IBM Plex Mono',monospace;font-size:10.5px;letter-spacing:.14em;text-transform:uppercase;background:var(--amber-soft);color:var(--amber);border:1px solid #e9d3ad;padding:5px 11px;border-radius:100px}
  .stepper{display:flex;margin:26px 0 30px;border:1px solid var(--line);border-radius:12px;overflow:hidden;background:var(--card);box-shadow:var(--shadow)}
  .step{flex:1;padding:16px 18px;border-right:1px solid var(--line)}
  .step:last-child{border-right:none}
  .step.active{background:var(--green)}
  .step .n{font-family:'IBM Plex Mono',monospace;font-size:11px;color:var(--ink-soft)}
  .step.active .n{color:rgba(255,255,255,.7)}
  .step .t{font-family:'Fraunces',serif;font-size:16px;font-weight:500;margin-top:3px}
  .step.active .t{color:#fff}
  .panel{display:none;animation:rise .4s ease both}.panel.show{display:block}
  @keyframes rise{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}
  .card{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:28px 30px;box-shadow:var(--shadow);margin-bottom:22px}
  .card h2{font-family:'Fraunces',serif;font-weight:600;font-size:21px;color:var(--green-deep);margin-bottom:6px}
  .card p.lead{color:var(--ink-soft);font-size:16px;margin-bottom:18px}
  .drops{display:grid;grid-template-columns:1fr 1fr;gap:16px}
  .bucketpick{margin-top:20px;padding:18px 20px;border:1px solid var(--line);border-radius:12px;background:#fcfdfc}
  .bucketpick .bpt{font-family:'Fraunces',serif;font-weight:600;font-size:16px;color:var(--green-deep)}
  .bucketpick .bpd{font-size:13px;color:var(--ink-soft);margin:4px 0 12px}
  .bpgrid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}
  @media(max-width:720px){.bpgrid{grid-template-columns:1fr 1fr}}
  .bpgrid label{display:flex;align-items:center;gap:8px;font-size:14px;padding:8px 10px;border:1px solid var(--line);border-radius:8px;background:#fff;cursor:pointer}
  .bpgrid label:hover{background:var(--green-soft)}
  .bchk{width:15px;height:15px;accent-color:var(--green);cursor:pointer}
  .bchip{display:inline-block;background:var(--green-soft);color:var(--green-deep);border:1px solid #bcd6c8;border-radius:100px;padding:2px 10px;font-size:12px;font-family:'IBM Plex Mono',monospace;margin-right:4px}
  @media(max-width:720px){.drops{grid-template-columns:1fr}.stepper{flex-wrap:wrap}.step{min-width:50%}}
  .drop{border:1.5px dashed #b9c7bf;border-radius:12px;padding:26px 22px;text-align:center;background:#fcfdfc;position:relative}
  .drop.loaded{border-style:solid;border-color:var(--green);background:var(--green-soft)}
  .drop .ico{width:42px;height:42px;margin:0 auto 12px;border-radius:10px;background:var(--green-soft);display:grid;place-items:center;color:var(--green);font-family:'IBM Plex Mono',monospace}
  .drop.loaded .ico{background:var(--green);color:#fff}
  .drop .dt{font-family:'Fraunces',serif;font-size:16px;font-weight:500}
  .drop .dd{font-size:13.5px;color:var(--ink-soft);margin-top:4px}
  .drop input{position:absolute;inset:0;opacity:0;cursor:pointer}
  .filechip{display:none;margin-top:12px;font-family:'IBM Plex Mono',monospace;font-size:12px;color:var(--green-deep);background:#fff;border:1px solid var(--line);border-radius:8px;padding:7px 10px}
  .drop.loaded .filechip{display:inline-block}
  .btn{font-family:'IBM Plex Mono',monospace;font-size:12.5px;border:none;border-radius:9px;padding:13px 22px;cursor:pointer;transition:all .15s}
  .btn-primary{background:var(--green);color:#fff}.btn-primary:hover{background:var(--green-deep)}
  .btn-ghost{background:transparent;color:var(--green-deep);border:1px solid var(--line)}.btn-ghost:hover{background:var(--green-soft)}
  .btn-danger{background:var(--red);color:#fff}
  .btn:disabled{opacity:.4;cursor:not-allowed}
  .btn-row{display:flex;gap:12px;margin-top:22px;flex-wrap:wrap;align-items:center}
  .tiles{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:22px}
  @media(max-width:720px){.tiles{grid-template-columns:1fr 1fr}}
  .tile{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:18px;box-shadow:var(--shadow)}
  .tile .v{font-family:'Fraunces',serif;font-size:30px;font-weight:600;color:var(--green-deep);line-height:1}
  .tile .l{font-size:13px;color:var(--ink-soft);margin-top:7px}
  .tile.warn .v{color:var(--amber)}.tile.mute .v{color:var(--ink-soft)}
  .tbl-wrap{border:1px solid var(--line);border-radius:12px;overflow:hidden;background:var(--card);box-shadow:var(--shadow)}
  table{width:100%;border-collapse:collapse;font-size:14px}
  thead th{background:var(--green-deep);color:#fff;text-align:left;font-family:'IBM Plex Mono',monospace;font-weight:500;font-size:11px;letter-spacing:.06em;text-transform:uppercase;padding:12px 14px}
  tbody td{padding:11px 14px;border-top:1px solid var(--line)}
  tbody tr:hover{background:var(--green-soft)}
  .amt{font-family:'IBM Plex Mono',monospace;font-weight:500}
  .mailbox{display:grid;grid-template-columns:240px 1fr;border:1px solid var(--line);border-radius:12px;overflow:hidden;background:var(--card);box-shadow:var(--shadow)}
  @media(max-width:720px){.mailbox{grid-template-columns:1fr}}
  .maillist{border-right:1px solid var(--line);max-height:520px;overflow-y:auto}
  .mailitem{padding:13px 15px;border-bottom:1px solid var(--line);cursor:pointer}
  .mailitem:hover{background:var(--green-soft)}
  .mailitem.sel{background:var(--green-soft);border-left:3px solid var(--green)}
  .mailitem .mn{font-family:'Fraunces',serif;font-weight:500;font-size:15px}
  .mailitem .md{font-size:12px;color:var(--ink-soft);margin-top:2px}
  .mailview{padding:22px 24px;max-height:520px;overflow-y:auto}
  .mailmeta{font-size:13.5px;margin-bottom:4px}
  .mailmeta b{font-family:'IBM Plex Mono',monospace;font-weight:500;font-size:12px;color:var(--ink-soft)}
  .mailmeta .val{font-family:'IBM Plex Mono',monospace;font-size:12.5px;color:var(--green-deep)}
  .mailbody{font-size:14px;white-space:pre-wrap;margin-top:8px}
  .mailbody table{font-size:11px;margin:12px 0}
  .mailbody thead th{font-size:9px;padding:6px 7px}
  .mailbody tbody td{padding:5px 7px}
  .notice{border-radius:10px;padding:14px 16px;font-size:14px;display:flex;gap:11px;margin-bottom:16px}
  .notice.green{background:var(--green-soft);border:1px solid #bcd6c8;color:var(--green-deep)}
  .notice.amber{background:var(--amber-soft);border:1px solid #e9d3ad;color:#7a4d12}
  .notice.red{background:#f7e7e5;border:1px solid #e3b5b0;color:var(--red)}
  .spinner{display:inline-block;width:16px;height:16px;border:2px solid var(--green-soft);border-top-color:var(--green);border-radius:50%;animation:spin .7s linear infinite;vertical-align:middle}
  @keyframes spin{to{transform:rotate(360deg)}}
  .sendchoice{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin:8px 0}
  @media(max-width:720px){.sendchoice{grid-template-columns:1fr}}
  .choicecard{border:1px solid var(--line);border-radius:12px;padding:20px 22px;background:#fcfdfc}
  .choicecard.real{border-color:#e3b5b0;background:#fdf6f5}
  .choicecard.dummy{border-color:#b9c9dd;background:#f5f8fc}
  .choicecard .ct{font-family:'Fraunces',serif;font-size:17px;font-weight:600;color:var(--green-deep);margin-bottom:6px}
  .choicecard.real .ct{color:var(--red)}
  .choicecard .cd{font-size:13.5px;color:var(--ink-soft);margin-bottom:16px;line-height:1.5}
  .modal-bg{display:none;position:fixed;inset:0;background:rgba(10,40,28,.4);backdrop-filter:blur(3px);z-index:50;place-items:center;padding:20px}
  .modal-bg.show{display:grid}
  .modal{background:var(--card);border-radius:16px;padding:30px 32px;max-width:440px;box-shadow:0 24px 60px rgba(10,40,28,.3);text-align:center}
  .modal h3{font-family:'Fraunces',serif;font-size:22px;color:var(--red);margin-bottom:8px}
  .modal p{color:var(--ink-soft);font-size:15px;margin-bottom:22px}
  .modal .big{font-family:'Fraunces',serif;font-size:40px;color:var(--red);font-weight:600}
  footer{text-align:center;color:var(--ink-soft);font-size:12.5px;margin-top:40px;font-family:'IBM Plex Mono',monospace}
</style></head><body><div class="wrap">
  <header class="mast"><div class="mast-row">
    <div><h1>iExpense Compliance Automation Suite</h1><div class="sub">Finance Operations · Corporate Card Compliance</div></div>
    <span class="flag" id="modeflag">● Checking environment…</span>
  </div></header>

  <div class="stepper" id="stepper">
    <div class="step active" data-p="1"><div class="n">STEP 01</div><div class="t">Upload Files</div></div>
    <div class="step" data-p="2"><div class="n">STEP 02</div><div class="t">Review Summary</div></div>
    <div class="step" data-p="3"><div class="n">STEP 03</div><div class="t">Preview Emails</div></div>
    <div class="step" data-p="4"><div class="n">STEP 04</div><div class="t">Send</div></div>
  </div>

  <!-- STEP 1 -->
  <section class="panel show" data-panel="1"><div class="card">
    <h2>Upload this month's two files</h2>
    <p class="lead">Drop in the Outstanding CC Transactions file and the iExpense data export. The fiscal-officer reference file is already built in.</p>
    <div class="drops">
      <label class="drop" id="drop1"><div class="ico">XLS</div><div class="dt">Outstanding CC Transactions</div><div class="dd">The file you prepare each month</div><div class="filechip" id="chip1"></div><input type="file" id="file1" accept=".xlsx,.xls"></label>
      <label class="drop" id="drop2"><div class="ico">CSV</div><div class="dt">iExpense Data Export</div><div class="dd">The OAC export with NetIDs &amp; approvers</div><div class="filechip" id="chip2"></div><input type="file" id="file2" accept=".csv"></label>
    </div>

    <div class="bucketpick">
      <div class="bpt">Which aging groups should get emails?</div>
      <div class="bpd">Tick the buckets to include. Anyone whose <b>oldest</b> item falls in a ticked bucket gets an email listing all their items. You can change this after processing too.</div>
      <div class="bpgrid" id="bucketGrid">
        <label><input type="checkbox" class="bchk" value="< 30 Days"> &lt; 30 Days</label>
        <label><input type="checkbox" class="bchk" value="30 - 60 Days"> 30 – 60 Days</label>
        <label><input type="checkbox" class="bchk" value="60 - 90 Days"> 60 – 90 Days</label>
        <label><input type="checkbox" class="bchk" value="90 - 120 Days"> 90 – 120 Days</label>
        <label><input type="checkbox" class="bchk" value="120 - 150 Days" checked> 120 – 150 Days</label>
        <label><input type="checkbox" class="bchk" value="> 150 Days" checked> &gt; 150 Days</label>
      </div>
    </div>

    <div class="btn-row">
      <button class="btn btn-primary" id="processBtn" disabled onclick="doProcess()">Process Files →</button>
      <span id="procStatus" style="font-size:13px;color:var(--ink-soft)">Upload both files to continue. Nothing is sent at this stage.</span>
    </div>
    <div id="procError"></div>
  </div></section>

  <!-- STEP 2 -->
  <section class="panel" data-panel="2">
    <div class="tiles" id="tiles"></div>
    <div class="notice green"><span>✓</span><span><b>Processed on your real files.</b> Everyone below has a valid cardholder email and fiscal-officer copy. Arts &amp; Sciences cardholders were left out automatically.</span></div>
    <div class="notice amber" id="flagNotice"><span>▲</span><span></span></div>
    <div class="card" id="flaggedCard" style="display:none">
      <h2>Flagged for review — not emailed</h2>
      <p class="lead">These transactions were held back and did <b>not</b> get an email. Two reasons only: the transaction is <b>Pending System Administrator Action</b>, or no NetID could be found for the cardholder (so there's no address). Handle these manually.</p>
      <div class="tbl-wrap"><table><thead><tr><th>Cardholder</th><th>Department</th><th>Trx Id</th><th>Bucket</th><th>Reason</th><th>Amount</th></tr></thead><tbody id="flaggedBody"></tbody></table></div>
    </div>
    <div class="card">
      <h2>Who will receive a notice</h2>
      <p class="lead" id="reviewBuckets"></p>
      <div class="tbl-wrap"><table><thead><tr><th>Cardholder</th><th>Department</th><th>Items</th><th>Amount</th><th>Oldest</th><th>Statuses</th></tr></thead><tbody id="reviewBody"></tbody></table></div>
      <div class="btn-row"><button class="btn btn-ghost" onclick="go(1)">← Back</button><button class="btn btn-primary" onclick="loadEmails();go(3)">Preview the Emails →</button></div>
    </div>
  </section>

  <!-- STEP 3 -->
  <section class="panel" data-panel="3"><div class="card">
    <h2>Read each email before anything goes out</h2>
    <p class="lead">Every notice is built from your files, in your wording. Click a name to read theirs. Nothing has been sent.</p>
    <div class="mailbox"><div class="maillist" id="maillist"></div><div class="mailview" id="mailview"></div></div>
    <div class="btn-row"><button class="btn btn-ghost" onclick="go(2)">← Back</button><button class="btn btn-ghost" onclick="previewBrowser()">Preview in Browser</button><button class="btn btn-primary" onclick="go(4)">Looks Good — Continue →</button></div>
    <div id="previewNote" style="font-size:13px;color:var(--ink-soft);margin-top:10px"></div>
  </div></section>

  <!-- STEP 4 -->
  <section class="panel" data-panel="4"><div class="card">
    <h2>Send the notices</h2>
    <p class="lead">Choose how to run this. You always pick consciously — nothing sends on its own.</p>
    <div class="sendchoice">
      <div class="choicecard">
        <div class="ct">Test</div>
        <div class="cd">Really sends the emails, but every message goes only to <b>the process owner</b> so she can review the real formatting. No cardholder is contacted.</div>
        <button class="btn btn-ghost" id="testBtn" onclick="doSend('TEST')">Test</button>
      </div>
      <div class="choicecard dummy" id="dummyCard">
        <div class="ct">Dummy Test</div>
        <div class="cd">Really sends the emails, but every message goes only to <b id="dummyWho">you</b>. Lets you eyeball the true output. No cardholder is contacted.</div>
        <button class="btn btn-ghost" id="dummyBtn" onclick="doSend('DUMMY')">Dummy Test</button>
      </div>
      <div class="choicecard real">
        <div class="ct">Production</div>
        <div class="cd">Actually emails every cardholder, copying fiscal officers &amp; managers. This cannot be undone.</div>
        <button class="btn btn-danger" id="sendBtn" onclick="confirmReal()">Production</button>
      </div>
    </div>
    <div id="platformNote" style="font-size:13px;color:var(--ink-soft);margin-top:14px"></div>
    <div class="btn-row"><button class="btn btn-ghost" onclick="go(3)">← Back</button><span id="sendStatus" style="font-size:13px;color:var(--ink-soft)"></span></div>
    <div id="sendResult"></div>
  </div></section>

  <!-- Confirm modal for real send -->
  <div class="modal-bg" id="modal">
    <div class="modal">
      <h3>Run in Production?</h3>
      <div class="big" id="modalCount">0</div>
      <p>emails will be sent through Outlook right now. Fiscal officers and managers will be copied. This cannot be undone.</p>
      <div style="display:flex;gap:12px;justify-content:center">
        <button class="btn btn-ghost" onclick="closeModal()">Cancel</button>
        <button class="btn btn-danger" onclick="closeModal();doSend('PRODUCTION')">Yes, send now</button>
      </div>
    </div>
  </div>

  <footer>iExpense Compliance Automation Suite · Built by Yeshwanth Lolla, Financial Systems &amp; Services · Questions or edits? Reach out anytime.</footer>
</div>
<script>
let EMAILS=[];
function go(p){document.querySelectorAll('.panel').forEach(x=>x.classList.remove('show'));document.querySelector('.panel[data-panel="'+p+'"]').classList.add('show');document.querySelectorAll('.step').forEach(s=>s.classList.toggle('active',s.dataset.p==p));window.scrollTo({top:0,behavior:'smooth'});}
document.querySelectorAll('.step').forEach(s=>s.onclick=()=>go(s.dataset.p));

function hookFile(n){const inp=document.getElementById('file'+n);inp.addEventListener('change',()=>{if(inp.files.length){document.getElementById('drop'+n).classList.add('loaded');document.getElementById('chip'+n).textContent='✓ '+inp.files[0].name;}checkReady();});}
hookFile(1);hookFile(2);
function checkReady(){document.getElementById('processBtn').disabled=!(document.getElementById('file1').files.length&&document.getElementById('file2').files.length);}

function getBuckets(){return Array.from(document.querySelectorAll('.bchk:checked')).map(c=>c.value);}

function doProcess(){
  const buckets=getBuckets();
  if(buckets.length===0){document.getElementById('procError').innerHTML='<div class="notice amber"><span>▲</span><span>Please tick at least one aging group to email.</span></div>';return;}
  const fd=new FormData();
  fd.append('outstanding',document.getElementById('file1').files[0]);
  fd.append('iexpense',document.getElementById('file2').files[0]);
  fd.append('buckets',buckets.join('||'));
  document.getElementById('procStatus').innerHTML='<span class="spinner"></span> Processing… please allow a few moments.';
  document.getElementById('procError').innerHTML='';
  fetch('/process',{method:'POST',body:fd}).then(r=>r.json()).then(d=>{
    if(!d.ok){document.getElementById('procStatus').textContent='';document.getElementById('procError').innerHTML='<div class="notice red"><span>✕</span><span><b>Could not process.</b> '+d.error+'</span></div>';return;}
    renderResults(d);
    document.getElementById('procStatus').innerHTML='✓ Done — '+d.counts.selected_people+' people selected.';
    go(2);
  }).catch(e=>{document.getElementById('procStatus').textContent='';document.getElementById('procError').innerHTML='<div class="notice red"><span>'+'\u2715'+'</span><span>'+e+'</span></div>';});
}

function renderResults(d){
  const c=d.counts;
  document.getElementById('tiles').innerHTML=
    tile(c.selected_people,'People to be emailed')+
    tile('$'+Math.round(c.selected_dollars).toLocaleString(),'Total outstanding (selected)')+
    tile(c.flagged,'Flagged for review','warn')+
    tile(c.as_people,'Arts &amp; Sciences (excluded)','mute');
  document.querySelector('#flagNotice span:last-child').innerHTML='<b>'+c.flagged+' transaction(s) flagged for review</b> — held back and not emailed. See the list below for who and why.';
  document.getElementById('reviewBody').innerHTML=d.review.map(r=>'<tr><td>'+r.name+'</td><td>'+r.dept+'</td><td>'+r.items+'</td><td class="amt">$'+r.amount+'</td><td>'+r.bucket+'</td><td style="font-size:12px">'+r.statuses+'</td></tr>').join('');
  // flagged-for-review list (shown right in the app, with reason)
  const fb=document.getElementById('flaggedBody'); const fc=document.getElementById('flaggedCard');
  const fl=d.flagged||[];
  if(fl.length){fc.style.display='';fb.innerHTML=fl.map(r=>'<tr><td>'+r.name+'</td><td>'+r.dept+'</td><td style="font-size:12px">'+r.trx+'</td><td>'+r.bucket+'</td><td style="font-size:12px"><b>'+r.reason+'</b></td><td class="amt">$'+r.amount+'</td></tr>').join('');}
  else{fc.style.display='none';}
  // reflect selected buckets on the review screen chips
  const chips=(c.buckets||[]).map(b=>'<span class="bchip">'+b+'</span>').join(' ');
  const rb=document.getElementById('reviewBuckets'); if(rb) rb.innerHTML='Currently emailing: '+chips+' &nbsp; <a href="#" onclick="go(1);return false;" style="color:var(--green-deep)">change</a>';
}
function tile(v,l,cls){return '<div class="tile '+(cls||'')+'"><div class="v">'+v+'</div><div class="l">'+l+'</div></div>';}

function loadEmails(){fetch('/emails').then(r=>r.json()).then(d=>{EMAILS=d.emails;const list=document.getElementById('maillist');list.innerHTML=EMAILS.map((e,i)=>'<div class="mailitem'+(i==0?' sel':'')+'" onclick="showMail('+i+',this)"><div class="mn">'+e.name+'</div><div class="md">$'+e.total+' · '+e.rows.length+' items</div></div>').join('');if(EMAILS.length)showMail(0,list.firstChild);});}
function previewBrowser(){const n=document.getElementById('previewNote');n.textContent='Opening preview…';fetch('/preview_browser',{method:'POST'}).then(r=>r.json()).then(d=>{if(!d.ok){n.innerHTML='<span style="color:var(--red)">'+d.error+'</span>';return;}if(d.opened){n.innerHTML='✓ Opened <b>'+d.name+'</b> in your browser (saved as '+d.file+' in the app folder).';}else{n.innerHTML='Saved as <b>'+d.file+'</b> in the app folder — open it manually. Path: '+d.path;}});}
function showMail(i,el){document.querySelectorAll('.mailitem').forEach(m=>m.classList.remove('sel'));if(el)el.classList.add('sel');const e=EMAILS[i];let cols=(e.columns&&e.columns.length)?e.columns:Object.keys(e.rows[0]||{});const esc=s=>String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');let th=cols.map(c=>'<th>'+esc(c)+'</th>').join('');let tr=e.rows.map(r=>'<tr>'+cols.map(c=>'<td>'+esc(r[c])+'</td>').join('')+'</tr>').join('');const intro=esc(e.intro).replace(/\n/g,'<br>');const footer=esc(e.footer).replace(/\n/g,'<br>');document.getElementById('mailview').innerHTML='<div class="mailmeta"><b>TO&nbsp;&nbsp;</b> <span class="val">'+esc(e.to)+'</span></div><div class="mailmeta"><b>CC&nbsp;&nbsp;</b> <span class="val">'+(e.cc.map(esc).join(', ')||'—')+'</span></div><div class="mailmeta"><b>SUBJ</b> <span class="val">'+esc(e.subject)+'</span></div><hr style="border:none;border-top:1px solid var(--line);margin:14px 0"><div class="mailbody">'+intro+'<table><thead><tr>'+th+'</tr></thead><tbody>'+tr+'</tbody></table>'+footer+'</div>';}

// detect platform to guide the user
fetch('/platform').then(r=>r.json()).then(p=>{
  const note=document.getElementById('platformNote');
  const flag=document.getElementById('modeflag');
  const dummyWho=document.getElementById('dummyWho');
  if(p.is_windows){
    // Windows: Test + Production (no Dummy card)
    const dc=document.getElementById('dummyCard'); if(dc) dc.style.display='none';
    if(dummyWho) dummyWho.textContent='the process owner';
    note.innerHTML='● Running on <b>Windows</b> (Outlook). Test sends to the process owner; Production sends to all cardholders.';
    flag.textContent='● Windows — Outlook';
  } else {
    // Mac: Test + Dummy + Production, all via Apple Mail
    if(dummyWho) dummyWho.textContent='you (demo.user@example.edu)';
    note.innerHTML='● Running on <b>'+p.system+'</b> (Apple Mail). Test → the process owner, Dummy → you, Production → all cardholders.';
    flag.textContent='● '+p.system+' — Apple Mail';
  }
});
function openModal(){document.getElementById('modalCount').textContent=EMAILS.length;document.getElementById('modal').classList.add('show');}
function closeModal(){document.getElementById('modal').classList.remove('show');}
function confirmReal(){openModal();}
function doSend(mode){
  document.getElementById('testBtn').disabled=true;
  const db=document.getElementById('dummyBtn'); if(db) db.disabled=true;
  const sb=document.getElementById('sendBtn'); if(sb) sb.disabled=true;
  const label = mode==='PRODUCTION'?'Sending to all cardholders…':(mode==='DUMMY'?'Sending dummy copies…':'Sending test copies to the process owner…');
  document.getElementById('sendStatus').innerHTML='<span class="spinner"></span> '+label;
  document.getElementById('sendResult').innerHTML='';
  fetch('/send',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({mode:mode})}).then(r=>r.json()).then(d=>{
    document.getElementById('testBtn').disabled=false;
    if(db) db.disabled=false;
    if(sb) sb.disabled=false;
    document.getElementById('sendStatus').textContent='';
    if(!d.ok){document.getElementById('sendResult').innerHTML='<div class="notice red"><span>✕</span><span>'+d.error+'</span></div>';return;}
    const audit = d.audit_file?(' Audit workbook: <b>'+d.audit_file+'</b>.'):'';
    let msg;
    if(d.mode==='TEST'){
      msg='<b>Test complete — '+d.sent_ok+' of '+d.count+' emails sent to the process owner for review.</b> No cardholder was contacted.';
    } else if(d.mode==='DUMMY'){
      msg='<b>Dummy Test complete — '+d.sent_ok+' of '+d.count+' copies sent to you.</b> No cardholder was contacted.';
    } else {
      msg='<b>Production — sent '+d.sent_ok+' of '+d.count+' emails to cardholders.</b>';
    }
    if(d.failed>0) msg+=' '+d.failed+' failed (see '+d.log+').';
    document.getElementById('sendResult').innerHTML='<div class="notice '+(d.failed>0?'amber':'green')+'"><span>'+(d.failed>0?'▲':'✓')+'</span><span>'+msg+audit+'</span></div>';
  }).catch(e=>{document.getElementById('testBtn').disabled=false;if(db) db.disabled=false;if(sb) sb.disabled=false;document.getElementById('sendStatus').textContent='';document.getElementById('sendResult').innerHTML='<div class="notice red"><span>✕</span><span>'+e+'</span></div>';});
}
</script></body></html>
"""

if __name__ == "__main__":
    print("=" * 60)
    print("  iExpense Compliance Automation Suite")
    print("  Open http://127.0.0.1:5000 in your browser")
    print("  (Test = to the process owner · Dummy = to you · Production = all cardholders)")
    print("=" * 60)
    app.run(debug=True, port=5000)
