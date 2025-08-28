# app.py — Streamlit + Google Sheets (Present + Absent in SAME sheet)
import streamlit as st
st.set_page_config(page_title="Attendance", page_icon="✅", layout="centered")

import pandas as pd
from datetime import datetime, time
import pytz, re
from io import BytesIO

# ===== Google Sheets =====
import gspread
from google.oauth2.service_account import Credentials

# ---------------- Time/Session helpers ----------------
TZ = pytz.timezone("Asia/Bangkok")

WEEKDAY_EN_TO_TH = {
    "Monday": "จันทร์", "Tuesday": "อังคาร", "Wednesday": "พุธ",
    "Thursday": "พฤหัสบดี", "Friday": "ศุกร์", "Saturday": "เสาร์", "Sunday": "อาทิตย์",
}
def get_session_th_and_cutoff(t: time):
    if t < time(12, 0, 0): return "เช้า", time(9, 10, 0)
    return "บ่าย", time(13, 10, 0)

def roster_sheet_name_for_now(now_dt: datetime) -> str:
    # Roster tabs ใช้ชื่อไทยเดิม (เช่น จันทร์เช้า/อังคารบ่าย)
    return f"{WEEKDAY_EN_TO_TH[now_dt.strftime('%A')]}{get_session_th_and_cutoff(now_dt.time())[0]}"

def normalize_seat(s: str) -> str:
    return (s or "").strip().upper()

def extract_student_id(raw: str):
    digits = re.findall(r"\\d+", raw or "")
    if not digits: return None
    d = "".join(digits)
    if len(d) == 14: return d[3:13]
    if len(d) == 10: return d
    if 1 <= len(d) <= 3: return d
    return None

# ---------------- Streamlit UI ----------------
st.title("Attendance")
st.header("Gen Chem 2302163 2302113 2302178")

now = datetime.now(TZ)
today_str = now.strftime("%Y-%m-%d")
day_en = now.strftime("%A")
if day_en in ["Saturday","Sunday"]:
    st.info(f"Today is {day_en} (attendance closed)")
    st.stop()

session_th, cutoff_time = get_session_th_and_cutoff(now.time())
session_en = "Morning" if session_th == "เช้า" else "Afternoon"

# ---- EN section key (no Thai at all for logs) ----
day_key = {
    "Monday":"mon","Tuesday":"tue","Wednesday":"wed","Thursday":"thu","Friday":"fri"
}[day_en]
session_key = {"Morning":"morning","Afternoon":"afternoon"}[session_en]
section_key = f"{day_key}_{session_key}"               # e.g., tue_afternoon
# Roster sheet (Thai tab names)
roster_ws_name = roster_sheet_name_for_now(now)

# ---------------- Secrets / Config ----------------
ROSTER_SHEET_KEY = st.secrets.get("SHEET_ROSTER_KEY", "")
LOG_KEYS_RAW = st.secrets.get("log_keys", {})
DEFAULT_LOG_KEY = st.secrets.get("SHEET_LOG_DEFAULT_KEY", "")  # optional fallback

if not ROSTER_SHEET_KEY:
    st.error("SHEET_ROSTER_KEY is not set in secrets.toml")
    st.stop()

from collections.abc import Mapping
if not (isinstance(LOG_KEYS_RAW, Mapping) and LOG_KEYS_RAW):
    st.error(
        "[log_keys] in secrets.toml must be a TOML table (dict). "
        f"(Found type={type(LOG_KEYS_RAW).__name__}; "
        f"available keys={list(st.secrets.keys())})"
    )
    st.stop()
LOG_KEYS = dict(LOG_KEYS_RAW)  # ensure plain dict

# choose log spreadsheet
if section_key in LOG_KEYS:
    CURRENT_LOG_SHEET_KEY = LOG_KEYS[section_key]
    DAILY_WS_NAME = today_str                            # daily tab in per-section file
else:
    if not DEFAULT_LOG_KEY:
        st.error(
            f"No mapping for section '{section_key}' in [log_keys] and SHEET_LOG_DEFAULT_KEY is not set."
        )
        st.stop()
    CURRENT_LOG_SHEET_KEY = DEFAULT_LOG_KEY
    DAILY_WS_NAME = f"{section_key} {today_str}"         # avoid mixing sections in default file

# ---------------- Google client ----------------
@st.cache_resource
def gs_client():
    sa_info = st.secrets.get("gcp_service_account", None)
    if not sa_info:
        st.error("No [gcp_service_account] in secrets.toml")
        st.stop()
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
    return gspread.authorize(creds)

def open_ws(spreadsheet_key: str, worksheet_name: str, create_if_missing=False):
    gc = gs_client()
    sh = gc.open_by_key(spreadsheet_key)
    try:
        return sh.worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        if create_if_missing:
            return sh.add_worksheet(title=worksheet_name, rows=2000, cols=20)
        raise

def ensure_headers(ws, headers):
    vals = ws.get_all_values()
    if not vals:
        ws.append_row(headers, value_input_option="USER_ENTERED")
    elif vals[0] != headers:
        ws.update("A1", [headers])

# ---------------- Load Roster ----------------
@st.cache_data(ttl=300)
def load_roster_dicts():
    ws = open_ws(ROSTER_SHEET_KEY, roster_ws_name)
    rows = ws.get_all_values()
    if not rows or len(rows) < 2:
        raise RuntimeError(f"Roster sheet '{roster_ws_name}' is empty or missing header")

    # Expect first 3 columns: A=Student ID, B=Full Name, C=Seat
    data = rows[1:]
    roster_by_id, roster_by_seat = {}, {}
    for r in data:
        sid = (r[0] if len(r)>0 else "").strip()
        name = (r[1] if len(r)>1 else "").strip()
        seat = normalize_seat(r[2] if len(r)>2 else "")
        if not sid: continue
        roster_by_id[sid] = (name, seat)
        if seat:
            roster_by_seat[seat] = (sid, name)
    return roster_by_id, roster_by_seat

try:
    ROSTER_BY_ID, ROSTER_BY_SEAT = load_roster_dicts()
except Exception as e:
    st.error(f"Failed to load roster '{roster_ws_name}': {e}")
    st.stop()

# ---------------- Daily worksheet (per day) ----------------
LOG_HEADERS = ["date","session","student_id","full_name","seat","time","status"]

@st.cache_data(ttl=30)
def read_today_log_df():
    ws = open_ws(CURRENT_LOG_SHEET_KEY, DAILY_WS_NAME, create_if_missing=True)
    vals = ws.get_all_values()
    if not vals:
        df = pd.DataFrame(columns=LOG_HEADERS)
    else:
        headers = vals[0]; rows = vals[1:]
        df = pd.DataFrame(rows, columns=headers)
        for c in LOG_HEADERS:
            if c not in df.columns: df[c] = ""
        df = df[LOG_HEADERS]
    return df

def append_today_row(row: dict):
    ws = open_ws(CURRENT_LOG_SHEET_KEY, DAILY_WS_NAME, create_if_missing=True)
    ensure_headers(ws, LOG_HEADERS)
    ws.append_row([row.get(h,"") for h in LOG_HEADERS], value_input_option="USER_ENTERED")
    read_today_log_df.clear()

# ---------------- Absent helpers (same-sheet style) ----------------
def compute_absent_df(roster_by_id: dict, today_df: pd.DataFrame) -> pd.DataFrame:
    """คืน DataFrame ของนักศึกษาที่ 'ยังไม่เช็กชื่อ' สำหรับวัน/ช่วงปัจจุบัน"""
    if today_df is None or today_df.empty:
        checked = set()
    else:
        mask = (today_df["date"] == today_str) & (today_df["session"] == session_en)
        checked = set(today_df.loc[mask, "student_id"].astype(str).tolist())

    rows = []
    for sid, (name, seat) in roster_by_id.items():
        if sid not in checked:
            rows.append({"student_id": sid, "full_name": name, "seat": seat, "status": "Absent"})
    df = pd.DataFrame(rows, columns=["student_id","full_name","seat","status"])
    return df.sort_values("student_id", ignore_index=True)

def write_absent_into_main_sheet_same_tab(absent_df: pd.DataFrame):
    """
    เขียนบล็อก Absent ไว้ท้ายชีต DAILY_WS_NAME เดียวกับ Present:
    - เก็บเฉพาะ log ของวันนี้ + ช่วงนี้ (Present = On time/Late หรืออะไรก็ได้ที่ไม่ใช่ Absent)
    - ลบ/แทนที่บล็อก Absent เดิม
    - ต่อท้ายด้วย "ABSENT LIST" และรายการ Absent
    """
    ws = open_ws(CURRENT_LOG_SHEET_KEY, DAILY_WS_NAME, create_if_missing=True)

    # อ่านข้อมูลทั้งหมดในชีต
    vals = ws.get_all_values()  # list[list[str]]
    headers = LOG_HEADERS[:]  # ["date","session","student_id","full_name","seat","time","status"]

    # กรอง "Present rows" เฉพาะของวันนี้/ช่วงนี้ และ "ไม่ใช่ Absent"
    present_rows = []
    if vals:
        start_idx = 1 if vals[0] == headers else 0
        for r in vals[start_idx:]:
            row = (r + [""]*len(headers))[:len(headers)]
            d, s, sid, name, seat, t, stat = row
            if d == today_str and s == session_en and str(stat).strip().lower() != "absent":
                present_rows.append([d,s,sid,name,seat,t,stat])

    # เคลียร์ทั้งชีต แล้วเขียนกลับ: headers + present + ABSENT LIST + absents
    ws.clear()
    ws.update("A1", [headers])
    if present_rows:
        ws.append_rows(present_rows, value_input_option="USER_ENTERED")

    # หัวคั่น Absent
    ws.append_row(["ABSENT LIST","","","","","",""], value_input_option="USER_ENTERED")

    if not absent_df.empty:
        rows = absent_df.rename(columns={
            "student_id":"Student ID","full_name":"Full Name","seat":"Seat","status":"Status"
        })[["Student ID","Full Name","Seat","Status"]].values.tolist()

        # map ให้เข้ากริดหลัก 7 คอลัมน์ของ LOG_HEADERS
        normalized = []
        for sid, name, seat, status in rows:
            normalized.append(["", "", sid, name, seat, "", status])
        ws.append_rows(normalized, value_input_option="USER_ENTERED")

def absent_to_excel_bytes(absent_df: pd.DataFrame) -> bytes:
    from openpyxl import Workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    wb = Workbook(); ws = wb.active
    ws.title = "Absent"
    ws.append([f"Absent list • {today_str} • {session_en}"])
    ws.append(["Student ID","Full Name","Seat","Status"])
    mini = absent_df.rename(columns={
        "student_id":"Student ID","full_name":"Full Name","seat":"Seat","status":"Status"
    })[["Student ID","Full Name","Seat","Status"]]
    for r in dataframe_to_rows(mini, index=False, header=False):
        if r and any(x is not None for x in r): ws.append(r)
    bio = BytesIO(); wb.save(bio); return bio.getvalue(), f"Absent {day_en} {session_en}.xlsx"

# ---------------- UI: Inputs (with on_change) ----------------
def handle_submit():
    sid_raw  = st.session_state.get("sid_input", "")
    seat_raw = st.session_state.get("seat_input", "")

    sid = extract_student_id(sid_raw) if sid_raw else None
    seat_in = normalize_seat(seat_raw) if seat_raw else None

    if not sid and not seat_in:
        st.warning("Please fill at least one field: Student ID or Seat")
        st.session_state.sid_input = ""
        st.session_state.seat_input = ""
        return

    log_df = read_today_log_df()

    def already_checked(_sid):
        if log_df.empty: return False
        m = (log_df["date"] == today_str) & (log_df["session"] == session_en) & (log_df["student_id"] == _sid)
        return bool(m.any())

    def seat_used(_seat):
        if log_df.empty: return False
        m = (log_df["date"] == today_str) & (log_df["session"] == session_en) & (log_df["seat"].str.upper() == _seat)
        return bool(m.any())

    # Case 1: ID only
    if sid and not seat_in:
        if sid not in ROSTER_BY_ID:
            st.warning(f"{sid} not in roster sheet '{roster_ws_name}'")
        else:
            full_name, seat_from = ROSTER_BY_ID[sid]
            seat_from = normalize_seat(seat_from)

            if already_checked(sid):
                st.warning(f"{sid} already checked in")
            elif seat_from and seat_used(seat_from):
                st.warning(f"Seat {seat_from} already used")
            else:
                status = "On time" if now.time() <= cutoff_time else "Late"
                row = dict(date=today_str, session=session_en, student_id=sid, full_name=full_name,
                           seat=seat_from, time=now.strftime("%H:%M:%S"), status=status)
                append_today_row(row)
                st.success(f"✅ {sid} ({full_name}) | Seat: {seat_from or '-'} | {row['time']} ({status})")

                # --- Auto-update ABSENT block at end of same sheet ---
                log_df = read_today_log_df()
                absent_df = compute_absent_df(ROSTER_BY_ID, log_df)
                write_absent_into_main_sheet_same_tab(absent_df)

    # Case 2: Seat only
    elif seat_in and not sid:
        if seat_in not in ROSTER_BY_SEAT:
            st.warning(f"Seat {seat_in} not in roster sheet '{roster_ws_name}'")
        else:
            sid2, full_name = ROSTER_BY_SEAT[seat_in]
            if already_checked(sid2):
                st.warning(f"{sid2} already checked in")
            elif seat_used(seat_in):
                st.warning(f"Seat {seat_in} already used")
            else:
                status = "On time" if now.time() <= cutoff_time else "Late"
                row = dict(date=today_str, session=session_en, student_id=sid2, full_name=full_name,
                           seat=seat_in, time=now.strftime("%H:%M:%S"), status=status)
                append_today_row(row)
                st.success(f"✅ {sid2} ({full_name}) | Seat: {seat_in} | {row['time']} ({status})")

                # --- Auto-update ABSENT block at end of same sheet ---
                log_df = read_today_log_df()
                absent_df = compute_absent_df(ROSTER_BY_ID, log_df)
                write_absent_into_main_sheet_same_tab(absent_df)

    # Case 3: both
    else:
        if sid not in ROSTER_BY_ID:
            st.warning(f"{sid} not in roster sheet '{roster_ws_name}'")
        else:
            full_name, seat_from = ROSTER_BY_ID[sid]
            seat_from = normalize_seat(seat_from)

            if seat_from and seat_in and seat_in != seat_from:
                st.warning(f"{sid} is assigned to seat {seat_from} (not {seat_in})")
            elif already_checked(sid):
                st.warning(f"{sid} already checked in")
            elif seat_in and seat_used(seat_in):
                st.warning(f"Seat {seat_in} already used")
            else:
                seat_final = seat_in or seat_from
                status = "On time" if now.time() <= cutoff_time else "Late"
                row = dict(date=today_str, session=session_en, student_id=sid, full_name=full_name,
                           seat=seat_final, time=now.strftime("%H:%M:%S"), status=status)
                append_today_row(row)
                st.success(f"✅ {sid} ({full_name}) | Seat: {seat_final or '-'} | {row['time']} ({status})")

                # --- Auto-update ABSENT block at end of same sheet ---
                log_df = read_today_log_df()
                absent_df = compute_absent_df(ROSTER_BY_ID, log_df)
                write_absent_into_main_sheet_same_tab(absent_df)

    # เคลียร์ช่องกรอกทุกครั้ง
    st.session_state.sid_input = ""
    st.session_state.seat_input = ""

# Inputs
col1, col2 = st.columns([2,1])
with col1:
    st.text_input("Student ID / Barcode", key="sid_input",
                  placeholder="e.g. 6530xxxxxx", on_change=handle_submit)
with col2:
    st.text_input("Seat", key="seat_input",
                  placeholder="e.g. A12", on_change=handle_submit)

if st.button("Check-in", type="primary"):
    handle_submit()

# ---------------- View + Download (today only) ----------------
st.divider()
st.subheader(f"Today's check-ins • {today_str} ({session_en}) • section '{section_key}'")

log_df = read_today_log_df()
today_df = log_df[(log_df["date"] == today_str) & (log_df["session"] == session_en)]

st.dataframe(
    today_df[["student_id","full_name","seat","time","status"]].rename(columns={
        "student_id":"Student ID","full_name":"Full Name","seat":"Seat","time":"Check-in Time","status":"Status"
    }),
    use_container_width=True, hide_index=True
)

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    from openpyxl import Workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    wb = Workbook(); ws = wb.active
    ws.title = today_str
    ws.append([f"Attendance date {today_str} | Session {session_en}"])
    ws.append(["Student ID", "Full Name", "Seat", "Check-in Time", "Status"])
    mini = df[["student_id","full_name","seat","time","status"]].rename(columns={
        "student_id":"Student ID","full_name":"Full Name","seat":"Seat","time":"Check-in Time","status":"Status"
    })
    for r in dataframe_to_rows(mini, index=False, header=False):
        if r and any(x is not None for x in r): ws.append(r)
    bio = BytesIO(); wb.save(bio); return bio.getvalue(), f"Attendance {day_en} {session_en}.xlsx"

excel_bytes, download_name = to_excel_bytes(today_df)
st.download_button("Download Excel (Present)", data=excel_bytes,
                   file_name=download_name,
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---------------- Absent list (UI preview) ----------------
st.divider()
st.subheader(f"Absent list • {today_str} ({session_en})")

absent_df = compute_absent_df(ROSTER_BY_ID, log_df)

st.dataframe(
    absent_df.rename(columns={
        "student_id":"Student ID","full_name":"Full Name","seat":"Seat","status":"Status"
    }),
    use_container_width=True, hide_index=True
)

# Download Absent as Excel
def absent_to_excel_bytes_ui(df: pd.DataFrame):
    b, name = absent_to_excel_bytes(df)
    return b, name

absent_xlsx_bytes, absent_name = absent_to_excel_bytes_ui(absent_df)
st.download_button("Download Excel (Absent)", data=absent_xlsx_bytes,
                   file_name=absent_name,
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Manual update button (in case of edits)
if st.button("Update Absent block (same sheet)"):
    write_absent_into_main_sheet_same_tab(absent_df)
    st.success("Updated Absent block at end of the same sheet.")