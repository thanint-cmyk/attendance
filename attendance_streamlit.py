import os, re
from datetime import datetime, time

import streamlit as st
# Webcam via WebRTC (optional)
try:
    from streamlit_webrtc import webrtc_streamer, WebRtcMode
    HAS_WEBRTC = True
except Exception:
    HAS_WEBRTC = False


# --- พยายามโหลด OpenCV; ถ้าไม่มีจะปิดฟีเจอร์สแกนแต่แอปยังทำงานได้ ---
try:
    import cv2
    HAS_CV2 = True
except Exception:
    HAS_CV2 = False

import gspread
from google.oauth2.service_account import Credentials

# ------------------ Login ------------------
def login_gate():
    pwd_cfg = st.secrets.get("APP_PASSWORD", os.environ.get("APP_PASSWORD", "")).strip()
    if not pwd_cfg:
        return True
    if "authed" not in st.session_state:
        st.session_state.authed = False
    if st.session_state.authed:
        return True

    st.title("Login")
    pwd = st.text_input("Password", type="password")
    if st.button("Enter"):
        if pwd == pwd_cfg:
            st.session_state.authed = True
            st.rerun()
        else:
            st.error("Wrong password.")
    st.stop()

login_gate()

# ------------------ Utils ------------------
WEEKDAY_EN_TO_TH = {
    "Monday": "จันทร์", "Tuesday": "อังคาร", "Wednesday": "พุธ",
    "Thursday": "พฤหัสบดี", "Friday": "ศุกร์", "Saturday": "เสาร์", "Sunday": "อาทิตย์",
}

def get_session_th_and_cutoff(now_t: time):
    if now_t < time(12, 0, 0):
        return "เช้า", time(9, 10, 0)
    return "บ่าย", time(13, 10, 0)

def roster_sheet_name_for_now(now_dt: datetime) -> str:
    day_th = WEEKDAY_EN_TO_TH[now_dt.strftime("%A")]
    session_th, _ = get_session_th_and_cutoff(now_dt.time())
    return f"{day_th}{session_th}"

def normalize_seat(s: str) -> str:
    return (s or "").strip().upper()

def extract_student_id(raw_input: str):
    digits = re.findall(r"\d+", raw_input or "")
    if not digits:
        return None
    d = "".join(digits)
    if len(d) == 14:
        return d[3:13]
    if len(d) == 10:
        return d
    if 1 <= len(d) <= 3:
        return d
    return None

# ------------------ Google Sheets helpers ------------------
def get_gspread_client_from_secrets(readonly=False):
    sa_info = st.secrets["GOOGLE_SERVICE_ACCOUNT"]
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly" if readonly
        else "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
    return gspread.authorize(creds)

@st.cache_data(show_spinner=False)
def load_roster_id_and_seat_from_gsheet(sheet_id: str, worksheet_name: str):
    gc = get_gspread_client_from_secrets(readonly=True)
    sh = gc.open_by_key(sheet_id)
    try:
        ws = sh.worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        raise FileNotFoundError(f"Google Sheet has no worksheet named '{worksheet_name}'")

    values = ws.get_all_values()
    if not values:
        return {}, {}

    header = [h.strip() for h in values[0]]
    rows = values[1:]

    def idx_of(name):
        try: return header.index(name)
        except ValueError: return None

    idx_sid, idx_name, idx_seat = idx_of("Student ID"), idx_of("Full Name"), idx_of("Seat")
    if None in (idx_sid, idx_name, idx_seat):
        raise ValueError("Header must include: 'Student ID', 'Full Name', 'Seat'")

    roster_by_id, roster_by_seat = {}, {}
    for r in rows:
        if not r or idx_sid >= len(r): continue
        sid = (r[idx_sid] or "").strip()
        if not sid: continue
        full_name = (r[idx_name] or "").strip() if idx_name < len(r) else ""
        seat = (r[idx_seat] or "").strip() if idx_seat < len(r) else ""
        roster_by_id[sid] = (full_name, seat)
        if seat:
            roster_by_seat[seat.upper()] = (sid, full_name)
    return roster_by_id, roster_by_seat

def open_or_create_log_worksheet(sh: gspread.Spreadsheet, name: str) -> gspread.Worksheet:
    try:
        ws = sh.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=name, rows=2000, cols=10)
        ws.append_row([name], value_input_option="USER_ENTERED")
        ws.append_row(["Student ID", "Full Name", "Seat", "Check-in Time", "Status"], value_input_option="USER_ENTERED")
    return ws

# ------------------ Page & State ------------------
st.set_page_config(page_title="QR/Barcode Attendance", page_icon="✅")
st.title("QR/Barcode Attendance (Google Sheets)")

now = datetime.now()
day_en = now.strftime("%A")
if day_en in ["Saturday", "Sunday"]:
    st.error(f"Today is {day_en}. Attendance is closed.")
    st.stop()

session_th, cutoff_time = get_session_th_and_cutoff(now.time())
session_en = "Morning" if session_th == "เช้า" else "Afternoon"
roster_sheet = roster_sheet_name_for_now(now)

SHEET_ID = st.secrets.get("ROSTER_SHEET_ID", "").strip()
if not SHEET_ID:
    st.error("Missing ROSTER_SHEET_ID in Secrets.")
    st.stop()

# Load roster
try:
    ROSTER_BY_ID, ROSTER_BY_SEAT = load_roster_id_and_seat_from_gsheet(SHEET_ID, roster_sheet)
except Exception as e:
    st.error(f"Failed to load roster '{roster_sheet}': {e}")
    st.stop()

gc_write = get_gspread_client_from_secrets(readonly=False)
SH = gc_write.open_by_key(SHEET_ID)
today_str = now.strftime("%Y-%m-%d")
LOG_WS_NAME = f"Log {today_str} {session_en}"
LOG_WS = open_or_create_log_worksheet(SH, LOG_WS_NAME)

# preload duplicates
scanned_ids, used_seats = set(), set()
try:
    vals = LOG_WS.get_all_values()
    for r in vals[2:]:
        if not r: continue
        sid = (r[0].strip() if len(r) > 0 and r[0] else "")
        seat = (r[2].strip() if len(r) > 2 and r[2] else "")
        if sid: scanned_ids.add(sid)
        if seat: used_seats.add(normalize_seat(seat))
except Exception:
    pass

st.caption(f"Roster: **{roster_sheet}**  |  Write-back: **{LOG_WS_NAME}**")
st.link_button("Open Google Sheet", f"https://docs.google.com/spreadsheets/d/{SHEET_ID}")

# ------------------ Camera Scanner ------------------
st.subheader("Scan with phone camera (1D & QR)")
if not HAS_CV2 or not HAS_WEBRTC:
    st.warning(
        "Camera scanning disabled: missing cv2 or streamlit-webrtc.\n"
        "• ใช้งานได้ด้วยการกรอก Student ID / Seat ด้านล่างตามปกติ\n"
        "• แก้ให้ใช้กล้อง: เพิ่มแพ็กเกจ `opencv-python-headless` (หรือ `opencv-contrib-python-headless` "
        "ถ้าต้องการ 1D barcode) ใน `requirements.txt` แล้ว Deploy ใหม่"
    )
else:
    if "student_id_prefill" not in st.session_state:
        st.session_state.student_id_prefill = ""
    if "last_decode" not in st.session_state:
        st.session_state.last_decode = ""
    auto_submit = st.checkbox("Auto submit when a valid Student ID is detected", value=False)

    def decode_any_barcode(img_bgr) -> str:
        # 1) BarcodeDetector (1D+2D)
        try:
            bd = cv2.barcode_BarcodeDetector()
            ok, infos, types, corners = bd.detectAndDecode(img_bgr)
            if ok and infos:
                return (infos[0] or "").strip()
        except Exception:
            pass
        # 2) Fallback to QR only
        try:
            data, _, _ = cv2.QRCodeDetector().detectAndDecode(img_bgr)
            return (data or "").strip()
        except Exception:
            return ""

    def video_frame_callback(frame):
        img = frame.to_ndarray(format="bgr24")
        text = decode_any_barcode(img)
        if text and text != st.session_state.last_decode:
            st.session_state.last_decode = text
            sid_auto = extract_student_id(text)
            if sid_auto:
                st.session_state.student_id_prefill = sid_auto
                if auto_submit:
                    st.session_state["do_submit"] = True
        return frame

    webrtc_streamer(
        key="qr-cam",
        mode=WebRtcMode.SENDRECV,
        video_frame_callback=video_frame_callback,
        media_stream_constraints={"video": {"facingMode": "environment"}, "audio": False},
    )
    if st.session_state.last_decode:
        st.info(f"Scanned: {st.session_state.last_decode}")

# ------------------ Form ------------------
msg = st.empty()
with st.form("checkin_form", clear_on_submit=True):
    c1, c2 = st.columns([2, 1])
    with c1:
        sid_in = st.text_input(
            "Student ID",
            value=st.session_state.get("student_id_prefill", ""),
            placeholder="Scan or type 10 digits"
        )
        st.caption("Enter Student ID OR leave blank if using Seat.")
    with c2:
        seat_in = st.text_input("Seat", "", placeholder="e.g., A12").upper()
        st.caption("Enter Seat OR leave blank if using Student ID.")
    submitted = st.form_submit_button("Check-in ✅")

if st.session_state.get("do_submit"):
    submitted = True
    st.session_state["do_submit"] = False

def save_row(sid: str, full_name: str, seat: str):
    t = datetime.now()
    time_str = t.strftime("%H:%M:%S")
    status = "On time" if t.time() <= cutoff_time else "Late"
    try:
        LOG_WS.append_row([sid, full_name, seat, time_str, status], value_input_option="USER_ENTERED")
    except Exception as e:
        msg.error(f"Cannot write to Google Sheets log: {e}")
        return
    scanned_ids.add(sid)
    if seat: used_seats.add(seat)
    msg.success(f"✅ {sid} ({full_name}) • Seat: {seat or '-'} • {time_str} ({status})")

if submitted:
    # switch session/sheets if crossed noon/day
    now2 = datetime.now()
    new_session_th, new_cutoff = get_session_th_and_cutoff(now2.time())
    new_session_en = "Morning" if new_session_th == "เช้า" else "Afternoon"
    new_roster_sheet = roster_sheet_name_for_now(now2)

    if (new_roster_sheet != roster_sheet) or (new_session_en != session_en):
        try:
            ROSTER_BY_ID, ROSTER_BY_SEAT = load_roster_id_and_seat_from_gsheet(SHEET_ID, new_roster_sheet)
        except Exception as e:
            st.error(f"Failed to switch roster '{new_roster_sheet}': {e}")
            st.stop()
        try:
            new_name = f"Log {now2.strftime('%Y-%m-%d')} {new_session_en}"
            LOG_WS_NAME = new_name
            LOG_WS = open_or_create_log_worksheet(SH, LOG_WS_NAME)
            cutoff_time = new_cutoff
            session_en = new_session_en
            roster_sheet = new_roster_sheet
            scanned_ids.clear(); used_seats.clear()
            vals2 = LOG_WS.get_all_values()
            for r in vals2[2:]:
                if not r: continue
                sid = (r[0].strip() if len(r) > 0 and r[0] else "")
                seat = (r[2].strip() if len(r) > 2 and r[2] else "")
                if sid: scanned_ids.add(sid)
                if seat: used_seats.add(normalize_seat(seat))
            st.rerun()
        except Exception as e:
            st.error(f"Cannot open/create log sheet: {e}")
            st.stop()

    sid = extract_student_id(sid_in) if sid_in else None
    seat_norm = normalize_seat(seat_in) if seat_in else None

    if not sid and not seat_norm:
        msg.warning("Please enter either a Student ID or a Seat (or both).")
    else:
        if sid and not seat_norm:
            if sid not in ROSTER_BY_ID:
                msg.error(f"Student ID {sid} is not in roster sheet '{roster_sheet}'.")
            else:
                full_name, seat_from_table = ROSTER_BY_ID[sid]
                seat_from_table = normalize_seat(seat_from_table)
                if sid in scanned_ids:
                    msg.warning(f"{sid} has already checked in.")
                elif seat_from_table and seat_from_table in used_seats:
                    msg.error(f"Seat {seat_from_table} is already used.")
                else:
                    save_row(sid, full_name, seat_from_table)

        elif seat_norm and not sid:
            if seat_norm not in ROSTER_BY_SEAT:
                msg.error(f"Seat {seat_norm} is not in roster sheet '{roster_sheet}'.")
            else:
                sid_from_seat, full_name = ROSTER_BY_SEAT[seat_norm]
                if sid_from_seat in scanned_ids:
                    msg.warning(f"{sid_from_seat} has already checked in.")
                elif seat_norm in used_seats:
                    msg.error(f"Seat {seat_norm} is already used.")
                else:
                    save_row(sid_from_seat, full_name, seat_norm)

        elif sid and seat_norm:
            if sid not in ROSTER_BY_ID:
                msg.error(f"Student ID {sid} is not in roster sheet '{roster_sheet}'.")
            else:
                full_name, seat_from_table = ROSTER_BY_ID[sid]
                seat_from_table = normalize_seat(seat_from_table)
                if seat_from_table and (seat_norm != seat_from_table):
                    msg.error(f"{sid} is assigned to seat {seat_from_table} (not {seat_norm}).")
                elif sid in scanned_ids:
                    msg.warning(f"{sid} has already checked in.")
                elif seat_norm in used_seats:
                    msg.error(f"Seat {seat_norm} is already used.")
                else:
                    save_row(sid, full_name, seat_norm)
