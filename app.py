# ====================================================
# 📋 ระบบติดตามการลาและไปราชการ สคร.9
# ✨ v3.0 — Full Feature Upgrade
# ====================================================
#
# 🆕 ฟีเจอร์ใหม่ใน v3.0:
#
# 📊 REPORTS & ANALYTICS
#   [R1] วันลาคงเหลือ: คำนวณสิทธิ์ลาตาม พ.ร.บ. แสดง progress bar
#   [R2] แนวโน้มการลารายเดือน: line chart เปรียบเทียบย้อนหลัง 12 เดือน
#   [R3] อัตราการขาดงาน / มาสาย แยกรายกลุ่มงาน
#   [R4] Heatmap ปฏิทินกลางหน่วยงาน
#   [R5] Export รายงานสรุปผู้บริหาร (Excel หลายชีต)
#
# 👤 STAFF MANAGEMENT
#   [S1] ฐานข้อมูลบุคลากร (staff_master.xlsx) — เพิ่ม/แก้ไข/ปิดใช้งาน
#   [S2] Dropdown ชื่อดึงจาก Master แทน aggregate จาก data
#   [S3] แสดงสถานะบุคลากร (ปฏิบัติงาน / ลาออก / ยืมตัว)
#   [S4] ตรวจสอบ quota ก่อนบันทึกลา + แจ้งเตือนเมื่อใกล้หมดสิทธิ์
#
# 🔔 NOTIFICATIONS
#   [N1] LINE Notify เมื่อบันทึกการลา (ส่งไปกลุ่ม LINE)
#   [N2] LINE Notify เมื่อบันทึกไปราชการ
#   [N3] In-app Activity Feed (10 รายการล่าสุด)
#   [N4] Alert แดง เมื่อบุคลากรลาเกินสิทธิ์
#
# 📱 UX / UI
#   [U1] Custom CSS: card layout, mobile-responsive, badge, color scheme
#   [U2] Sticky sidebar navigation
#   [U3] สถานะ indicator (🟢🟡🔴) ในทุกตาราง
#   [U4] Collapsible sections ใน form ยาวๆ
#   [U5] Toast notification (st.toast) แทน st.success ธรรมดา
#
# ====================================================

import io
import time
import logging
import datetime as dt
import requests
from typing import Dict, List, Optional, Tuple, Tuple

import numpy as np
import pandas as pd
import altair as alt
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# ===========================
# 🔧 Logging
# ===========================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler()],
)
logger = logging.getLogger(__name__)

# ===========================
# [U1] 📱 Custom CSS — Mobile Responsive + Modern UI
# ===========================
CUSTOM_CSS = """
<style>
/* ── Global ── */
html, body, [class*="css"] { font-family: 'Sarabun', sans-serif; }

/* ── Sidebar ── */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0f172a 0%, #1e293b 100%);
    color: white;
}
section[data-testid="stSidebar"] * { color: white !important; }
section[data-testid="stSidebar"] .stRadio > label { 
    background: rgba(255,255,255,0.05);
    border-radius: 8px;
    padding: 6px 12px;
    margin: 2px 0;
    display: block;
    transition: background 0.2s;
}
section[data-testid="stSidebar"] .stRadio > label:hover {
    background: rgba(255,255,255,0.15);
}

/* ── Cards (Metric boxes) ── */
div[data-testid="metric-container"] {
    background: white;
    border-radius: 12px;
    padding: 16px;
    border: 1px solid #e2e8f0;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08);
}

/* ── Status badges ── */
.badge-green  { background:#dcfce7; color:#166534; padding:2px 10px; border-radius:999px; font-size:0.78rem; font-weight:600; }
.badge-yellow { background:#fef9c3; color:#854d0e; padding:2px 10px; border-radius:999px; font-size:0.78rem; font-weight:600; }
.badge-red    { background:#fee2e2; color:#991b1b; padding:2px 10px; border-radius:999px; font-size:0.78rem; font-weight:600; }
.badge-blue   { background:#dbeafe; color:#1e40af; padding:2px 10px; border-radius:999px; font-size:0.78rem; font-weight:600; }
.badge-gray   { background:#f1f5f9; color:#475569; padding:2px 10px; border-radius:999px; font-size:0.78rem; font-weight:600; }

/* ── Section header ── */
.section-header {
    background: linear-gradient(90deg, #0ea5e9, #6366f1);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    font-size: 1.4rem;
    font-weight: 700;
    margin-bottom: 1rem;
}

/* ── Activity feed ── */
.activity-item {
    padding: 10px 14px;
    border-left: 3px solid #6366f1;
    background: #f8fafc;
    border-radius: 0 8px 8px 0;
    margin-bottom: 8px;
    font-size: 0.87rem;
}

/* ── Progress bar (quota) ── */
.quota-bar-wrap { background:#e2e8f0; border-radius:999px; height:10px; margin:4px 0; }
.quota-bar-fill { height:10px; border-radius:999px; transition: width 0.4s; }

/* ── Mobile: collapse sidebar ── */
@media (max-width: 768px) {
    div[data-testid="metric-container"] { margin-bottom: 8px; }
    .block-container { padding: 1rem !important; }
}
</style>
<link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap" rel="stylesheet">
"""

# ===========================
# 🔐 App Init
# ===========================
st.set_page_config(
    page_title="สคร.9 — HR Tracking v3",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

if "gcp_service_account" not in st.secrets:
    st.error("❌ ไม่พบ gcp_service_account ใน secrets.toml")
    st.stop()

# ===========================
# [R1] Leave Quota Config
# ===========================
LEAVE_QUOTA: Dict[str, int] = {
    "ลาป่วย":                       90,
    "ลากิจส่วนตัว":                  45,
    "ลาพักผ่อน":                     10,
    "ลาคลอดบุตร":                    98,
    "ลาอุปสมบท":                    120,
    "ลาช่วยเหลือภริยาที่คลอดบุตร":   15,
}

STAFF_GROUPS: List[str] = [
    "กลุ่มบริหารทั่วไป",
    "กลุ่มบริหารทั่วไป (งานธุรการ)",
    "กลุ่มบริหารทั่วไป (งานการเงินและบัญชี)",
    "กลุ่มบริหารทั่วไป (งานการเจ้าหน้าที่)",
    "กลุ่มบริหารทั่วไป (งานพัสดุและยานพาหนะ (งานพัสดุ))",
    "กลุ่มบริหารทั่วไป (งานพัสดุและยานพาหนะ (งานยานพาหนะ))",
    "กลุ่มบริหารทั่วไป (งานพัสดุและยานพาหนะ (งานอาคารสถานที่))",
    "กลุ่มยุทธศาสตร์และแผนงาน",
    "กลุ่มระบาดวิทยาและตอบโต้ภาวะฉุกเฉินทางสาธารณสุข",
    "กลุ่มโรคติดต่อ", "กลุ่มโรคไม่ติดต่อ", "กลุ่มโรคติดต่อเรื้อรัง",
    "กลุ่มโรคติดต่อนำโดยแมลง",
    "กลุ่มโรคติดต่อนำโดยแมลง (ศตม. 9.1 จ.ชัยภูมิ)",
    "กลุ่มโรคติดต่อนำโดยแมลง (ศตม. 9.2 จ.บุรีรัมย์)",
    "กลุ่มโรคติดต่อนำโดยแมลง (ศตม. 9.3 จ.สุรินทร์)",
    "กลุ่มโรคติดต่อนำโดยแมลง (ศตม. 9.4 อ.ปากช่อง)",
    "กลุ่มโรคจากการประกอบอาชีพและสิ่งแวดล้อม",
    "กลุ่มห้องปฏิบัติการทางการแพทย์ด้านควบคุมโรค",
    "กลุ่มสื่อสารความเสี่ยงโรคและภัยสุขภาพ",
    "กลุ่มพัฒนานวัตกรรมและวิจัย", "กลุ่มพัฒนาองค์กร",
    "ศูนย์ฝึกอบรมนักระบาดวิทยาภาคสนาม", "ศูนย์บริการเวชศาสตร์ป้องกัน",
    "งานกฎหมาย", "งานเภสัชกรรม", "ด่านควบคุมโรคติดต่อระหว่างประเทศ", "อื่นๆ",
]

LEAVE_TYPES: List[str] = list(LEAVE_QUOTA.keys())

COLUMN_MAPPING: Dict[str, str] = {
    "ชื่อพนักงาน": "ชื่อ-สกุล",
    "ชื่อ": "ชื่อ-สกุล",
    "fullname": "ชื่อ-สกุล",
}

FILE_ATTEND   = "attendance_report.xlsx"
FILE_LEAVE    = "leave_report.xlsx"
FILE_TRAVEL   = "travel_report.xlsx"
FILE_STAFF    = "staff_master.xlsx"        # [S1] NEW
FILE_NOTIFY   = "activity_log.xlsx"        # [N3] NEW
FILE_HOLIDAYS = "special_holidays.xlsx"    # [H1] NEW — วันหยุดพิเศษ
FOLDER_ID   = "1YFJZvs59ahRHmlRrKcQwepWJz6A-4B7d"
ATTACHMENT_FOLDER_NAME = "Attachments_Leave_App"

# ===========================
# Google Drive Service
# ===========================
@st.cache_resource
def init_drive_service():
    for attempt in range(3):
        try:
            creds = service_account.Credentials.from_service_account_info(
                st.secrets["gcp_service_account"],
                scopes=["https://www.googleapis.com/auth/drive"],
            )
            svc = build("drive", "v3", credentials=creds)
            logger.info("Drive connected")
            return svc
        except Exception as e:
            if attempt == 2:
                st.error("❌ เชื่อมต่อ Google Drive ไม่สำเร็จ")
                st.stop()
            time.sleep(2 ** attempt)

service = init_drive_service()

# ===========================
# Drive Helpers
# ===========================
def get_file_id(filename: str, parent_id: str = FOLDER_ID) -> Optional[str]:
    try:
        res = service.files().list(
            q=f"name='{filename}' and '{parent_id}' in parents and trashed=false",
            fields="files(id)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        files = res.get("files", [])
        return files[0]["id"] if files else None
    except Exception as e:
        logger.error(f"get_file_id({filename}): {e}")
        return None

def get_or_create_folder(folder_name: str, parent_id: str) -> Optional[str]:
    try:
        res = service.files().list(
            q=f"name='{folder_name}' and '{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
            fields="files(id)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        folders = res.get("files", [])
        if folders:
            return folders[0]["id"]
        meta = {"name": folder_name, "parents": [parent_id], "mimeType": "application/vnd.google-apps.folder"}
        new = service.files().create(body=meta, supportsAllDrives=True, fields="id").execute()
        return new.get("id")
    except Exception as e:
        logger.error(f"get_or_create_folder({folder_name}): {e}")
        return None

@st.cache_data(ttl=300)
def read_excel_from_drive(filename: str) -> pd.DataFrame:
    for attempt in range(3):
        try:
            fid = get_file_id(filename)
            if not fid:
                return pd.DataFrame()
            req = service.files().get_media(fileId=fid, supportsAllDrives=True)
            fh = io.BytesIO()
            dl = MediaIoBaseDownload(fh, req)
            done = False
            while not done:
                _, done = dl.next_chunk()
            fh.seek(0)
            return pd.read_excel(fh, engine="openpyxl")
        except Exception as e:
            if attempt == 2:
                st.error(f"อ่านไฟล์ {filename} ไม่สำเร็จ")
                return pd.DataFrame()
            time.sleep(2 ** attempt)
    return pd.DataFrame()

def write_excel_to_drive(filename: str, df: pd.DataFrame) -> bool:
    try:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            df.to_excel(w, index=False)
        buf.seek(0)
        fid = get_file_id(filename)
        media = MediaIoBaseUpload(buf, mimetype=EXCEL_MIME)
        if fid:
            service.files().update(fileId=fid, media_body=media, supportsAllDrives=True).execute()
        else:
            service.files().create(
                body={"name": filename, "parents": [FOLDER_ID]},
                media_body=media,
                supportsAllDrives=True,
            ).execute()
        st.cache_data.clear()
        return True
    except Exception as e:
        logger.error(f"write_excel_to_drive({filename}): {e}")
        st.error(f"บันทึกไฟล์ล้มเหลว: {e}")
        return False

def backup_excel(filename: str, df: pd.DataFrame) -> None:
    if df.empty:
        return
    try:
        fid = get_file_id(filename)
        if fid:
            ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
            service.files().copy(
                fileId=fid,
                body={"name": f"BAK_{ts}_{filename}", "parents": [FOLDER_ID]},
                supportsAllDrives=True,
            ).execute()
    except Exception as e:
        logger.warning(f"backup_excel({filename}): {e}")

def upload_pdf_to_drive(uploaded_file, new_filename: str, folder_id: str) -> str:
    try:
        meta = {"name": new_filename, "parents": [folder_id]}
        media = MediaIoBaseUpload(io.BytesIO(uploaded_file.getvalue()), mimetype="application/pdf", resumable=True)
        created = service.files().create(body=meta, media_body=media, supportsAllDrives=True, fields="id, webViewLink").execute()
        service.permissions().create(
            fileId=created["id"],
            body={"type": "anyone", "role": "reader"},
            supportsAllDrives=True,
        ).execute()
        return created.get("webViewLink", "-")
    except Exception as e:
        logger.error(f"upload_pdf: {e}")
        return "-"

# ===========================
# Data Processing
# ===========================
def normalize_date_col(df: pd.DataFrame, col: str) -> pd.DataFrame:
    if not df.empty and col in df.columns:
        df[col] = pd.to_datetime(df[col], errors="coerce").dt.normalize()
    return df

def clean_names(df: pd.DataFrame, col: str) -> pd.DataFrame:
    if not df.empty and col in df.columns:
        df[col] = df[col].astype(str).str.strip()
    return df

def preprocess_dataframes(df_leave, df_travel, df_att):
    # Rename columns
    for df in [df_att]:
        for old, new in COLUMN_MAPPING.items():
            if old in df.columns:
                df.rename(columns={old: new}, inplace=True)
    # Dates
    for col in ["วันที่เริ่ม", "วันที่สิ้นสุด"]:
        df_leave  = normalize_date_col(df_leave,  col)
        df_travel = normalize_date_col(df_travel, col)
    df_att = normalize_date_col(df_att, "วันที่")
    # Names
    for df in [df_leave, df_travel, df_att]:
        df = clean_names(df, "ชื่อ-สกุล")
    return df_leave, df_travel, df_att

def count_weekdays(start_date, end_date, extra_holidays: Optional[List[dt.date]] = None) -> int:
    """นับวันทำการ (จ-ศ) หักวันหยุดพิเศษด้วย (ถ้ามี)"""
    if not start_date or not end_date:
        return 0
    if isinstance(start_date, dt.datetime):
        start_date = start_date.date()
    if isinstance(end_date, dt.datetime):
        end_date = end_date.date()
    base = int(np.busday_count(start_date, end_date + dt.timedelta(days=1)))
    if extra_holidays:
        overlap = sum(
            1 for h in extra_holidays
            if start_date <= h <= end_date and h.weekday() < 5
        )
        base = max(0, base - overlap)
    return base

# ===========================
# [H1] Special Holiday Helpers
# ===========================

FIXED_THAI_HOLIDAYS: List[Tuple[int, int, str]] = [
    (1,  1,  "วันขึ้นปีใหม่"),
    (4,  6,  "วันจักรี"),
    (4,  13, "วันสงกรานต์"),
    (4,  14, "วันสงกรานต์"),
    (4,  15, "วันสงกรานต์"),
    (5,  1,  "วันแรงงานแห่งชาติ"),
    (5,  5,  "วันฉัตรมงคล"),
    (6,  3,  "วันเฉลิมพระชนมพรรษา สมเด็จพระราชินี"),
    (7,  28, "วันเฉลิมพระชนมพรรษา ร.10"),
    (8,  12, "วันแม่แห่งชาติ"),
    (10, 13, "วันคล้ายวันสวรรคต ร.9"),
    (10, 23, "วันปิยมหาราช"),
    (12, 5,  "วันพ่อแห่งชาติ / วันชาติ"),
    (12, 10, "วันรัฐธรรมนูญ"),
    (12, 31, "วันสิ้นปี"),
]

HOLIDAY_TYPE_OPTIONS = ["วันหยุดราชการ", "วันหยุดพิเศษ (ผนวก)", "วันหยุดประจำหน่วยงาน", "อื่นๆ"]
HOLIDAY_COLS = ["วันที่", "ชื่อวันหยุด", "ประเภท", "หมายเหตุ"]

def _can_make_date(year: int, month: int, day: int) -> bool:
    """ตรวจว่า date(year, month, day) valid หรือไม่"""
    try:
        dt.date(year, month, day)
        return True
    except ValueError:
        return False

def get_fixed_holidays_for_year(year: int) -> pd.DataFrame:
    """สร้าง DataFrame วันหยุดราชการตายตัวสำหรับปีที่กำหนด"""
    rows = []
    for month, day, name in FIXED_THAI_HOLIDAYS:
        try:
            d = dt.date(year, month, day)
            rows.append({
                "วันที่":      pd.Timestamp(d),
                "ชื่อวันหยุด": name,
                "ประเภท":     "วันหยุดราชการ",
                "หมายเหตุ":   "กำหนดโดยระบบ (แก้ไขไม่ได้)",
            })
        except ValueError:
            pass
    return pd.DataFrame(rows)

@st.cache_data(ttl=300)
def load_holidays_raw() -> pd.DataFrame:
    """โหลดเฉพาะวันหยุดที่ Admin กำหนดเองจาก Drive"""
    df = read_excel_from_drive(FILE_HOLIDAYS)
    if not df.empty:
        df["วันที่"] = pd.to_datetime(df["วันที่"], errors="coerce")
        df = df.dropna(subset=["วันที่"])
        for col in HOLIDAY_COLS:
            if col not in df.columns:
                df[col] = ""
    return df

def load_holidays_all(year: Optional[int] = None) -> pd.DataFrame:
    """รวมวันหยุดจาก Drive + วันหยุดราชการตายตัว"""
    df_custom = load_holidays_raw()
    frames: List[pd.DataFrame] = []
    if year:
        frames.append(get_fixed_holidays_for_year(year))
    if not df_custom.empty:
        if year:
            df_y = df_custom[df_custom["วันที่"].dt.year == year]
        else:
            df_y = df_custom
        frames.append(df_y)
    if not frames:
        return pd.DataFrame(columns=HOLIDAY_COLS)
    df_all = pd.concat(frames, ignore_index=True)
    df_all = df_all.drop_duplicates(subset=["วันที่"]).sort_values("วันที่").reset_index(drop=True)
    return df_all

def get_holiday_dates(year: Optional[int] = None) -> List[dt.date]:
    """คืน list ของ dt.date วันหยุดทั้งหมด (ราชการ + พิเศษ)"""
    df_h = load_holidays_all(year)
    if df_h.empty:
        return []
    return pd.to_datetime(df_h["วันที่"], errors="coerce").dropna().dt.date.tolist()

def get_holiday_name(d: dt.date, holiday_df: pd.DataFrame) -> str:
    """หาชื่อวันหยุดจาก DataFrame — คืน '' ถ้าไม่ใช่วันหยุด"""
    if holiday_df.empty:
        return ""
    match = holiday_df[pd.to_datetime(holiday_df["วันที่"], errors="coerce").dt.date == d]
    if not match.empty:
        return str(match.iloc[0].get("ชื่อวันหยุด", "วันหยุดพิเศษ"))
    return ""


    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None
    if isinstance(val, dt.time):
        return val
    try:
        return pd.to_datetime(str(val)).time()
    except Exception:
        return None

# ===========================
# [S2] Staff Master Helpers
# ===========================
STAFF_MASTER_COLS = ["ชื่อ-สกุล", "กลุ่มงาน", "ตำแหน่ง", "ประเภทบุคลากร", "วันเริ่มงาน", "สถานะ"]

def get_active_staff(df_staff: pd.DataFrame) -> List[str]:
    """ดึงรายชื่อบุคลากรที่ยังปฏิบัติงานอยู่จาก master"""
    if df_staff.empty or "ชื่อ-สกุล" not in df_staff.columns:
        return []
    if "สถานะ" in df_staff.columns:
        df_active = df_staff[df_staff["สถานะ"] == "ปฏิบัติงาน"]
    else:
        df_active = df_staff
    return sorted(df_active["ชื่อ-สกุล"].dropna().astype(str).str.strip().unique().tolist())

def get_all_names_fallback(df_leave, df_travel, df_att) -> List[str]:
    """Fallback: รวมชื่อจากข้อมูลทั้งหมด (กรณียังไม่มี master)"""
    all_names: set = set()
    for df in [df_leave, df_travel, df_att]:
        if not df.empty and "ชื่อ-สกุล" in df.columns:
            all_names.update(df["ชื่อ-สกุล"].dropna().astype(str).str.strip().unique())
    return sorted([n for n in all_names if n.lower() != "nan"])

# ===========================
# [R1] Leave Quota Functions
# ===========================
def get_leave_used(name: str, leave_type: str, df_leave: pd.DataFrame, year: int) -> int:
    if df_leave.empty or "ชื่อ-สกุล" not in df_leave.columns:
        return 0
    mask = (
        (df_leave["ชื่อ-สกุล"] == name)
        & (df_leave["ประเภทการลา"] == leave_type)
        & (df_leave["วันที่เริ่ม"].dt.year == year)
    )
    return int(df_leave.loc[mask, "จำนวนวันลา"].sum())

def get_quota_status(used: int, quota: int) -> Tuple[str, str]:
    """คืน (emoji_indicator, badge_class)"""
    pct = used / quota if quota > 0 else 1
    if pct >= 1.0:
        return "🔴", "badge-red"
    elif pct >= 0.8:
        return "🟡", "badge-yellow"
    else:
        return "🟢", "badge-green"

def quota_bar_html(used: int, quota: int) -> str:
    pct = min(used / quota, 1.0) if quota > 0 else 1.0
    color = "#22c55e" if pct < 0.8 else ("#f59e0b" if pct < 1.0 else "#ef4444")
    return (
        f'<div class="quota-bar-wrap">'
        f'<div class="quota-bar-fill" style="width:{pct*100:.0f}%;background:{color};"></div>'
        f'</div>'
    )

def check_leave_quota(name: str, leave_type: str, days_req: int, df_leave: pd.DataFrame, year: int) -> Optional[str]:
    quota = LEAVE_QUOTA.get(leave_type, 9999)
    used  = get_leave_used(name, leave_type, df_leave, year)
    remaining = quota - used
    if days_req > remaining:
        return f"❌ ลาเกินสิทธิ์! คงเหลือ {remaining} วัน (ขอ {days_req} วัน, ใช้ไปแล้ว {used}/{quota} วัน)"
    if (used + days_req) / quota >= 0.8:
        return f"⚠️ เตือน: หากลาครั้งนี้ จะใช้สิทธิ์ลา{leave_type}ไปแล้ว {used+days_req}/{quota} วัน (ใกล้หมดสิทธิ์)"
    return None

# ===========================
# [N1][N2] LINE Notify
# ===========================
def send_line_notify(message: str) -> bool:
    token = st.secrets.get("line_notify_token", "")
    if not token:
        logger.info("LINE Notify: token not configured, skipping")
        return False
    try:
        resp = requests.post(
            "https://notify-api.line.me/api/notify",
            headers={"Authorization": f"Bearer {token}"},
            data={"message": message},
            timeout=5,
        )
        return resp.status_code == 200
    except Exception as e:
        logger.warning(f"LINE Notify failed: {e}")
        return False

def format_leave_notify(record: dict) -> str:
    return (
        f"\n🔔 แจ้งการลา — สคร.9"
        f"\n👤 {record.get('ชื่อ-สกุล','')}  ({record.get('กลุ่มงาน','')})"
        f"\n📌 {record.get('ประเภทการลา','')}  {record.get('จำนวนวันลา','')} วัน"
        f"\n📅 {record.get('วันที่เริ่ม','').strftime('%d/%m/%Y') if hasattr(record.get('วันที่เริ่ม'),'strftime') else record.get('วันที่เริ่ม','')} "
        f"ถึง {record.get('วันที่สิ้นสุด','').strftime('%d/%m/%Y') if hasattr(record.get('วันที่สิ้นสุด'),'strftime') else record.get('วันที่สิ้นสุด','')}"
        f"\n📝 {record.get('เหตุผล','')}"
        f"\n⏰ {record.get('Timestamp','')}"
    )

def format_travel_notify(persons: List[str], project: str, location: str, d_start, d_end, days: int) -> str:
    names_str = ", ".join(persons[:5])
    if len(persons) > 5:
        names_str += f" และอีก {len(persons)-5} คน"
    return (
        f"\n✈️ แจ้งไปราชการ — สคร.9"
        f"\n👥 {names_str}"
        f"\n📌 {project}"
        f"\n📍 {location}"
        f"\n📅 {d_start.strftime('%d/%m/%Y')} ถึง {d_end.strftime('%d/%m/%Y')} ({days} วันทำการ)"
        f"\n⏰ {dt.datetime.now().strftime('%d/%m/%Y %H:%M')}"
    )

# ===========================
# [N3] Activity Log
# ===========================
ACTIVITY_LOG_COLS = ["Timestamp", "ประเภท", "รายละเอียด", "ผู้เกี่ยวข้อง"]

def log_activity(action_type: str, detail: str, persons: str) -> None:
    try:
        df_log = read_excel_from_drive(FILE_NOTIFY)
        if df_log.empty:
            df_log = pd.DataFrame(columns=ACTIVITY_LOG_COLS)
        new_row = {
            "Timestamp": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "ประเภท": action_type,
            "รายละเอียด": detail,
            "ผู้เกี่ยวข้อง": persons,
        }
        df_log = pd.concat([df_log, pd.DataFrame([new_row])], ignore_index=True)
        # เก็บแค่ 500 รายการล่าสุด
        df_log = df_log.tail(500).reset_index(drop=True)
        write_excel_to_drive(FILE_NOTIFY, df_log)
    except Exception as e:
        logger.warning(f"log_activity failed: {e}")

# ===========================
# Validation
# ===========================
def validate_leave_data(name, start_date, end_date, reason, df_leave) -> List[str]:
    errors: List[str] = []
    if not name or not name.strip():
        errors.append("❌ กรุณาเลือกชื่อ-สกุล")
    if start_date > end_date:
        errors.append("❌ วันที่เริ่มต้องน้อยกว่าหรือเท่ากับวันที่สิ้นสุด")
    if not reason or len(reason.strip()) < 5:
        errors.append("❌ กรุณาระบุเหตุผลอย่างน้อย 5 ตัวอักษร")
    if not df_leave.empty and name:
        s, e = pd.to_datetime(start_date), pd.to_datetime(end_date)
        overlap = df_leave[
            (df_leave["ชื่อ-สกุล"] == name)
            & (df_leave["วันที่เริ่ม"] <= e)
            & (df_leave["วันที่สิ้นสุด"] >= s)
        ]
        if not overlap.empty:
            errors.append("❌ มีการลาซ้ำในช่วงเวลานี้แล้ว")
    return errors

def validate_travel_data(staff_list, project, location, start_date, end_date) -> List[str]:
    errors: List[str] = []
    if not staff_list:
        errors.append("❌ กรุณาเลือกผู้เดินทางอย่างน้อย 1 คน")
    if not project or len(project.strip()) < 3:
        errors.append("❌ กรุณาระบุชื่อโครงการ/กิจกรรม")
    if not location or len(location.strip()) < 3:
        errors.append("❌ กรุณาระบุสถานที่")
    if start_date > end_date:
        errors.append("❌ วันที่เริ่มต้องน้อยกว่าหรือเท่ากับวันที่สิ้นสุด")
    return errors

# ===========================
# Security
# ===========================
def check_admin_password(password: str) -> bool:
    correct = st.secrets.get("admin_password", "204486")
    return password == correct

# ===========================
# 🚀 Main App
# ===========================
# Sidebar
with st.sidebar:
    st.markdown("## 🏥 สคร.9 HR System")
    st.markdown("---")
    menu = st.radio(
        "เมนูใช้งาน",
        [
            "🏠 หน้าหลัก",
            "📊 Dashboard & รายงาน",
            "📅 ตรวจสอบการปฏิบัติงาน",
            "📅 ปฏิทินกลาง",
            "🧭 บันทึกไปราชการ",
            "🕒 บันทึกการลา",
            "📈 วันลาคงเหลือ",
            "👤 จัดการบุคลากร",
            "🔔 กิจกรรมล่าสุด",
            "⚙️ ผู้ดูแลระบบ",
        ],
        label_visibility="collapsed",
    )
    st.markdown("---")
    st.caption(f"v3.0 | {dt.date.today().strftime('%d/%m/%Y')}")

# ============================================================
# 🏠 หน้าหลัก
# ============================================================
if menu == "🏠 หน้าหลัก":
    st.markdown('<div class="section-header">🏥 ระบบติดตามการลา ไปราชการ และการปฏิบัติงาน<br>สำนักงานป้องกันควบคุมโรคที่ 9</div>', unsafe_allow_html=True)

    with st.spinner("กำลังโหลดภาพรวม..."):
        df_leave  = read_excel_from_drive(FILE_LEAVE)
        df_travel = read_excel_from_drive(FILE_TRAVEL)
        df_att    = read_excel_from_drive(FILE_ATTEND)
        df_leave, df_travel, df_att = preprocess_dataframes(df_leave, df_travel, df_att)

    # Quick KPIs
    c1, c2, c3, c4 = st.columns(4)
    this_month = dt.date.today().strftime("%Y-%m")

    leave_this_month = 0
    travel_this_month = 0
    if not df_leave.empty and "วันที่เริ่ม" in df_leave.columns:
        leave_this_month = len(df_leave[df_leave["วันที่เริ่ม"].dt.strftime("%Y-%m") == this_month])
    if not df_travel.empty and "วันที่เริ่ม" in df_travel.columns:
        travel_this_month = len(df_travel[df_travel["วันที่เริ่ม"].dt.strftime("%Y-%m") == this_month])

    c1.metric("📋 ลาเดือนนี้", f"{leave_this_month} ครั้ง")
    c2.metric("🚗 ราชการเดือนนี้", f"{travel_this_month} ครั้ง")
    c3.metric("📋 ลารวมทั้งหมด", f"{len(df_leave)} ครั้ง")
    c4.metric("🚗 ราชการรวมทั้งหมด", f"{len(df_travel)} ครั้ง")

    st.markdown("---")
    col_news, col_feat = st.columns([2, 1])

    with col_news:
        st.subheader("🆕 อัปเดต v3.0")
        st.markdown("""
        | ฟีเจอร์ | สถานะ |
        |--------|------|
        | 📊 วิเคราะห์แนวโน้มการลารายเดือน | ✅ ใหม่ |
        | 🗓️ ปฏิทินกลางหน่วยงาน (Heatmap) | ✅ ใหม่ |
        | 📈 วันลาคงเหลือ / สิทธิ์ลา | ✅ ใหม่ |
        | 👤 จัดการฐานข้อมูลบุคลากร | ✅ ใหม่ |
        | 🔔 LINE Notify แจ้งเตือนอัตโนมัติ | ✅ ใหม่ |
        | 📝 Activity Feed กิจกรรมล่าสุด | ✅ ใหม่ |
        | 📱 UI ใหม่ รองรับ Mobile | ✅ ใหม่ |
        """)

    with col_feat:
        st.subheader("⚙️ สถานะการเชื่อมต่อ")
        line_token = st.secrets.get("line_notify_token", "")
        st.markdown(f"LINE Notify: {'🟢 เชื่อมต่อแล้ว' if line_token else '🔴 ยังไม่ตั้งค่า'}")
        st.markdown(f"Google Drive: 🟢 เชื่อมต่อแล้ว")
        st.markdown(f"Staff Master: {'🟢 มีข้อมูล' if not read_excel_from_drive(FILE_STAFF).empty else '🟡 ยังไม่มีข้อมูล'}")

# ============================================================
# 📊 Dashboard & รายงาน
# ============================================================
elif menu == "📊 Dashboard & รายงาน":
    st.markdown('<div class="section-header">📊 Dashboard & วิเคราะห์ข้อมูล</div>', unsafe_allow_html=True)

    with st.spinner("กำลังโหลดข้อมูล..."):
        df_leave  = read_excel_from_drive(FILE_LEAVE)
        df_travel = read_excel_from_drive(FILE_TRAVEL)
        df_att    = read_excel_from_drive(FILE_ATTEND)
        df_leave, df_travel, df_att = preprocess_dataframes(df_leave, df_travel, df_att)

    # KPIs
    c1, c2, c3 = st.columns(3)
    c1.metric("📋 จำนวนการลา (ทั้งหมด)", len(df_leave))
    c2.metric("🚗 จำนวนไปราชการ (ทั้งหมด)", len(df_travel))
    c3.metric("👆 ข้อมูลสแกนนิ้ว", len(df_att))

    st.divider()
    tab_a, tab_b, tab_c, tab_late, tab_d = st.tabs([
        "🏆 Top Charts", "📈 แนวโน้มรายเดือน",
        "🏢 แยกกลุ่มงาน", "⏰ สถิติการมาสาย", "📥 Export รายงาน",
    ])

    with tab_a:
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("วันลารวมแยกตามกลุ่มงาน (Top 10)")
            if not df_leave.empty and "กลุ่มงาน" in df_leave.columns:
                df_c = df_leave.groupby("กลุ่มงาน")["จำนวนวันลา"].sum().nlargest(10).reset_index()
                chart = alt.Chart(df_c).mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4).encode(
                    x=alt.X("จำนวนวันลา:Q", title="วันลารวม"),
                    y=alt.Y("กลุ่มงาน:N", sort="-x", title=""),
                    color=alt.value("#6366f1"),
                    tooltip=["กลุ่มงาน", "จำนวนวันลา"],
                ).properties(height=320)
                st.altair_chart(chart, use_container_width=True)
            else:
                st.info("ไม่มีข้อมูล")

        with col2:
            st.subheader("ประเภทการลา (สัดส่วน)")
            if not df_leave.empty and "ประเภทการลา" in df_leave.columns:
                df_pie = df_leave["ประเภทการลา"].value_counts().reset_index()
                df_pie.columns = ["ประเภท", "จำนวนครั้ง"]
                pie = alt.Chart(df_pie).mark_arc(innerRadius=50).encode(
                    theta=alt.Theta("จำนวนครั้ง:Q"),
                    color=alt.Color("ประเภท:N", legend=alt.Legend(orient="bottom")),
                    tooltip=["ประเภท", "จำนวนครั้ง"],
                ).properties(height=280)
                st.altair_chart(pie, use_container_width=True)
            else:
                st.info("ไม่มีข้อมูล")

    with tab_b:
        # [R2] แนวโน้มรายเดือน
        st.subheader("📈 แนวโน้มการลาและไปราชการ (ย้อนหลัง 12 เดือน)")

        if not df_leave.empty and "วันที่เริ่ม" in df_leave.columns:
            df_leave["เดือน"] = df_leave["วันที่เริ่ม"].dt.to_period("M").astype(str)
            df_travel["เดือน"] = df_travel["วันที่เริ่ม"].dt.to_period("M").astype(str) if not df_travel.empty and "วันที่เริ่ม" in df_travel.columns else ""

            df_trend_l = df_leave.groupby("เดือน")["จำนวนวันลา"].sum().reset_index()
            df_trend_l.columns = ["เดือน", "จำนวน"]
            df_trend_l["ประเภท"] = "วันลา"

            df_trend_t = pd.DataFrame()
            if not df_travel.empty and "เดือน" in df_travel.columns:
                df_trend_t = df_travel.groupby("เดือน")["จำนวนวัน"].sum().reset_index()
                df_trend_t.columns = ["เดือน", "จำนวน"]
                df_trend_t["ประเภท"] = "วันราชการ"

            df_trend = pd.concat([df_trend_l, df_trend_t], ignore_index=True).tail(24)

            trend_chart = alt.Chart(df_trend).mark_line(point=True).encode(
                x=alt.X("เดือน:O", title="เดือน"),
                y=alt.Y("จำนวน:Q", title="จำนวนวัน"),
                color=alt.Color("ประเภท:N"),
                tooltip=["เดือน", "ประเภท", "จำนวน"],
            ).properties(height=320)
            st.altair_chart(trend_chart, use_container_width=True)
        else:
            st.info("ยังไม่มีข้อมูลเพียงพอ")

    with tab_c:
        # [R3] แยกกลุ่มงาน
        st.subheader("🏢 สถิติแยกรายกลุ่มงาน")
        if not df_leave.empty and "กลุ่มงาน" in df_leave.columns and "ประเภทการลา" in df_leave.columns:
            df_grp = df_leave.groupby(["กลุ่มงาน", "ประเภทการลา"])["จำนวนวันลา"].sum().reset_index()
            grouped_chart = alt.Chart(df_grp).mark_bar().encode(
                x=alt.X("จำนวนวันลา:Q"),
                y=alt.Y("กลุ่มงาน:N", sort="-x"),
                color=alt.Color("ประเภทการลา:N"),
                tooltip=["กลุ่มงาน", "ประเภทการลา", "จำนวนวันลา"],
            ).properties(height=420)
            st.altair_chart(grouped_chart, use_container_width=True)
        else:
            st.info("ไม่มีข้อมูล")

    # ─────────────────────────────────────────────
    # ⏰ Tab: สถิติการมาสาย
    # ─────────────────────────────────────────────
    with tab_late:
        st.subheader("⏰ สถิติการมาสาย / ออกก่อน")

        if df_att.empty:
            st.info("ยังไม่มีข้อมูลสแกนนิ้วในระบบ")
        else:
            # ── filter: ปี / กลุ่มงาน
            late_col1, late_col2, late_col3 = st.columns([1, 1, 2])
            att_years = sorted(
                df_att["วันที่"].dropna()
                .pipe(lambda s: pd.to_datetime(s, errors="coerce"))
                .dt.year.dropna().astype(int).unique().tolist(),
                reverse=True,
            )
            with late_col1:
                sel_year = st.selectbox(
                    "ปี", att_years,
                    key="late_year",
                )
            with late_col2:
                sel_late_group = st.selectbox(
                    "กลุ่มงาน", ["ทุกกลุ่ม"] + STAFF_GROUPS,
                    key="late_group",
                )
            with late_col3:
                sel_late_status = st.multiselect(
                    "สถานะที่ต้องการดู",
                    ["มาสาย", "ออกก่อน", "มาสายและออกก่อน", "ขาดงาน"],
                    default=["มาสาย", "มาสายและออกก่อน"],
                    key="late_status",
                )

            WORK_START_L = dt.time(8, 30)
            WORK_END_L   = dt.time(16, 30)

            df_att_y = df_att.copy()
            df_att_y["วันที่"] = pd.to_datetime(df_att_y["วันที่"], errors="coerce")
            df_att_y = df_att_y[df_att_y["วันที่"].dt.year == sel_year]

            if df_att_y.empty:
                st.info(f"ไม่มีข้อมูลสแกนในปี {sel_year}")
            else:
                # กรองกลุ่มงาน (ถ้ามี staff master)
                df_staff_dash = read_excel_from_drive(FILE_STAFF)
                names_in_group = None
                if sel_late_group != "ทุกกลุ่ม" and not df_staff_dash.empty and "กลุ่มงาน" in df_staff_dash.columns:
                    names_in_group = set(
                        df_staff_dash[df_staff_dash["กลุ่มงาน"] == sel_late_group]["ชื่อ-สกุล"]
                        .astype(str).str.strip().tolist()
                    )
                    df_att_y = df_att_y[df_att_y["ชื่อ-สกุล"].isin(names_in_group)]

                # คำนวณสถานะรายแถว
                def calc_late_status(row) -> str:
                    if pd.to_datetime(row["วันที่"], errors="coerce").weekday() >= 5:
                        return "วันหยุด"
                    t_in  = parse_time(row.get("เวลาเข้า", ""))
                    t_out = parse_time(row.get("เวลาออก", ""))
                    if not t_in and not t_out:
                        return "ขาดงาน"
                    late = t_in and t_in > WORK_START_L
                    early = not t_out or t_out < WORK_END_L
                    if late and early:  return "มาสายและออกก่อน"
                    if late:            return "มาสาย"
                    if early:           return "ออกก่อน"
                    return "มาปกติ"

                df_att_y["สถานะสแกน"] = df_att_y.apply(calc_late_status, axis=1)
                df_att_y["เดือน"] = df_att_y["วันที่"].dt.strftime("%Y-%m")

                # ── KPI row
                total_days   = len(df_att_y[df_att_y["สถานะสแกน"] != "วันหยุด"])
                total_late   = len(df_att_y[df_att_y["สถานะสแกน"].isin(["มาสาย","มาสายและออกก่อน"])])
                total_early  = len(df_att_y[df_att_y["สถานะสแกน"].isin(["ออกก่อน","มาสายและออกก่อน"])])
                total_absent = len(df_att_y[df_att_y["สถานะสแกน"] == "ขาดงาน"])
                late_rate    = f"{total_late/total_days*100:.1f}%" if total_days > 0 else "—"

                kc1, kc2, kc3, kc4 = st.columns(4)
                kc1.metric("⏰ มาสาย",         f"{total_late} ครั้ง",  f"อัตรา {late_rate}")
                kc2.metric("🚪 ออกก่อนเวลา",   f"{total_early} ครั้ง")
                kc3.metric("❌ ขาดงาน",         f"{total_absent} ครั้ง")
                kc4.metric("📅 วันทำการ (รวม)", f"{total_days} วัน")

                st.divider()

                # ── กราฟแนวโน้มรายเดือน: มาสาย/ขาดงาน/ออกก่อน
                df_late_trend = (
                    df_att_y[df_att_y["สถานะสแกน"].isin(
                        ["มาสาย","ออกก่อน","มาสายและออกก่อน","ขาดงาน"]
                    )]
                    .groupby(["เดือน","สถานะสแกน"])
                    .size()
                    .reset_index(name="จำนวนครั้ง")
                )

                if not df_late_trend.empty:
                    late_chart = (
                        alt.Chart(df_late_trend)
                        .mark_bar()
                        .encode(
                            x=alt.X("เดือน:O", title="เดือน", sort=None),
                            y=alt.Y("จำนวนครั้ง:Q", title="จำนวนครั้ง", stack="zero"),
                            color=alt.Color(
                                "สถานะสแกน:N",
                                scale=alt.Scale(
                                    domain=["มาสาย","ออกก่อน","มาสายและออกก่อน","ขาดงาน"],
                                    range=["#F59E0B","#EF4444","#F97316","#991B1B"],
                                ),
                                legend=alt.Legend(orient="bottom"),
                            ),
                            tooltip=["เดือน","สถานะสแกน","จำนวนครั้ง"],
                        )
                        .properties(height=280, title=f"แนวโน้มการมาสาย/ขาดงาน ปี {sel_year}")
                    )
                    st.altair_chart(late_chart, use_container_width=True)
                else:
                    st.success("🎉 ไม่พบข้อมูลการมาสายหรือขาดงานในปีนี้")

                st.divider()

                # ── Top ผู้มาสายบ่อย (จำแนกตามสถานะที่เลือก)
                if sel_late_status:
                    df_top_late = (
                        df_att_y[df_att_y["สถานะสแกน"].isin(sel_late_status)]
                        .groupby("ชื่อ-สกุล")
                        .size()
                        .reset_index(name="จำนวนครั้ง")
                        .sort_values("จำนวนครั้ง", ascending=False)
                        .head(15)
                    )
                    if not df_top_late.empty:
                        st.subheader(f"🏅 Top 15 ผู้มีสถานะ {' / '.join(sel_late_status)} บ่อยที่สุด")
                        top_chart = (
                            alt.Chart(df_top_late)
                            .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
                            .encode(
                                x=alt.X("จำนวนครั้ง:Q", title="จำนวนครั้ง"),
                                y=alt.Y("ชื่อ-สกุล:N", sort="-x", title=""),
                                color=alt.value("#F59E0B"),
                                tooltip=["ชื่อ-สกุล","จำนวนครั้ง"],
                            )
                            .properties(height=max(200, len(df_top_late) * 24))
                        )
                        st.altair_chart(top_chart, use_container_width=True)

                        # ตาราง export
                        st.subheader("📋 รายละเอียดรายคน")
                        df_person_detail = (
                            df_att_y[df_att_y["สถานะสแกน"].isin(sel_late_status)]
                            .groupby(["ชื่อ-สกุล","สถานะสแกน"])
                            .size()
                            .reset_index(name="จำนวนครั้ง")
                            .pivot_table(
                                index="ชื่อ-สกุล",
                                columns="สถานะสแกน",
                                values="จำนวนครั้ง",
                                fill_value=0,
                            )
                            .reset_index()
                        )
                        st.dataframe(df_person_detail, use_container_width=True)

                        buf_late = io.BytesIO()
                        with pd.ExcelWriter(buf_late, engine="xlsxwriter") as w:
                            df_person_detail.to_excel(w, index=False, sheet_name="สรุปรายคน")
                            df_att_y[df_att_y["สถานะสแกน"].isin(sel_late_status)].to_excel(
                                w, index=False, sheet_name="รายการทั้งหมด"
                            )
                        buf_late.seek(0)
                        st.download_button(
                            "📥 Export รายงานการมาสาย",
                            buf_late,
                            f"LateReport_{sel_year}_{sel_late_group}.xlsx",
                            mime=EXCEL_MIME,
                        )
                    else:
                        st.success("🎉 ไม่พบข้อมูลสถานะที่เลือกในปีนี้")

    with tab_d:
        # [R5] Export Excel หลายชีต
        st.subheader("📥 Export รายงานสรุปผู้บริหาร")

        today = dt.date.today()
        month_opts = pd.date_range(
            start=f"{today.year - 2}-01-01",
            end=f"{today.year + 1}-12-31",
            freq="MS"
        ).strftime("%Y-%m").tolist()
        cur_ym = today.strftime("%Y-%m")
        default_i = month_opts.index(cur_ym) if cur_ym in month_opts else 0
        export_month = st.selectbox("เลือกเดือน", month_opts, index=default_i)

        if st.button("📊 สร้างรายงาน Excel", use_container_width=True, type="primary"):
            with st.spinner("กำลังสร้างรายงาน..."):
                m_start = pd.to_datetime(export_month + "-01")
                m_end   = m_start + pd.offsets.MonthEnd(0)

                df_lm = df_leave[(df_leave["วันที่เริ่ม"] >= m_start) & (df_leave["วันที่เริ่ม"] <= m_end)] if not df_leave.empty else pd.DataFrame()
                df_tm = df_travel[(df_travel["วันที่เริ่ม"] >= m_start) & (df_travel["วันที่เริ่ม"] <= m_end)] if not df_travel.empty else pd.DataFrame()

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    wb = writer.book

                    # Sheet สรุป
                    summary_data = {
                        "รายการ": ["จำนวนการลา (ครั้ง)", "จำนวนวันลารวม", "จำนวนไปราชการ (ครั้ง)", "จำนวนวันราชการรวม"],
                        "จำนวน": [
                            len(df_lm),
                            int(df_lm["จำนวนวันลา"].sum()) if not df_lm.empty and "จำนวนวันลา" in df_lm.columns else 0,
                            len(df_tm),
                            int(df_tm["จำนวนวัน"].sum()) if not df_tm.empty and "จำนวนวัน" in df_tm.columns else 0,
                        ],
                    }
                    pd.DataFrame(summary_data).to_excel(writer, sheet_name="📋 สรุป", index=False)
                    if not df_lm.empty:
                        df_lm.to_excel(writer, sheet_name="การลา", index=False)
                    if not df_tm.empty:
                        df_tm.to_excel(writer, sheet_name="ไปราชการ", index=False)
                    # Sheet วันลารายประเภท
                    if not df_lm.empty and "ประเภทการลา" in df_lm.columns:
                        df_lm.groupby("ประเภทการลา")["จำนวนวันลา"].sum().reset_index().to_excel(writer, sheet_name="ลาแยกประเภท", index=False)

                st.download_button(
                    "⬇️ ดาวน์โหลดรายงาน",
                    output.getvalue(),
                    f"HR_Report_{export_month}.xlsx",
                    mime=EXCEL_MIME,
                    use_container_width=True,
                )

# ============================================================
# 📅 ตรวจสอบการปฏิบัติงาน
# ============================================================
elif menu == "📅 ตรวจสอบการปฏิบัติงาน":
    st.markdown('<div class="section-header">📅 สรุปการมาปฏิบัติงานรายวัน</div>', unsafe_allow_html=True)

    with st.spinner("กำลังโหลดข้อมูล..."):
        df_att    = read_excel_from_drive(FILE_ATTEND)
        df_leave  = read_excel_from_drive(FILE_LEAVE)
        df_travel = read_excel_from_drive(FILE_TRAVEL)
        df_staff  = read_excel_from_drive(FILE_STAFF)
        df_leave, df_travel, df_att = preprocess_dataframes(df_leave, df_travel, df_att)

        all_names = get_active_staff(df_staff) or get_all_names_fallback(df_leave, df_travel, df_att)

    if df_att.empty:
        months = [dt.datetime.now().strftime("%Y-%m")]
    else:
        df_att["เดือน"] = df_att["วันที่"].dt.strftime("%Y-%m")
        months = sorted(df_att["เดือน"].dropna().unique().tolist()) or [dt.datetime.now().strftime("%Y-%m")]

    name_col = next((c for c in ["ชื่อ-สกุล","ชื่อพนักงาน","ชื่อ"] if not df_att.empty and c in df_att.columns), "ชื่อ-สกุล")

    # ── ตัวกรอง ────────────────────────────────────────────
    st.markdown("#### 🔍 ตัวกรองข้อมูล")
    fc1, fc2 = st.columns([2, 2])

    with fc1:
        # Preset shortcuts
        preset = st.radio(
            "ช่วงเวลาด่วน",
            ["เลือกเอง", "เดือนปัจจุบัน", "3 เดือนล่าสุด", "ครึ่งปีแรก", "ครึ่งปีหลัง", "ทั้งปีนี้"],
            horizontal=True,
            key="att_preset",
        )

    # สร้าง default ตาม preset
    today_m = dt.date.today().strftime("%Y-%m")
    def last_n_months(n):
        result = []
        d = dt.date.today().replace(day=1)
        for _ in range(n):
            result.append(d.strftime("%Y-%m"))
            d = (d - dt.timedelta(days=1)).replace(day=1)
        return sorted(result)

    cur_year = dt.date.today().year
    preset_map = {
        "เดือนปัจจุบัน":  [today_m],
        "3 เดือนล่าสุด":  last_n_months(3),
        "ครึ่งปีแรก":     [f"{cur_year}-{m:02d}" for m in range(1, 7)],
        "ครึ่งปีหลัง":    [f"{cur_year}-{m:02d}" for m in range(7, 13)],
        "ทั้งปีนี้":       [f"{cur_year}-{m:02d}" for m in range(1, 13)],
    }
    default_months = preset_map.get(preset, [today_m]) if preset != "เลือกเอง" else [today_m]
    # กรองเฉพาะเดือนที่มีในข้อมูล
    valid_defaults = [m for m in default_months if m in months] or [months[-1]]

    with fc2:
        if preset == "เลือกเอง":
            selected_months = st.multiselect(
                "📅 เลือกเดือน (เลือกได้หลายเดือน)",
                months,
                default=valid_defaults,
                key="att_months_multi",
            )
        else:
            selected_months = valid_defaults
            st.info(f"📅 เดือนที่เลือก: **{', '.join(selected_months)}**")

    if not selected_months:
        st.warning("กรุณาเลือกอย่างน้อย 1 เดือน")
        st.stop()

    fc3, fc4 = st.columns([2, 2])
    with fc3:
        selected_names = st.multiselect("👥 บุคลากร (ว่าง = ทุกคน)", all_names, key="att_names")
    with fc4:
        sel_att_group = st.selectbox("🏢 กลุ่มงาน", ["ทุกกลุ่ม"] + STAFF_GROUPS, key="att_group")

    # ── กรองชื่อตามกลุ่มงาน ────────────────────────────────
    df_staff_att = read_excel_from_drive(FILE_STAFF)
    names_to_process = selected_names or all_names
    if sel_att_group != "ทุกกลุ่ม" and not df_staff_att.empty and "กลุ่มงาน" in df_staff_att.columns:
        grp_set = set(
            df_staff_att[df_staff_att["กลุ่มงาน"] == sel_att_group]["ชื่อ-สกุล"]
            .astype(str).str.strip().tolist()
        )
        names_to_process = [n for n in names_to_process if n in grp_set]

    if not names_to_process:
        st.warning("ไม่มีข้อมูลบุคลากร")
        st.stop()

    # ── สร้าง date_range จากทุกเดือนที่เลือก ────────────────
    all_dates = pd.DatetimeIndex([])
    for ym in selected_months:
        ms = pd.to_datetime(ym + "-01")
        me = ms + pd.offsets.MonthEnd(0)
        all_dates = all_dates.append(pd.date_range(ms, me, freq="D"))
    all_dates = all_dates.sort_values()

    # ── โหลดข้อมูลสแกนเฉพาะเดือนที่เลือก ───────────────────
    if not df_att.empty and "เดือน" in df_att.columns:
        df_months_att = df_att[df_att["เดือน"].isin(selected_months)].copy()
    else:
        df_months_att = pd.DataFrame()
    if not df_months_att.empty:
        df_months_att[name_col] = df_months_att[name_col].astype(str).str.strip()

    WORK_START = dt.time(8, 30)
    WORK_END   = dt.time(16, 30)

    # ── โหลดวันหยุดพิเศษ (เดือนที่เลือก) ───────────────────
    sel_years = list({int(ym[:4]) for ym in selected_months})
    holiday_dates_set: set = set()
    holiday_df_lookup = pd.DataFrame()
    for yr in sel_years:
        holiday_dates_set.update(get_holiday_dates(yr))
        hdf = load_holidays_all(yr)
        holiday_df_lookup = pd.concat([holiday_df_lookup, hdf], ignore_index=True)

    # ── คำนวณข้อมูลรายวัน ────────────────────────────────────
    records = []
    prog = st.progress(0, text="กำลังประมวลผล...")
    for i, name in enumerate(names_to_process):
        prog.progress((i + 1) / len(names_to_process), text=f"กำลังประมวลผล {name}...")
        for d in all_dates:
            d_date = d.date()
            rec = {"ชื่อพนักงาน": name, "วันที่": d_date, "เดือน": d.strftime("%Y-%m"),
                   "เวลาเข้า": "", "เวลาออก": "", "สถานะ": ""}
            att = df_months_att[(df_months_att[name_col] == name) & (df_months_att["วันที่"] == d)] if not df_months_att.empty else pd.DataFrame()

            # ตรวจวันหยุดพิเศษก่อน (มีความสำคัญสูงกว่าสถานะอื่น ยกเว้นการลา)
            is_special_hday = d_date in holiday_dates_set
            special_hday_name = get_holiday_name(d_date, holiday_df_lookup) if is_special_hday else ""

            in_leave, leave_type = False, ""
            ul = df_leave[df_leave["ชื่อ-สกุล"] == name] if not df_leave.empty else pd.DataFrame()
            if not ul.empty:
                ml = ul[(ul["วันที่เริ่ม"] <= d) & (ul["วันที่สิ้นสุด"] >= d)]
                if not ml.empty:
                    in_leave, leave_type = True, ml.iloc[0].get("ประเภทการลา", "")

            in_travel = False
            ut = df_travel[df_travel["ชื่อ-สกุล"] == name] if not df_travel.empty else pd.DataFrame()
            if not ut.empty:
                mt = ut[(ut["วันที่เริ่ม"] <= d) & (ut["วันที่สิ้นสุด"] >= d)]
                in_travel = not mt.empty

            if in_leave:
                rec["สถานะ"] = f"ลา ({leave_type})"
            elif in_travel:
                rec["สถานะ"] = "ไปราชการ"
            elif d.weekday() >= 5:
                rec["สถานะ"] = "วันหยุด"
            elif is_special_hday:
                # วันหยุดพิเศษ — แสดงชื่อ
                rec["สถานะ"] = f"วันหยุด ({special_hday_name})"
            elif not att.empty:
                row = att.iloc[0]
                rec["เวลาเข้า"] = row.get("เวลาเข้า", "")
                rec["เวลาออก"] = row.get("เวลาออก", "")
                t_in  = parse_time(rec["เวลาเข้า"])
                t_out = parse_time(rec["เวลาออก"])
                if not t_in and not t_out:
                    rec["สถานะ"] = "ขาดงาน"
                elif t_in and t_in > WORK_START:
                    rec["สถานะ"] = "มาสายและออกก่อน" if (not t_out or t_out < WORK_END) else "มาสาย"
                elif not t_out or t_out < WORK_END:
                    rec["สถานะ"] = "ออกก่อน"
                else:
                    rec["สถานะ"] = "มาปกติ"
            else:
                rec["สถานะ"] = "วันหยุด" if d.weekday() >= 5 else "ขาดงาน"
            records.append(rec)
    prog.empty()

    df_daily = pd.DataFrame(records).sort_values(["ชื่อพนักงาน","วันที่"])

    def simplify_status(s):
        return "ลา" if isinstance(s, str) and s.startswith("ลา") else s
    df_daily["สถานะย่อ"] = df_daily["สถานะ"].apply(simplify_status)

    STATUS_COLORS = {
        "มาปกติ": "background-color:#d4edda",
        "มาสาย": "background-color:#ffeeba",
        "ออกก่อน": "background-color:#f8d7da",
        "มาสายและออกก่อน": "background-color:#fcd5b5",
        "ลา": "background-color:#d1ecf1",
        "ไปราชการ": "background-color:#fff3cd",
        "วันหยุด": "background-color:#e2e3e5",
        "ขาดงาน": "background-color:#f5c6cb",
    }
    def color_status(val):
        for k, v in STATUS_COLORS.items():
            if k in str(val):
                return v
        return ""

    # ── แสดงผลแบบ Tab ─────────────────────────────────────────
    att_tab1, att_tab2, att_tab3 = st.tabs([
        "📋 ตารางรายวัน",
        "📊 สรุปรายเดือน (ภาพรวม)",
        "📈 กราฟแนวโน้ม",
    ])

    with att_tab1:
        month_filter_daily = st.selectbox(
            "กรองดูเฉพาะเดือน (ว่าง = ทุกเดือนที่เลือก)",
            ["ทั้งหมด"] + selected_months,
            key="daily_month_filter",
        )
        df_show_daily = df_daily if month_filter_daily == "ทั้งหมด" else df_daily[df_daily["เดือน"] == month_filter_daily]
        st.caption(f"แสดง {len(df_show_daily)} แถว | {len(selected_months)} เดือน | {len(names_to_process)} คน")
        st.dataframe(
            df_show_daily.drop(columns=["สถานะย่อ"], errors="ignore")
            .style.map(color_status, subset=["สถานะ"]),
            use_container_width=True,
            height=480,
        )

    with att_tab2:
        # ── สรุปภาพรวมหลายเดือน ───────────────────────────────
        required_cols = ["มาปกติ","มาสาย","ออกก่อน","มาสายและออกก่อน","ลา","ไปราชการ","วันหยุด","ขาดงาน"]

        # pivot รายคน-รายเดือน (เลือกได้)
        view_mode = st.radio(
            "มุมมอง",
            ["รวมทุกเดือน (ภาพรวม)", "แยกรายเดือน"],
            horizontal=True,
            key="att_view_mode",
        )

        if view_mode == "รวมทุกเดือน (ภาพรวม)":
            summary = df_daily.pivot_table(
                index="ชื่อพนักงาน", columns="สถานะย่อ", aggfunc="size", fill_value=0
            )
            for col in required_cols:
                if col not in summary.columns:
                    summary[col] = 0
            summary = summary[[c for c in required_cols if c in summary.columns]].reset_index()

            # เพิ่มคอลัมน์ % อัตราการมาปกติ
            work_days = summary.get("มาปกติ", 0)
            all_work  = summary[[c for c in ["มาปกติ","มาสาย","ออกก่อน","มาสายและออกก่อน","ขาดงาน"] if c in summary.columns]].sum(axis=1)
            summary["% มาปกติ"] = (work_days / all_work.replace(0, 1) * 100).round(1).astype(str) + "%"

            st.caption(f"📅 ช่วงที่เลือก: {selected_months[0]} ถึง {selected_months[-1]} ({len(selected_months)} เดือน)")
            st.dataframe(summary, use_container_width=True, height=420)

        else:  # แยกรายเดือน
            summary_monthly = df_daily.pivot_table(
                index=["ชื่อพนักงาน","เดือน"], columns="สถานะย่อ", aggfunc="size", fill_value=0
            )
            for col in required_cols:
                if col not in summary_monthly.columns:
                    summary_monthly[col] = 0
            summary_monthly = summary_monthly[[c for c in required_cols if c in summary_monthly.columns]].reset_index()
            st.caption(f"แสดงข้อมูลแยกทุกเดือน: {len(selected_months)} เดือน | {len(names_to_process)} คน")
            st.dataframe(summary_monthly, use_container_width=True, height=480)

        # ── ปุ่ม Export รวม ─────────────────────────────────────
        st.divider()
        out = io.BytesIO()
        period_label = f"{selected_months[0]}_to_{selected_months[-1]}"
        with pd.ExcelWriter(out, engine="xlsxwriter") as w:
            df_daily.drop(columns=["สถานะย่อ"], errors="ignore").to_excel(
                w, index=False, sheet_name="รายวัน"
            )
            if view_mode == "รวมทุกเดือน (ภาพรวม)":
                summary.to_excel(w, index=False, sheet_name="สรุปภาพรวม")
            else:
                summary_monthly.to_excel(w, index=False, sheet_name="สรุปแยกเดือน")
        out.seek(0)
        st.download_button(
            "📥 ดาวน์โหลด Excel",
            out,
            f"Attendance_{period_label}.xlsx",
            mime=EXCEL_MIME,
            use_container_width=True,
        )

    with att_tab3:
        # ── กราฟแนวโน้มสถานะรายเดือน ─────────────────────────
        st.subheader("📈 แนวโน้มสถานะรายเดือน")

        trend_status_sel = st.multiselect(
            "เลือกสถานะที่ต้องการดูในกราฟ",
            required_cols,
            default=["มาปกติ","มาสาย","ขาดงาน","ลา","ไปราชการ"],
            key="trend_status_sel",
        )

        if trend_status_sel:
            df_trend_att = (
                df_daily[df_daily["สถานะย่อ"].isin(trend_status_sel)]
                .groupby(["เดือน","สถานะย่อ"])
                .size()
                .reset_index(name="จำนวนวัน")
            )
            if not df_trend_att.empty:
                trend_att_chart = (
                    alt.Chart(df_trend_att)
                    .mark_line(point=alt.OverlayMarkDef(size=60))
                    .encode(
                        x=alt.X("เดือน:O", title="เดือน", sort=None),
                        y=alt.Y("จำนวนวัน:Q", title="จำนวนวัน-คน"),
                        color=alt.Color("สถานะย่อ:N", legend=alt.Legend(orient="bottom")),
                        tooltip=["เดือน","สถานะย่อ","จำนวนวัน"],
                    )
                    .properties(height=320, title="แนวโน้มสถานะรายเดือน")
                )
                st.altair_chart(trend_att_chart, use_container_width=True)

                # Stacked bar เพื่อดูสัดส่วน
                stack_chart = (
                    alt.Chart(df_trend_att)
                    .mark_bar()
                    .encode(
                        x=alt.X("เดือน:O", sort=None),
                        y=alt.Y("จำนวนวัน:Q", stack="normalize", title="สัดส่วน (%)"),
                        color=alt.Color("สถานะย่อ:N", legend=alt.Legend(orient="bottom")),
                        tooltip=["เดือน","สถานะย่อ","จำนวนวัน"],
                    )
                    .properties(height=220, title="สัดส่วนสถานะแต่ละเดือน (%)")
                )
                st.altair_chart(stack_chart, use_container_width=True)
            else:
                st.info("ไม่มีข้อมูลสำหรับสถานะที่เลือก")

# ============================================================
# [R4] 📅 ปฏิทินกลาง — Heatmap
# ============================================================
elif menu == "📅 ปฏิทินกลาง":
    st.markdown('<div class="section-header">📅 ปฏิทินกลางหน่วยงาน</div>', unsafe_allow_html=True)

    with st.spinner("กำลังโหลดข้อมูล..."):
        df_leave  = read_excel_from_drive(FILE_LEAVE)
        df_travel = read_excel_from_drive(FILE_TRAVEL)
        df_staff  = read_excel_from_drive(FILE_STAFF)
        df_leave, df_travel, _ = preprocess_dataframes(df_leave, df_travel, pd.DataFrame())
        all_names = get_active_staff(df_staff) or get_all_names_fallback(df_leave, df_travel, pd.DataFrame())

    today = dt.date.today()
    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        cal_month = st.selectbox(
            "เดือน",
            pd.date_range(f"{today.year-1}-01-01", f"{today.year+1}-12-31", freq="MS").strftime("%Y-%m").tolist(),
            index=pd.date_range(f"{today.year-1}-01-01", f"{today.year+1}-12-31", freq="MS").strftime("%Y-%m").tolist().index(today.strftime("%Y-%m")),
        )
    with col_f2:
        cal_group = st.selectbox("กลุ่มงาน (ว่าง = ทุกกลุ่ม)", ["ทุกกลุ่ม"] + STAFF_GROUPS)
    with col_f3:
        cal_names = st.multiselect("เลือกบุคลากร (ว่าง = ทุกคน)", all_names)

    m_start = pd.to_datetime(cal_month + "-01")
    m_end   = m_start + pd.offsets.MonthEnd(0)
    date_range = pd.date_range(m_start, m_end, freq="D")

    # Filter staff
    names_to_show = cal_names or all_names
    if cal_group != "ทุกกลุ่ม" and not df_staff.empty and "กลุ่มงาน" in df_staff.columns:
        grp_names = df_staff[df_staff["กลุ่มงาน"] == cal_group]["ชื่อ-สกุล"].tolist()
        names_to_show = [n for n in names_to_show if n in grp_names]

    cal_records = []
    for name in names_to_show:
        for d in date_range:
            status = "วันหยุด" if d.weekday() >= 5 else "ปฏิบัติงาน"

            ul = df_leave[df_leave["ชื่อ-สกุล"] == name] if not df_leave.empty else pd.DataFrame()
            if not ul.empty:
                ml = ul[(ul["วันที่เริ่ม"] <= d) & (ul["วันที่สิ้นสุด"] >= d)]
                if not ml.empty:
                    status = "ลา"

            ut = df_travel[df_travel["ชื่อ-สกุล"] == name] if not df_travel.empty else pd.DataFrame()
            if not ut.empty:
                mt = ut[(ut["วันที่เริ่ม"] <= d) & (ut["วันที่สิ้นสุด"] >= d)]
                if not mt.empty:
                    status = "ไปราชการ"

            cal_records.append({"ชื่อ-สกุล": name, "วันที่": d.strftime("%d"), "สถานะ": status, "วันที่เต็ม": d})

    if cal_records:
        df_cal = pd.DataFrame(cal_records)
        heatmap = alt.Chart(df_cal).mark_rect(stroke="white", strokeWidth=1).encode(
            x=alt.X("วันที่:O", title="วันที่", sort=None),
            y=alt.Y("ชื่อ-สกุล:N", title=""),
            color=alt.Color(
                "สถานะ:N",
                scale=alt.Scale(
                    domain=["ปฏิบัติงาน", "ลา", "ไปราชการ", "วันหยุด"],
                    range=["#22c55e", "#60a5fa", "#f59e0b", "#e2e8f0"],
                ),
                legend=alt.Legend(orient="bottom"),
            ),
            tooltip=["ชื่อ-สกุล", "วันที่เต็ม", "สถานะ"],
        ).properties(height=max(200, len(names_to_show) * 22), title=f"ปฏิทินการปฏิบัติงาน — {cal_month}")
        st.altair_chart(heatmap, use_container_width=True)

        # [N4] Alert: วันที่มีคนลา/ราชการมากผิดปกติ
        df_alert = df_cal[df_cal["สถานะ"].isin(["ลา","ไปราชการ"])].groupby("วันที่เต็ม")["ชื่อ-สกุล"].count().reset_index()
        df_alert.columns = ["วันที่", "จำนวนคน"]
        df_alert_risk = df_alert[df_alert["จำนวนคน"] >= max(3, len(names_to_show) * 0.3)]
        if not df_alert_risk.empty:
            st.warning(f"⚠️ พบ {len(df_alert_risk)} วันที่มีบุคลากรลา/ราชการพร้อมกัน ≥ {df_alert_risk['จำนวนคน'].min()} คน")
            st.dataframe(df_alert_risk, use_container_width=True)
    else:
        st.info("ไม่มีข้อมูล")

# ============================================================
# 🧭 บันทึกไปราชการ
# ============================================================
elif menu == "🧭 บันทึกไปราชการ":
    st.markdown('<div class="section-header">🧭 บันทึกการเดินทางไปราชการ</div>', unsafe_allow_html=True)

    with st.spinner("กำลังโหลดข้อมูล..."):
        df_travel = read_excel_from_drive(FILE_TRAVEL)
        df_leave  = read_excel_from_drive(FILE_LEAVE)
        df_att    = read_excel_from_drive(FILE_ATTEND)
        df_staff  = read_excel_from_drive(FILE_STAFF)
        df_leave, df_travel, df_att = preprocess_dataframes(df_leave, df_travel, df_att)
        ALL_NAMES = get_active_staff(df_staff) or get_all_names_fallback(df_leave, df_travel, df_att)

    with st.form("form_travel"):
        col1, col2 = st.columns(2)
        with col1:
            group_job = st.selectbox("กลุ่มงาน", STAFF_GROUPS)
            project   = st.text_input("ชื่อโครงการ/กิจกรรม *", placeholder="ระบุชื่อโครงการ")
            location  = st.text_input("สถานที่ *", placeholder="เช่น กรุงเทพฯ / โรงแรม...")
        with col2:
            d_start = st.date_input("วันที่เริ่ม *", value=dt.date.today())
            d_end   = st.date_input("วันที่สิ้นสุด *", value=dt.date.today())

        st.markdown("---")
        st.markdown("**👥 รายชื่อผู้เดินทาง**")
        selected_staff   = st.multiselect("เลือกจากระบบ", ALL_NAMES)
        extra_staff_text = st.text_area("เพิ่มชื่อที่ไม่มีในระบบ (คั่นด้วย , หรือขึ้นบรรทัดใหม่)")
        uploaded_pdf     = st.file_uploader("แนบเอกสารขออนุมัติ (PDF)", type=["pdf"])
        submitted = st.form_submit_button("💾 บันทึกข้อมูล", use_container_width=True, type="primary")

        if submitted:
            final_staff = list(selected_staff)
            if extra_staff_text:
                extras = [n.strip() for n in extra_staff_text.replace("\n", ",").split(",") if n.strip()]
                final_staff.extend(extras)
            final_staff = sorted(set(final_staff))

            errors = validate_travel_data(final_staff, project, location, d_start, d_end)
            if errors:
                for e in errors:
                    st.error(e)
            else:
                with st.status("กำลังบันทึก...", expanded=True) as status:
                    try:
                        link = "-"
                        if uploaded_pdf:
                            st.write("📤 อัปโหลดไฟล์...")
                            fid = get_or_create_folder(ATTACHMENT_FOLDER_NAME, FOLDER_ID)
                            if fid:
                                fn = f"TRAVEL_{dt.datetime.now().strftime('%Y%m%d_%H%M')}_{len(final_staff)}pax.pdf"
                                link = upload_pdf_to_drive(uploaded_pdf, fn, fid)

                        st.write("💾 บันทึกข้อมูล...")
                        ts   = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        days = count_weekdays(d_start, d_end)
                        new_rows = [
                            {"Timestamp": ts, "กลุ่มงาน": group_job, "ชื่อ-สกุล": p,
                             "เรื่อง/กิจกรรม": project, "สถานที่": location,
                             "วันที่เริ่ม": pd.to_datetime(d_start), "วันที่สิ้นสุด": pd.to_datetime(d_end),
                             "จำนวนวัน": days, "ไฟล์แนบ": link}
                            for p in final_staff
                        ]
                        backup_excel(FILE_TRAVEL, df_travel)
                        df_upd = pd.concat([df_travel, pd.DataFrame(new_rows)], ignore_index=True)

                        if write_excel_to_drive(FILE_TRAVEL, df_upd):
                            # [N2] LINE Notify
                            st.write("🔔 ส่งแจ้งเตือน LINE...")
                            msg = format_travel_notify(final_staff, project, location, d_start, d_end, days)
                            sent = send_line_notify(msg)
                            # [N3] Log
                            log_activity("ไปราชการ", f"{project} @ {location}", ", ".join(final_staff[:3]))

                            status.update(label=f"✅ บันทึกสำเร็จ ({len(final_staff)} ท่าน) {'| LINE ✓' if sent else ''}", state="complete")
                            st.toast(f"✅ บันทึกไปราชการสำเร็จ {len(final_staff)} ท่าน", icon="✅")
                            time.sleep(1)
                            st.rerun()
                        else:
                            status.update(label="❌ บันทึกล้มเหลว", state="error")
                    except Exception as e:
                        logger.error(f"travel form: {e}")
                        status.update(label=f"❌ {e}", state="error")

    st.divider()
    st.subheader("📋 รายการล่าสุด")
    if not df_travel.empty:
        cols = [c for c in ["Timestamp","ชื่อ-สกุล","เรื่อง/กิจกรรม","สถานที่","วันที่เริ่ม","วันที่สิ้นสุด"] if c in df_travel.columns]
        st.dataframe(df_travel[cols].tail(5), use_container_width=True)

# ============================================================
# 🕒 บันทึกการลา
# ============================================================
elif menu == "🕒 บันทึกการลา":
    st.markdown('<div class="section-header">🕒 บันทึกการลา</div>', unsafe_allow_html=True)

    with st.spinner("กำลังโหลดข้อมูล..."):
        df_leave  = read_excel_from_drive(FILE_LEAVE)
        df_travel = read_excel_from_drive(FILE_TRAVEL)
        df_att    = read_excel_from_drive(FILE_ATTEND)
        df_staff  = read_excel_from_drive(FILE_STAFF)
        df_leave, df_travel, df_att = preprocess_dataframes(df_leave, df_travel, df_att)
        ALL_NAMES = get_active_staff(df_staff) or get_all_names_fallback(df_leave, df_travel, df_att)

    with st.form("form_leave"):
        col1, col2 = st.columns(2)
        with col1:
            l_name  = st.selectbox("ชื่อ-สกุล *", ALL_NAMES)
            l_group = st.selectbox("กลุ่มงาน", STAFF_GROUPS)
            l_type  = st.selectbox("ประเภทการลา *", LEAVE_TYPES)
        with col2:
            l_start  = st.date_input("วันที่เริ่มลา *", value=dt.date.today())
            l_end    = st.date_input("ถึงวันที่ *", value=dt.date.today())
            l_reason = st.text_area("เหตุผลการลา *", placeholder="ระบุเหตุผล (อย่างน้อย 5 ตัวอักษร)")

        l_file   = st.file_uploader("แนบใบลา (PDF)", type=["pdf"])
        l_submit = st.form_submit_button("💾 บันทึกการลา", use_container_width=True, type="primary")

        if l_submit:
            days_req = count_weekdays(l_start, l_end)
            errors   = validate_leave_data(l_name, l_start, l_end, l_reason, df_leave)

            # [S4] ตรวจสอบ quota
            quota_msg = check_leave_quota(l_name, l_type, days_req, df_leave, l_start.year) if l_name else None
            if quota_msg and quota_msg.startswith("❌"):
                errors.append(quota_msg)

            if errors:
                for e in errors:
                    st.error(e)
            else:
                if quota_msg:  # warning level
                    st.warning(quota_msg)

                with st.status("กำลังบันทึก...", expanded=True) as status:
                    try:
                        link = "-"
                        if l_file:
                            st.write("📤 อัปโหลดไฟล์...")
                            fid = get_or_create_folder(ATTACHMENT_FOLDER_NAME, FOLDER_ID)
                            if fid:
                                fn = f"LEAVE_{l_name}_{dt.datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
                                link = upload_pdf_to_drive(l_file, fn, fid)

                        st.write("💾 บันทึกข้อมูล...")
                        new_rec = {
                            "Timestamp": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "ชื่อ-สกุล": l_name, "กลุ่มงาน": l_group, "ประเภทการลา": l_type,
                            "วันที่เริ่ม": pd.to_datetime(l_start), "วันที่สิ้นสุด": pd.to_datetime(l_end),
                            "จำนวนวันลา": days_req, "เหตุผล": l_reason, "ไฟล์แนบ": link,
                        }
                        backup_excel(FILE_LEAVE, df_leave)
                        df_upd = pd.concat([df_leave, pd.DataFrame([new_rec])], ignore_index=True)

                        if write_excel_to_drive(FILE_LEAVE, df_upd):
                            # [N1] LINE Notify
                            st.write("🔔 ส่งแจ้งเตือน LINE...")
                            sent = send_line_notify(format_leave_notify(new_rec))
                            # [N3] Log
                            log_activity("การลา", f"{l_type} {days_req} วัน — {l_reason[:30]}", l_name)

                            status.update(label=f"✅ บันทึกสำเร็จ {'| LINE ✓' if sent else ''}", state="complete")
                            st.toast(f"✅ บันทึกการลาสำเร็จ ({l_type} {days_req} วัน)", icon="✅")
                            time.sleep(1)
                            st.rerun()
                        else:
                            status.update(label="❌ บันทึกล้มเหลว", state="error")
                    except Exception as e:
                        logger.error(f"leave form: {e}")
                        status.update(label=f"❌ {e}", state="error")

    st.divider()
    st.subheader("📋 รายการล่าสุด")
    if not df_leave.empty:
        cols = [c for c in ["Timestamp","ชื่อ-สกุล","ประเภทการลา","วันที่เริ่ม","วันที่สิ้นสุด","จำนวนวันลา"] if c in df_leave.columns]
        st.dataframe(df_leave[cols].tail(5), use_container_width=True)

# ============================================================
# [R1][S4] 📈 วันลาคงเหลือ
# ============================================================
elif menu == "📈 วันลาคงเหลือ":
    st.markdown('<div class="section-header">📈 สิทธิ์วันลาคงเหลือ</div>', unsafe_allow_html=True)

    with st.spinner("กำลังโหลดข้อมูล..."):
        df_leave = read_excel_from_drive(FILE_LEAVE)
        df_staff = read_excel_from_drive(FILE_STAFF)
        df_leave, _, _ = preprocess_dataframes(df_leave, pd.DataFrame(), pd.DataFrame())
        all_names = get_active_staff(df_staff) or get_all_names_fallback(df_leave, pd.DataFrame(), pd.DataFrame())

    selected_year = st.selectbox("ปี (พ.ศ.)", list(range(dt.date.today().year + 543, dt.date.today().year + 540, -1)))
    year_ad = selected_year - 543

    selected_person = st.selectbox("เลือกบุคลากร (ว่าง = ดูทุกคน)", ["— ทุกคน —"] + all_names)
    names_to_show = all_names if selected_person == "— ทุกคน —" else [selected_person]

    quota_rows = []
    for name in names_to_show:
        row = {"ชื่อ-สกุล": name}
        has_alert = False
        for ltype, quota in LEAVE_QUOTA.items():
            used      = get_leave_used(name, ltype, df_leave, year_ad)
            remaining = max(0, quota - used)
            indicator, _ = get_quota_status(used, quota)
            row[f"{ltype}_ใช้"] = used
            row[f"{ltype}_คงเหลือ"] = remaining
            row[f"{ltype}_สถานะ"] = indicator
            if used >= quota:
                has_alert = True
        row["⚠️"] = "🔴" if has_alert else ""
        quota_rows.append(row)

    df_quota = pd.DataFrame(quota_rows)

    # แสดง visual quota สำหรับคนที่เลือก
    if selected_person != "— ทุกคน —":
        st.subheader(f"📊 สิทธิ์ลาของ {selected_person} ปี {selected_year}")
        cols_q = st.columns(len(LEAVE_QUOTA))
        for i, (ltype, quota) in enumerate(LEAVE_QUOTA.items()):
            used = get_leave_used(selected_person, ltype, df_leave, year_ad)
            remaining = max(0, quota - used)
            indicator, badge_cls = get_quota_status(used, quota)
            with cols_q[i % len(cols_q)]:
                st.markdown(f"**{ltype}**")
                st.markdown(f"{indicator} ใช้ **{used}** / {quota} วัน")
                st.markdown(quota_bar_html(used, quota), unsafe_allow_html=True)
                st.markdown(f'<span class="{badge_cls}">คงเหลือ {remaining} วัน</span>', unsafe_allow_html=True)
        st.divider()

    st.subheader("📋 ตารางสรุปทุกคน")
    # แสดงเฉพาะคอลัมน์สำคัญ
    display_cols = ["ชื่อ-สกุล", "⚠️"]
    for ltype in LEAVE_QUOTA:
        if f"{ltype}_ใช้" in df_quota.columns:
            display_cols += [f"{ltype}_ใช้", f"{ltype}_คงเหลือ"]
    st.dataframe(df_quota[display_cols], use_container_width=True, height=400)

    # Export
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        df_quota.to_excel(w, index=False, sheet_name=f"โควต้าลา_{selected_year}")
    buf.seek(0)
    st.download_button("📥 ดาวน์โหลด Excel", buf, f"LeaveQuota_{selected_year}.xlsx", mime=EXCEL_MIME)

# ============================================================
# [S1][S3] 👤 จัดการบุคลากร
# ============================================================
elif menu == "👤 จัดการบุคลากร":
    st.markdown('<div class="section-header">👤 จัดการฐานข้อมูลบุคลากร</div>', unsafe_allow_html=True)

    with st.spinner("กำลังโหลด..."):
        df_staff = read_excel_from_drive(FILE_STAFF)

    if df_staff.empty:
        df_staff = pd.DataFrame(columns=STAFF_MASTER_COLS)

    tab_list, tab_add, tab_edit = st.tabs(["📋 รายชื่อทั้งหมด", "➕ เพิ่มบุคลากร", "✏️ แก้ไข / ปิดใช้งาน"])

    # Tab 1: แสดงรายชื่อ
    with tab_list:
        col_s, col_f = st.columns([1,2])
        with col_s:
            filter_status = st.selectbox("สถานะ", ["ทุกสถานะ","ปฏิบัติงาน","ลาออก","ยืมตัว"])
        with col_f:
            filter_group = st.selectbox("กลุ่มงาน", ["ทุกกลุ่ม"] + STAFF_GROUPS)

        df_show = df_staff.copy()
        if filter_status != "ทุกสถานะ" and "สถานะ" in df_show.columns:
            df_show = df_show[df_show["สถานะ"] == filter_status]
        if filter_group != "ทุกกลุ่ม" and "กลุ่มงาน" in df_show.columns:
            df_show = df_show[df_show["กลุ่มงาน"] == filter_group]

        st.caption(f"แสดง {len(df_show)} รายการ")

        def badge_status(val):
            m = {"ปฏิบัติงาน":"badge-green","ลาออก":"badge-red","ยืมตัว":"badge-yellow"}
            cls = m.get(str(val),"badge-gray")
            return f'<span class="{cls}">{val}</span>'

        if not df_show.empty:
            # [U3] Badge แสดงสถานะ
            df_display = df_show.copy()
            if "สถานะ" in df_display.columns:
                df_display["สถานะ (badge)"] = df_display["สถานะ"].apply(badge_status)
            st.dataframe(df_show, use_container_width=True, height=420)

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            df_staff.to_excel(w, index=False)
        buf.seek(0)
        st.download_button("📥 Export รายชื่อ", buf, "staff_master.xlsx", mime=EXCEL_MIME)

    # Tab 2: เพิ่มบุคลากรใหม่
    with tab_add:
        with st.form("form_add_staff"):
            c1, c2 = st.columns(2)
            with c1:
                s_name   = st.text_input("ชื่อ-สกุล *", placeholder="นายสมชาย ใจดี")
                s_group  = st.selectbox("กลุ่มงาน *", STAFF_GROUPS)
                s_pos    = st.text_input("ตำแหน่ง", placeholder="นักวิชาการสาธารณสุข")
            with c2:
                s_type   = st.selectbox("ประเภทบุคลากร", ["ข้าราชการ","พนักงานราชการ","ลูกจ้างประจำ","จ้างเหมา"])
                s_start  = st.date_input("วันเริ่มปฏิบัติงาน", value=dt.date.today())
                s_status = st.selectbox("สถานะ", ["ปฏิบัติงาน","ยืมตัว"])

            s_submit = st.form_submit_button("➕ เพิ่มบุคลากร", use_container_width=True, type="primary")
            if s_submit:
                if not s_name.strip():
                    st.error("❌ กรุณาระบุชื่อ-สกุล")
                elif not df_staff.empty and s_name.strip() in df_staff["ชื่อ-สกุล"].values:
                    st.error("❌ ชื่อนี้มีอยู่ในระบบแล้ว")
                else:
                    new_staff = {
                        "ชื่อ-สกุล": s_name.strip(), "กลุ่มงาน": s_group,
                        "ตำแหน่ง": s_pos, "ประเภทบุคลากร": s_type,
                        "วันเริ่มงาน": str(s_start), "สถานะ": s_status,
                    }
                    df_staff = pd.concat([df_staff, pd.DataFrame([new_staff])], ignore_index=True)
                    if write_excel_to_drive(FILE_STAFF, df_staff):
                        log_activity("เพิ่มบุคลากร", f"เพิ่ม {s_name} ({s_group})", s_name)
                        st.toast(f"✅ เพิ่ม {s_name} สำเร็จ", icon="✅")
                        st.rerun()

    # Tab 3: แก้ไขหรือปิดใช้งาน
    with tab_edit:
        if df_staff.empty:
            st.info("ยังไม่มีข้อมูลบุคลากร")
        else:
            edit_name = st.selectbox("เลือกบุคลากรที่ต้องการแก้ไข", df_staff["ชื่อ-สกุล"].tolist())
            row_idx   = df_staff[df_staff["ชื่อ-สกุล"] == edit_name].index

            if len(row_idx) > 0:
                idx = row_idx[0]
                with st.form("form_edit_staff"):
                    c1, c2 = st.columns(2)
                    with c1:
                        e_group  = st.selectbox("กลุ่มงาน", STAFF_GROUPS, index=STAFF_GROUPS.index(df_staff.at[idx,"กลุ่มงาน"]) if "กลุ่มงาน" in df_staff.columns and df_staff.at[idx,"กลุ่มงาน"] in STAFF_GROUPS else 0)
                        e_pos    = st.text_input("ตำแหน่ง", value=str(df_staff.at[idx,"ตำแหน่ง"]) if "ตำแหน่ง" in df_staff.columns else "")
                    with c2:
                        e_type   = st.selectbox("ประเภทบุคลากร", ["ข้าราชการ","พนักงานราชการ","ลูกจ้างประจำ","จ้างเหมา"])
                        e_status_opts = ["ปฏิบัติงาน","ลาออก","ยืมตัว"]
                        cur_status = str(df_staff.at[idx,"สถานะ"]) if "สถานะ" in df_staff.columns else "ปฏิบัติงาน"
                        e_status = st.selectbox("สถานะ", e_status_opts, index=e_status_opts.index(cur_status) if cur_status in e_status_opts else 0)

                    e_submit = st.form_submit_button("✅ บันทึกการแก้ไข", use_container_width=True)
                    if e_submit:
                        df_staff.at[idx,"กลุ่มงาน"]       = e_group
                        df_staff.at[idx,"ตำแหน่ง"]         = e_pos
                        df_staff.at[idx,"ประเภทบุคลากร"]   = e_type
                        df_staff.at[idx,"สถานะ"]           = e_status
                        if write_excel_to_drive(FILE_STAFF, df_staff):
                            log_activity("แก้ไขบุคลากร", f"อัปเดตข้อมูล {edit_name} สถานะ→{e_status}", edit_name)
                            st.toast(f"✅ อัปเดต {edit_name} สำเร็จ", icon="✅")
                            st.rerun()

# ============================================================
# [N3] 🔔 กิจกรรมล่าสุด (Activity Feed)
# ============================================================
elif menu == "🔔 กิจกรรมล่าสุด":
    st.markdown('<div class="section-header">🔔 กิจกรรมล่าสุดในระบบ</div>', unsafe_allow_html=True)

    with st.spinner("กำลังโหลด..."):
        df_log = read_excel_from_drive(FILE_NOTIFY)

    if df_log.empty:
        st.info("ยังไม่มีกิจกรรมในระบบ กิจกรรมจะถูกบันทึกเมื่อมีการบันทึกการลาหรือไปราชการ")
    else:
        df_log = df_log.sort_values("Timestamp", ascending=False).head(50)

        col_f1, col_f2 = st.columns(2)
        with col_f1:
            filter_type = st.selectbox("กรองตามประเภท", ["ทั้งหมด"] + df_log["ประเภท"].dropna().unique().tolist())
        with col_f2:
            search_name = st.text_input("ค้นหาชื่อ")

        df_show = df_log.copy()
        if filter_type != "ทั้งหมด":
            df_show = df_show[df_show["ประเภท"] == filter_type]
        if search_name:
            df_show = df_show[df_show["ผู้เกี่ยวข้อง"].str.contains(search_name, na=False)]

        TYPE_ICONS = {"การลา": "🕒", "ไปราชการ": "✈️", "เพิ่มบุคลากร": "➕", "แก้ไขบุคลากร": "✏️"}

        for _, row in df_show.iterrows():
            icon = TYPE_ICONS.get(str(row.get("ประเภท","")), "📌")
            ts   = str(row.get("Timestamp",""))[:16]
            st.markdown(
                f'<div class="activity-item">'
                f'{icon} <b>{row.get("ประเภท","")}</b> — {row.get("รายละเอียด","")}'
                f'<br><small style="color:#94a3b8;">👤 {row.get("ผู้เกี่ยวข้อง","")} &nbsp;|&nbsp; ⏰ {ts}</small>'
                f'</div>',
                unsafe_allow_html=True,
            )

# ============================================================
# ⚙️ ผู้ดูแลระบบ
# ============================================================
elif menu == "⚙️ ผู้ดูแลระบบ":
    st.markdown('<div class="section-header">⚙️ ผู้ดูแลระบบ</div>', unsafe_allow_html=True)
    password = st.text_input("🔑 รหัสผ่าน Admin", type="password")

    if password and check_admin_password(password):
        st.success("✅ เข้าสู่ระบบสำเร็จ")

        with st.spinner("กำลังโหลด..."):
            df_leave  = read_excel_from_drive(FILE_LEAVE)
            df_travel = read_excel_from_drive(FILE_TRAVEL)
            df_att    = read_excel_from_drive(FILE_ATTEND)
            df_staff  = read_excel_from_drive(FILE_STAFF)

        tab1, tab2, tab3, tab4, tab5, tab6, tab_hol = st.tabs([
            "📂 ไฟล์ลา", "📂 ไฟล์ราชการ",
            "📂 ไฟล์สแกนนิ้ว", "📂 ไฟล์บุคลากร",
            "🔧 ตั้งค่าระบบ", "👆 คีย์ลืมสแกนนิ้ว", "🎌 วันหยุดพิเศษ",
        ])

        def admin_file_panel(df, filename, tab_obj):
            with tab_obj:
                st.subheader(f"ไฟล์: {filename}")
                if df.empty:
                    st.warning("⚠️ ไม่มีข้อมูล")
                else:
                    st.dataframe(df.head(20), use_container_width=True)
                    st.caption(f"ทั้งหมด {len(df)} แถว | {len(df.columns)} คอลัมน์")

                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    if not df.empty:
                        buf = io.BytesIO()
                        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
                            df.to_excel(w, index=False)
                        st.download_button(f"⬇️ Excel", buf.getvalue(), filename, use_container_width=True)
                with col_d2:
                    if not df.empty:
                        csv = df.to_csv(index=False).encode("utf-8-sig")
                        st.download_button("⬇️ CSV", csv, filename.replace(".xlsx",".csv"), "text/csv", use_container_width=True)

                st.divider()
                st.warning("⚠️ การอัปโหลดจะเขียนทับข้อมูลเดิมทั้งหมด")
                up = st.file_uploader(f"อัปโหลดทับ {filename}", type=["xlsx"], key=f"up_{filename}")
                if up:
                    try:
                        new_df = pd.read_excel(up)
                        st.info(f"ไฟล์: {len(new_df)} แถว, {len(new_df.columns)} คอลัมน์")
                        st.dataframe(new_df.head(3))
                        if st.button(f"✅ ยืนยันอัปโหลด", key=f"confirm_{filename}", type="primary"):
                            backup_excel(filename, df)
                            if write_excel_to_drive(filename, new_df):
                                st.toast("✅ อัปเดตสำเร็จ", icon="✅")
                                time.sleep(1)
                                st.rerun()
                    except Exception as e:
                        st.error(f"❌ อ่านไฟล์ไม่ได้: {e}")

        admin_file_panel(df_leave,  FILE_LEAVE,  tab1)
        admin_file_panel(df_travel, FILE_TRAVEL, tab2)
        admin_file_panel(df_att,    FILE_ATTEND, tab3)
        admin_file_panel(df_staff,  FILE_STAFF,  tab4)

        with tab5:
            st.subheader("🔧 ตั้งค่าระบบ")
            st.info("ค่าต่อไปนี้ควรตั้งใน `.streamlit/secrets.toml` เพื่อความปลอดภัย")
            st.code("""
# .streamlit/secrets.toml

[gcp_service_account]
# ... Google Service Account JSON

admin_password = "รหัสผ่าน Admin ของคุณ"
line_notify_token = "Token จาก https://notify-bot.line.me/my/"
            """)
            st.subheader("📊 โควต้าวันลา (ปัจจุบัน)")
            df_quota_cfg = pd.DataFrame(
                [{"ประเภทการลา": k, "โควต้า (วัน/ปี)": v} for k, v in LEAVE_QUOTA.items()]
            )
            st.dataframe(df_quota_cfg, use_container_width=True)
            st.caption("หากต้องการเปลี่ยนโควต้า ให้แก้ไขค่า LEAVE_QUOTA ในไฟล์ app.py")

        # ============================================================
        # 👆 Tab 6 — คีย์ลืมสแกนนิ้ว
        # ============================================================
        with tab6:
            st.subheader("👆 บันทึกเวลาทำการสำหรับผู้ที่ลืมสแกนนิ้ว")
            st.info(
                "ใช้สำหรับกรณีที่เครื่องสแกนขัดข้อง หรือบุคลากรลืมสแกนนิ้ว "
                "ข้อมูลที่คีย์จะถูกเพิ่มเข้าไปใน `attendance_report.xlsx` "
                "พร้อมหมายเหตุว่า **Admin เพิ่มเอง**"
            )

            # ─── โหลดรายชื่อ ───
            all_staff_names = get_active_staff(df_staff) or get_all_names_fallback(
                df_leave, df_travel, df_att
            )

            # ─── ตรวจสอบ column ที่มีในไฟล์สแกน ───
            ATT_REQUIRED_COLS = ["ชื่อ-สกุล", "วันที่", "เวลาเข้า", "เวลาออก", "หมายเหตุ"]
            if df_att.empty:
                df_att = pd.DataFrame(columns=ATT_REQUIRED_COLS)
            for col in ATT_REQUIRED_COLS:
                if col not in df_att.columns:
                    df_att[col] = ""

            st.markdown("---")

            # ─── ฟอร์มคีย์ข้อมูล ───
            with st.form("form_manual_scan"):
                col_f1, col_f2 = st.columns(2)

                with col_f1:
                    ms_name = st.selectbox(
                        "ชื่อ-สกุล *",
                        all_staff_names,
                        help="เลือกชื่อบุคลากรที่ลืมสแกน",
                    )
                    ms_date = st.date_input(
                        "วันที่ลืมสแกน *",
                        value=dt.date.today(),
                        max_value=dt.date.today(),
                        help="ไม่สามารถคีย์วันในอนาคตได้",
                    )

                with col_f2:
                    ms_time_in  = st.time_input(
                        "เวลาเข้างาน *",
                        value=dt.time(8, 30),
                        step=60,
                    )
                    ms_time_out = st.time_input(
                        "เวลาออกงาน *",
                        value=dt.time(16, 30),
                        step=60,
                    )

                ms_note = st.text_input(
                    "หมายเหตุเพิ่มเติม",
                    value="Admin คีย์แทน — ลืมสแกนนิ้ว",
                    help="ระบบจะเพิ่มชื่อ Admin และเวลาที่คีย์ให้อัตโนมัติ",
                )

                ms_submit = st.form_submit_button(
                    "💾 บันทึกข้อมูลสแกนนิ้ว",
                    use_container_width=True,
                    type="primary",
                )

                if ms_submit:
                    # ─── Validation ───
                    scan_errors: List[str] = []
                    if not ms_name:
                        scan_errors.append("❌ กรุณาเลือกชื่อ")
                    if ms_time_in >= ms_time_out:
                        scan_errors.append("❌ เวลาเข้างานต้องน้อยกว่าเวลาออกงาน")

                    # เช็คว่ามีข้อมูลวันนั้นซ้ำอยู่แล้วหรือไม่
                    ms_date_ts = pd.to_datetime(ms_date)
                    existing = df_att[
                        (df_att["ชื่อ-สกุล"].astype(str).str.strip() == ms_name)
                        & (pd.to_datetime(df_att["วันที่"], errors="coerce").dt.normalize() == ms_date_ts)
                    ]
                    if not existing.empty:
                        scan_errors.append(
                            f"⚠️ มีข้อมูลสแกนของ {ms_name} วันที่ {ms_date} อยู่แล้ว "
                            f"(เวลาเข้า: {existing.iloc[0].get('เวลาเข้า','?')} "
                            f"เวลาออก: {existing.iloc[0].get('เวลาออก','?')}) "
                            f"— หากต้องการแก้ไข ให้ลบแถวเดิมก่อนในแท็บ '📂 ไฟล์สแกนนิ้ว'"
                        )

                    if scan_errors:
                        for err in scan_errors:
                            st.error(err)
                    else:
                        with st.status("กำลังบันทึก...", expanded=True) as status:
                            try:
                                note_full = (
                                    f"{ms_note} | คีย์โดย Admin "
                                    f"{dt.datetime.now().strftime('%d/%m/%Y %H:%M')}"
                                )
                                new_scan_row = {
                                    "ชื่อ-สกุล":  ms_name,
                                    "วันที่":      pd.to_datetime(ms_date),
                                    "เวลาเข้า":    ms_time_in.strftime("%H:%M"),
                                    "เวลาออก":     ms_time_out.strftime("%H:%M"),
                                    "หมายเหตุ":    note_full,
                                }
                                backup_excel(FILE_ATTEND, df_att)
                                df_att_upd = pd.concat(
                                    [df_att, pd.DataFrame([new_scan_row])],
                                    ignore_index=True,
                                )
                                # เรียงลำดับตามชื่อ + วันที่
                                df_att_upd = df_att_upd.sort_values(
                                    ["ชื่อ-สกุล", "วันที่"], ignore_index=True
                                )

                                if write_excel_to_drive(FILE_ATTEND, df_att_upd):
                                    log_activity(
                                        "คีย์สแกนนิ้ว",
                                        f"Admin คีย์ {ms_date} เข้า {ms_time_in.strftime('%H:%M')} ออก {ms_time_out.strftime('%H:%M')}",
                                        ms_name,
                                    )
                                    df_att = df_att_upd  # อัปเดต local copy
                                    status.update(label="✅ บันทึกสำเร็จ!", state="complete")
                                    st.toast(
                                        f"✅ บันทึกสแกนนิ้วของ {ms_name} วันที่ {ms_date} สำเร็จ",
                                        icon="✅",
                                    )
                                    time.sleep(1)
                                    st.rerun()
                                else:
                                    status.update(label="❌ บันทึกล้มเหลว", state="error")
                            except Exception as e:
                                logger.error(f"manual_scan error: {e}")
                                status.update(label=f"❌ {e}", state="error")

            st.divider()

            # ─── ค้นหาและตรวจสอบข้อมูลสแกนรายคน ───
            st.subheader("🔍 ตรวจสอบข้อมูลสแกนรายคน")

            col_s1, col_s2, col_s3 = st.columns([2, 1, 1])
            with col_s1:
                search_att_name = st.selectbox(
                    "เลือกชื่อบุคลากร", all_staff_names, key="search_att_name"
                )
            with col_s2:
                search_att_month_opts = (
                    sorted(
                        df_att["วันที่"]
                        .dropna()
                        .pipe(lambda s: pd.to_datetime(s, errors="coerce"))
                        .dt.strftime("%Y-%m")
                        .dropna()
                        .unique()
                        .tolist()
                    )
                    if not df_att.empty
                    else [dt.date.today().strftime("%Y-%m")]
                )
                search_att_month = st.selectbox(
                    "เดือน",
                    search_att_month_opts,
                    index=len(search_att_month_opts) - 1,
                    key="search_att_month",
                )
            with col_s3:
                st.write("")
                st.write("")
                show_all_flag = st.checkbox("แสดงทุกเดือน", key="show_all_att")

            # กรองข้อมูลสแกน
            if not df_att.empty:
                df_att_view = df_att[
                    df_att["ชื่อ-สกุล"].astype(str).str.strip() == search_att_name
                ].copy()
                df_att_view["วันที่"] = pd.to_datetime(df_att_view["วันที่"], errors="coerce")

                if not show_all_flag:
                    df_att_view = df_att_view[
                        df_att_view["วันที่"].dt.strftime("%Y-%m") == search_att_month
                    ]

                df_att_view = df_att_view.sort_values("วันที่", ascending=False)

                # เน้นแถวที่ Admin คีย์เอง
                def highlight_manual(row):
                    note = str(row.get("หมายเหตุ", ""))
                    if "Admin คีย์แทน" in note or "Admin คีย์" in note:
                        return ["background-color:#fef9c3"] * len(row)
                    return [""] * len(row)

                if df_att_view.empty:
                    st.info(f"ไม่พบข้อมูลสแกนของ {search_att_name} ในเดือน {search_att_month}")
                else:
                    st.caption(
                        f"พบ {len(df_att_view)} รายการ  |  "
                        f"🟡 = รายการที่ Admin คีย์เอง"
                    )
                    st.dataframe(
                        df_att_view.style.apply(highlight_manual, axis=1),
                        use_container_width=True,
                        height=350,
                    )

                    # ─── ลบรายการ (กรณีคีย์ผิด) ───
                    st.divider()
                    st.subheader("🗑️ ลบรายการที่คีย์ผิด")
                    st.warning(
                        "⚠️ ลบได้เฉพาะรายการที่ **Admin คีย์เอง** เท่านั้น "
                        "(แถวที่มีคำว่า 'Admin คีย์แทน' ในหมายเหตุ)"
                    )

                    admin_rows = df_att_view[
                        df_att_view["หมายเหตุ"].astype(str).str.contains(
                            "Admin คีย์", na=False
                        )
                    ].copy()

                    if admin_rows.empty:
                        st.info("ไม่มีรายการที่ Admin คีย์เองในช่วงที่เลือก")
                    else:
                        # สร้างป้ายให้เลือก
                        admin_rows["label"] = admin_rows.apply(
                            lambda r: (
                                f"{r['วันที่'].strftime('%d/%m/%Y') if pd.notna(r['วันที่']) else '?'} "
                                f"— เข้า {r.get('เวลาเข้า','?')} "
                                f"ออก {r.get('เวลาออก','?')}"
                            ),
                            axis=1,
                        )
                        del_label = st.selectbox(
                            "เลือกรายการที่ต้องการลบ",
                            admin_rows["label"].tolist(),
                            key="del_att_select",
                        )
                        del_row = admin_rows[admin_rows["label"] == del_label]

                        if st.button(
                            f"🗑️ ลบรายการนี้",
                            key="btn_del_att",
                            type="primary",
                        ):
                            if not del_row.empty:
                                idx_to_drop = del_row.index.tolist()
                                df_att_new = df_att.drop(index=idx_to_drop).reset_index(drop=True)
                                backup_excel(FILE_ATTEND, df_att)
                                if write_excel_to_drive(FILE_ATTEND, df_att_new):
                                    log_activity(
                                        "ลบสแกนนิ้ว",
                                        f"Admin ลบรายการ {del_label}",
                                        search_att_name,
                                    )
                                    df_att = df_att_new
                                    st.toast("✅ ลบรายการสำเร็จ", icon="🗑️")
                                    time.sleep(1)
                                    st.rerun()
            else:
                st.info("ยังไม่มีข้อมูลสแกนในระบบ")

        # ============================================================
        # 🎌 Tab: วันหยุดพิเศษ
        # ============================================================
        with tab_hol:
            st.subheader("🎌 จัดการวันหยุดพิเศษ / วันหยุดราชการ")
            st.caption(
                "วันหยุดที่กำหนดในหน้านี้จะถูกนำไปใช้ใน: "
                "① ตรวจสอบการปฏิบัติงาน (แสดงสถานะวันหยุด) "
                "② คำนวณวันลา/ราชการ (หักออกจากวันทำการ) "
                "③ ปฏิทินกลาง"
            )

            hol_yr_opts = list(range(dt.date.today().year + 1, dt.date.today().year - 3, -1))
            hol_col1, hol_col2 = st.columns([1, 3])
            with hol_col1:
                hol_view_year = st.selectbox("ดูปี (พ.ศ.)", [y + 543 for y in hol_yr_opts], key="hol_view_year")
                hol_view_year_ad = hol_view_year - 543

            with hol_col2:
                hol_show_fixed = st.checkbox("แสดงวันหยุดราชการตายตัว (กำหนดโดยระบบ)", value=True, key="hol_show_fixed")

            # โหลดข้อมูล
            df_hol_custom = load_holidays_raw()
            df_hol_fixed  = get_fixed_holidays_for_year(hol_view_year_ad)

            if hol_show_fixed:
                df_hol_display = load_holidays_all(hol_view_year_ad)
            else:
                df_hol_display = df_hol_custom[
                    pd.to_datetime(df_hol_custom["วันที่"], errors="coerce").dt.year == hol_view_year_ad
                ] if not df_hol_custom.empty else pd.DataFrame(columns=HOLIDAY_COLS)

            # ─── ตารางแสดงวันหยุด ──────────────────────────────
            st.markdown(f"#### 📋 วันหยุดทั้งหมดปี พ.ศ. {hol_view_year}")

            # นับจำนวนวันหยุดที่ตกในวันทำการ (จ-ศ)
            hol_dates_yr = get_holiday_dates(hol_view_year_ad)
            hol_workdays = [d for d in hol_dates_yr if d.weekday() < 5]
            hol_weekend  = [d for d in hol_dates_yr if d.weekday() >= 5]

            hkc1, hkc2, hkc3 = st.columns(3)
            hkc1.metric("📅 วันหยุดทั้งหมด",       f"{len(hol_dates_yr)} วัน")
            hkc2.metric("📌 ตกในวันทำการ (จ-ศ)",   f"{len(hol_workdays)} วัน", help="วันที่มีผลหักจากวันทำการ")
            hkc3.metric("🏖️ ตกวันเสาร์-อาทิตย์",   f"{len(hol_weekend)} วัน",  help="ไม่มีผลต่อวันทำการ")

            st.divider()

            if not df_hol_display.empty:
                # เพิ่มคอลัมน์ช่วยดู
                df_show_hol = df_hol_display.copy()
                df_show_hol["วันที่"] = pd.to_datetime(df_show_hol["วันที่"], errors="coerce")
                df_show_hol["วัน"] = df_show_hol["วันที่"].dt.strftime("%A").map({
                    "Monday":"จันทร์","Tuesday":"อังคาร","Wednesday":"พุธ",
                    "Thursday":"พฤหัสบดี","Friday":"ศุกร์","Saturday":"เสาร์","Sunday":"อาทิตย์",
                })
                df_show_hol["กระทบวันทำการ"] = df_show_hol["วันที่"].dt.weekday.apply(
                    lambda w: "✅ ใช่" if w < 5 else "—"
                )
                df_show_hol["วันที่"] = df_show_hol["วันที่"].dt.strftime("%d/%m/%Y")

                def hol_row_color(row):
                    if str(row.get("หมายเหตุ","")).startswith("กำหนดโดยระบบ"):
                        return ["background-color:#f0f4ff"] * len(row)
                    return ["background-color:#fffde7"] * len(row)

                st.dataframe(
                    df_show_hol[["วันที่","วัน","ชื่อวันหยุด","ประเภท","กระทบวันทำการ","หมายเหตุ"]]
                    .style.apply(hol_row_color, axis=1),
                    use_container_width=True,
                    height=320,
                )
                st.caption("🔵 น้ำเงินอ่อน = วันหยุดราชการตายตัว  |  🟡 เหลืองอ่อน = วันหยุดที่ Admin เพิ่มเอง")
            else:
                st.info(f"ยังไม่มีวันหยุดพิเศษสำหรับปี {hol_view_year}")

            st.divider()

            # ─── ฟอร์มเพิ่มวันหยุด ─────────────────────────────
            st.markdown("#### ➕ เพิ่มวันหยุดพิเศษ")
            st.info("วันหยุดราชการตายตัวจะถูกเพิ่มให้อัตโนมัติ ไม่ต้องกรอกซ้ำ")

            with st.form("form_add_holiday"):
                ha_col1, ha_col2 = st.columns(2)
                with ha_col1:
                    ha_date   = st.date_input("วันที่ *", value=dt.date.today(), key="ha_date")
                    ha_name   = st.text_input("ชื่อวันหยุด *", placeholder="เช่น วันพ่อแห่งชาติ, วันหยุดชดเชย", key="ha_name")
                with ha_col2:
                    ha_type   = st.selectbox("ประเภท *", HOLIDAY_TYPE_OPTIONS, key="ha_type")
                    ha_note   = st.text_input("หมายเหตุ", placeholder="ข้อมูลเพิ่มเติม (ถ้ามี)", key="ha_note")

                ha_submit = st.form_submit_button("➕ เพิ่มวันหยุด", use_container_width=True, type="primary")

                if ha_submit:
                    ha_errors: List[str] = []
                    if not ha_name.strip():
                        ha_errors.append("❌ กรุณาระบุชื่อวันหยุด")

                    # ตรวจซ้ำกับวันหยุดที่ Admin เพิ่มเองแล้ว
                    if not df_hol_custom.empty:
                        existing_dates = pd.to_datetime(df_hol_custom["วันที่"], errors="coerce").dt.date.tolist()
                        if ha_date in existing_dates:
                            ha_errors.append(f"❌ วันที่ {ha_date.strftime('%d/%m/%Y')} มีในระบบแล้ว")

                    # ตรวจซ้ำกับวันหยุดราชการตายตัว
                    fixed_dates = [dt.date(ha_date.year, m, d) for m, d, _ in FIXED_THAI_HOLIDAYS
                                   if _can_make_date(ha_date.year, m, d)]
                    if ha_date in fixed_dates:
                        ha_errors.append(
                            f"⚠️ วันที่ {ha_date.strftime('%d/%m/%Y')} ตรงกับวันหยุดราชการตายตัวที่ระบบกำหนดไว้แล้ว "
                            f"— ไม่จำเป็นต้องเพิ่มซ้ำ"
                        )

                    if ha_errors:
                        for e in ha_errors:
                            st.error(e)
                    else:
                        new_hol = {
                            "วันที่":      pd.Timestamp(ha_date),
                            "ชื่อวันหยุด": ha_name.strip(),
                            "ประเภท":     ha_type,
                            "หมายเหตุ":   ha_note.strip(),
                        }
                        df_hol_new = pd.concat(
                            [df_hol_custom, pd.DataFrame([new_hol])], ignore_index=True
                        ).sort_values("วันที่").reset_index(drop=True)

                        if write_excel_to_drive(FILE_HOLIDAYS, df_hol_new):
                            log_activity(
                                "เพิ่มวันหยุดพิเศษ",
                                f"{ha_name} ({ha_date.strftime('%d/%m/%Y')}) ประเภท {ha_type}",
                                "Admin",
                            )
                            st.toast(f"✅ เพิ่ม '{ha_name}' วันที่ {ha_date.strftime('%d/%m/%Y')} สำเร็จ", icon="🎌")
                            st.cache_data.clear()
                            time.sleep(0.5)
                            st.rerun()

            st.divider()

            # ─── ลบวันหยุด (เฉพาะที่ Admin เพิ่มเอง) ──────────
            st.markdown("#### 🗑️ ลบวันหยุดพิเศษ")
            st.warning("⚠️ ลบได้เฉพาะวันหยุดที่ **Admin เพิ่มเอง** เท่านั้น — วันหยุดราชการตายตัวลบไม่ได้")

            # โหลดใหม่ (อาจเพิ่งเพิ่มไป)
            df_hol_custom_fresh = load_holidays_raw()
            if df_hol_custom_fresh.empty:
                st.info("ยังไม่มีวันหยุดพิเศษที่ Admin เพิ่มเอง")
            else:
                df_hol_del = df_hol_custom_fresh.copy()
                df_hol_del["วันที่"] = pd.to_datetime(df_hol_del["วันที่"], errors="coerce")
                df_hol_del["label"] = df_hol_del.apply(
                    lambda r: f"{r['วันที่'].strftime('%d/%m/%Y') if pd.notna(r['วันที่']) else '?'} — {r.get('ชื่อวันหยุด','')} ({r.get('ประเภท','')})",
                    axis=1,
                )

                del_hol_col1, del_hol_col2 = st.columns([3, 1])
                with del_hol_col1:
                    del_hol_label = st.selectbox(
                        "เลือกวันหยุดที่ต้องการลบ",
                        df_hol_del["label"].tolist(),
                        key="del_hol_select",
                    )
                with del_hol_col2:
                    st.write("")
                    st.write("")
                    if st.button("🗑️ ลบวันหยุดนี้", key="btn_del_hol", type="primary"):
                        idx_del = df_hol_del[df_hol_del["label"] == del_hol_label].index.tolist()
                        if idx_del:
                            df_hol_after = df_hol_custom_fresh.drop(index=idx_del).reset_index(drop=True)
                            if write_excel_to_drive(FILE_HOLIDAYS, df_hol_after):
                                log_activity("ลบวันหยุดพิเศษ", del_hol_label, "Admin")
                                st.toast("✅ ลบวันหยุดสำเร็จ", icon="🗑️")
                                st.cache_data.clear()
                                time.sleep(0.5)
                                st.rerun()

            st.divider()

            # ─── Export ─────────────────────────────────────────
            st.markdown("#### 📥 Export ปฏิทินวันหยุด")
            exp_hol_col1, exp_hol_col2 = st.columns(2)
            with exp_hol_col1:
                exp_hol_yr_be = st.selectbox(
                    "เลือกปี (พ.ศ.) ที่ต้องการ Export",
                    [y + 543 for y in hol_yr_opts],
                    key="exp_hol_year",
                )
                exp_hol_yr_ad = exp_hol_yr_be - 543
            with exp_hol_col2:
                st.write("")
                st.write("")
                if st.button("📥 Export Excel", key="btn_exp_hol", use_container_width=True):
                    df_exp = load_holidays_all(exp_hol_yr_ad)
                    if df_exp.empty:
                        st.warning("ไม่มีข้อมูล")
                    else:
                        df_exp_out = df_exp.copy()
                        df_exp_out["วันที่"] = pd.to_datetime(df_exp_out["วันที่"], errors="coerce")
                        df_exp_out["วัน"] = df_exp_out["วันที่"].dt.strftime("%A").map({
                            "Monday":"จันทร์","Tuesday":"อังคาร","Wednesday":"พุธ",
                            "Thursday":"พฤหัสบดี","Friday":"ศุกร์","Saturday":"เสาร์","Sunday":"อาทิตย์",
                        })
                        df_exp_out["กระทบวันทำการ"] = df_exp_out["วันที่"].dt.weekday.apply(
                            lambda w: "ใช่" if w < 5 else "ไม่"
                        )
                        df_exp_out["วันที่"] = df_exp_out["วันที่"].dt.strftime("%d/%m/%Y")
                        buf_hol = io.BytesIO()
                        with pd.ExcelWriter(buf_hol, engine="xlsxwriter") as w:
                            df_exp_out[["วันที่","วัน","ชื่อวันหยุด","ประเภท","กระทบวันทำการ","หมายเหตุ"]].to_excel(
                                w, index=False, sheet_name=f"วันหยุด_{exp_hol_yr_be}"
                            )
                        buf_hol.seek(0)
                        st.download_button(
                            f"⬇️ ดาวน์โหลดปฏิทินวันหยุด {exp_hol_yr_be}",
                            buf_hol,
                            f"Holidays_{exp_hol_yr_be}.xlsx",
                            mime=EXCEL_MIME,
                            use_container_width=True,
                        )

    elif password:
        st.error("❌ รหัสผ่านไม่ถูกต้อง")
        st.info("💡 เปลี่ยนรหัสผ่านได้ที่ secrets.toml → admin_password")
