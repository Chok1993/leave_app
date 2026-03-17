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

FILE_ATTEND    = "attendance_report.xlsx"
FILE_LEAVE     = "leave_report.xlsx"
FILE_TRAVEL    = "travel_report.xlsx"
FILE_STAFF     = "staff_master.xlsx"
FILE_NOTIFY    = "activity_log.xlsx"
FILE_HOLIDAYS  = "special_holidays.xlsx"
FILE_MANUAL_SCAN = "manual_scan.xlsx"
FOLDER_ID   = "1YFJZvs59ahRHmlRrKcQwepWJz6A-4B7d"
ATTACHMENT_FOLDER_NAME = "Attachments_Leave_App"
BACKUP_FOLDER_NAME     = "Backup"   # 📁 Backup/ — โฟลเดอร์หลักสำหรับ backup

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
    """ค้นหา File ID ใน parent_id ที่กำหนด — ถ้ามี duplicate เก็บล่าสุด ลบที่เหลือ"""
    try:
        res = service.files().list(
            q=f"name='{filename}' and '{parent_id}' in parents and trashed=false",
            fields="files(id, modifiedTime)",
            orderBy="modifiedTime desc",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        files = res.get("files", [])
        if not files:
            return None
        keep_id = files[0]["id"]
        for dup in files[1:]:
            try:
                service.files().delete(fileId=dup["id"], supportsAllDrives=True).execute()
                logger.info(f"Deleted duplicate '{filename}' id={dup['id']}")
            except Exception:
                pass
        return keep_id
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
    """อ่านไฟล์ Excel จาก Drive พร้อม retry"""
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
            logger.warning(f"Read attempt {attempt + 1} failed for {filename}: {e}")
            if attempt == 2:
                st.error(f"อ่านไฟล์ {filename} ไม่สำเร็จ")
                return pd.DataFrame()
            time.sleep(2 ** attempt)
    return pd.DataFrame()


def read_excel_with_id(filename: str) -> Tuple[pd.DataFrame, Optional[str]]:
    """
    อ่านไฟล์ Excel พร้อมคืน (DataFrame, file_id) ไปพร้อมกัน
    ใช้คู่กับ write_excel_to_drive(known_file_id=fid) เพื่อป้องกัน
    race condition ที่ทำให้สร้างไฟล์ใหม่แทนการ update ของเดิม
    """
    for attempt in range(3):
        try:
            fid = get_file_id(filename)
            if not fid:
                return pd.DataFrame(), None
            req = service.files().get_media(fileId=fid, supportsAllDrives=True)
            fh = io.BytesIO()
            dl = MediaIoBaseDownload(fh, req)
            done = False
            while not done:
                _, done = dl.next_chunk()
            fh.seek(0)
            return pd.read_excel(fh, engine="openpyxl"), fid
        except Exception as e:
            logger.warning(f"read_excel_with_id attempt {attempt+1} failed for {filename}: {e}")
            if attempt == 2:
                return pd.DataFrame(), None
            time.sleep(2 ** attempt)
    return pd.DataFrame(), None


def _read_file_by_id(file_id: str) -> pd.DataFrame:
    """อ่านไฟล์ Excel จาก Drive โดยตรงจาก file_id"""
    try:
        req = service.files().get_media(fileId=file_id, supportsAllDrives=True)
        fh  = io.BytesIO()
        dl  = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = dl.next_chunk()
        fh.seek(0)
        return pd.read_excel(fh, engine="openpyxl")
    except Exception as e:
        logger.warning(f"_read_file_by_id({file_id}): {e}")
        return pd.DataFrame()


def read_excel_with_backup(filename: str,
                            dedup_cols: Optional[List[str]] = None) -> Tuple[pd.DataFrame, Optional[str]]:
    """
    อ่านข้อมูลจาก **ทั้ง main file และ Backup** แล้วรวมเป็น DataFrame เดียว
    ป้องกันข้อมูลสูญหายในกรณีที่ระบบเขียนข้อมูลลง Backup แทน main file

    โครงสร้าง Backup:
        Leave_App_Data/Backup/BAK_{filename}/BAK_{filename}.xlsx

    dedup_cols: คอลัมน์ที่ใช้ตรวจสอบ duplicate (None = ไม่ dedup)
    คืน: (DataFrame รวม, file_id ของ main file)
    """
    frames: List[pd.DataFrame] = []

    # ── 1. อ่าน main file ────────────────────────────────────────────
    df_main, main_fid = read_excel_with_id(filename)
    if not df_main.empty:
        df_main["_src"] = "main"
        frames.append(df_main)
        logger.info(f"read_excel_with_backup: main '{filename}' → {len(df_main)} rows")

    # ── 2. อ่าน Backup/BAK_{filename}/BAK_{filename}.xlsx ────────────
    bak_name = f"BAK_{filename}"
    try:
        backup_root = get_or_create_folder(BACKUP_FOLDER_NAME, FOLDER_ID)
        if backup_root:
            bak_subfolder = get_or_create_folder(bak_name, backup_root)
            if bak_subfolder:
                bak_fid = get_file_id(bak_name, bak_subfolder)
                if bak_fid:
                    df_bak = _read_file_by_id(bak_fid)
                    if not df_bak.empty:
                        df_bak["_src"] = "backup"
                        frames.append(df_bak)
                        logger.info(f"read_excel_with_backup: backup '{bak_name}' → {len(df_bak)} rows")
    except Exception as e:
        logger.warning(f"read_excel_with_backup: ไม่สามารถอ่าน backup ของ '{filename}': {e}")

    if not frames:
        return pd.DataFrame(), main_fid

    # ── 3. รวมและ dedup ──────────────────────────────────────────────
    df_all = pd.concat(frames, ignore_index=True)

    if dedup_cols:
        # เอา main ก่อน backup (sort by _src)
        df_all["_src_order"] = df_all["_src"].map({"main": 0, "backup": 1})
        df_all = (
            df_all
            .sort_values("_src_order")
            .drop_duplicates(subset=dedup_cols, keep="first")
            .drop(columns=["_src_order"], errors="ignore")
        )

    df_all = df_all.drop(columns=["_src"], errors="ignore").reset_index(drop=True)

    main_rows = len(df_main) if not df_main.empty else 0
    bak_rows  = len(df_all) - main_rows
    if bak_rows > 0:
        logger.info(f"read_excel_with_backup: ดึงข้อมูลเพิ่มจาก Backup {bak_rows} แถว → รวม {len(df_all)} แถว")

    return df_all, main_fid


def _normalize_name(val) -> str:
    """แปลงค่าชื่อให้เป็น string สะอาด — คืน '' ถ้าว่างหรือ nan"""
    import re as _re
    if val is None: return ""
    s = str(val).strip()
    if s.lower() in ("nan", "none", ""): return ""
    # collapse internal whitespace
    return _re.sub(r"\s+", " ", s)


def _normalize_date(val) -> Optional[dt.date]:
    """
    แปลงวันที่หลายรูปแบบเป็น dt.date
    รองรับ: datetime, Timestamp, string (d/m/Y, Y-m-d, d/m/Y H:M:S)
    """
    if val is None: return None
    if isinstance(val, dt.datetime): return val.date()
    if isinstance(val, dt.date): return val
    try:
        ts = pd.to_datetime(val, dayfirst=True, errors="coerce")
        if pd.isna(ts): return None
        return ts.date()
    except Exception:
        return None


def _normalize_time_value(val) -> str:
    """
    แปลงค่าเวลาจาก Excel ทุกรูปแบบให้เป็น string "HH:MM"
    คืน "" ถ้าแปลงไม่ได้หรือเป็น null
    """
    import re as _re, math
    if val is None: return ""
    # float NaN
    if isinstance(val, float):
        if math.isnan(val): return ""
        # Excel stores time as fraction of day (0.375 = 09:00)
        total_sec = int(round(val * 86400))
        h = (total_sec // 3600) % 24
        m = (total_sec % 3600) // 60
        return f"{h:02d}:{m:02d}"
    # timedelta / pd.Timedelta
    if isinstance(val, (pd.Timedelta, dt.timedelta)):
        total_sec = int(val.total_seconds())
        if total_sec < 0: return ""
        h = (total_sec // 3600) % 24
        m = (total_sec % 3600) // 60
        return f"{h:02d}:{m:02d}"
    # datetime → extract time part
    if isinstance(val, dt.datetime): return val.strftime("%H:%M")
    if isinstance(val, dt.time): return val.strftime("%H:%M")
    # string
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", "nat", ""): return ""
    # "0 days HH:MM:SS" หรือ "H:MM:SS" หรือ "HH:MM" หรือ "H:MM AM/PM"
    m_re = _re.search(r"(\d+):(\d{2})(?::(\d{2}))?(?:\s*(AM|PM))?", s, _re.IGNORECASE)
    if m_re:
        h  = int(m_re.group(1))
        mn = int(m_re.group(2))
        meridiem = (m_re.group(4) or "").upper()
        # days prefix
        d_m = _re.search(r"(\d+)\s+day", s, _re.IGNORECASE)
        if d_m: h += int(d_m.group(1)) * 24
        # AM/PM
        if meridiem == "PM" and h < 12: h += 12
        elif meridiem == "AM" and h == 12: h = 0
        return f"{h % 24:02d}:{mn:02d}"
    return ""


@st.cache_data(ttl=300)
def read_attendance_report() -> pd.DataFrame:
    """
    อ่านและ normalize attendance_report.xlsx อย่างละเอียด
    
    โครงสร้างไฟล์จริง:
      Col A = ชื่อพนักงาน  (ว่างสำหรับแถวที่ Admin คีย์แทน)
      Col B = วันที่        (mixed format: d/m/Y, Y-m-d, datetime)
      Col C = เวลาเข้า      (Timedelta / string / float)
      Col D = เวลาออก       (เหมือน เวลาเข้า)
      Col E = สาย
      Col F = ออกก่อน
      Col G = หมายเหตุ
      Col H = ชื่อ-สกุล    (มีค่าเฉพาะแถวที่ col A ว่าง)
    
    กฎการ resolve ชื่อ:
      - ถ้า col A มีค่า → ใช้ col A เป็นชื่อ-สกุล
      - ถ้า col A ว่าง แต่ col H มีค่า → ใช้ col H (Admin คีย์แทน)
      - ถ้าทั้งคู่ว่าง → ข้ามแถวนี้ไป
    """
    # ── อ่านไฟล์ดิบ ────────────────────────────────────────────
    fid = get_file_id(FILE_ATTEND)
    if not fid:
        logger.warning("attendance_report.xlsx: ไม่พบในไฟล์ Drive")
        return pd.DataFrame()

    try:
        req = service.files().get_media(fileId=fid, supportsAllDrives=True)
        fh  = io.BytesIO()
        dl  = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = dl.next_chunk()
        fh.seek(0)
        # dtype=str เพื่อป้องกัน pandas auto-cast ผิดพลาด ยกเว้น column ที่ต้องเป็น numeric
        df_raw = pd.read_excel(
            fh,
            engine="openpyxl",
            header=0,
            dtype=str,           # อ่านทุก column เป็น string ก่อน → parse เองในขั้นถัดไป
        )
    except Exception as e:
        logger.error(f"read_attendance_report: {e}")
        return pd.DataFrame()

    if df_raw.empty:
        return pd.DataFrame()

    # ── normalize column names ─────────────────────────────────
    df_raw.columns = [str(c).strip() for c in df_raw.columns]

    # map column ที่รู้จัก (ชื่ออาจต่างกันได้)
    COL_NAME_A = next((c for c in df_raw.columns if c in ("ชื่อพนักงาน", "ชื่อ")), None)
    COL_DATE   = next((c for c in df_raw.columns if c in ("วันที่", "date", "Date")), None)
    COL_IN     = next((c for c in df_raw.columns if c in ("เวลาเข้า", "เข้า", "check_in")), None)
    COL_OUT    = next((c for c in df_raw.columns if c in ("เวลาออก", "ออก", "check_out")), None)
    COL_NOTE   = next((c for c in df_raw.columns if c in ("หมายเหตุ", "note", "Note")), None)
    COL_NAME_H = next((c for c in df_raw.columns if c in ("ชื่อ-สกุล",)), None)

    if COL_DATE is None:
        logger.error("read_attendance_report: ไม่พบคอลัมน์วันที่")
        return pd.DataFrame()

    rows_out: list = []

    for _, row in df_raw.iterrows():
        # ── resolve ชื่อ ──────────────────────────────────────
        name_a = _normalize_name(row.get(COL_NAME_A, "")) if COL_NAME_A else ""
        name_h = _normalize_name(row.get(COL_NAME_H, "")) if COL_NAME_H else ""

        if name_a:
            name = name_a          # col A ก่อนเสมอ
        elif name_h:
            name = name_h          # col H fallback (Admin คีย์แทน)
        else:
            continue               # ไม่มีชื่อ → ข้าม

        # ── resolve วันที่ ────────────────────────────────────
        date_val = _normalize_date(row.get(COL_DATE, ""))
        if date_val is None:
            continue               # ไม่มีวันที่ valid → ข้าม

        # ── resolve เวลา ──────────────────────────────────────
        t_in_str  = _normalize_time_value(row.get(COL_IN,  "")) if COL_IN  else ""
        t_out_str = _normalize_time_value(row.get(COL_OUT, "")) if COL_OUT else ""

        # ── หมายเหตุ ──────────────────────────────────────────
        note = str(row.get(COL_NOTE, "") or "").strip()

        rows_out.append({
            "ชื่อ-สกุล": name,
            "วันที่":     pd.Timestamp(date_val),
            "เวลาเข้า":   t_in_str,
            "เวลาออก":    t_out_str,
            "หมายเหตุ":   note,
        })

    if not rows_out:
        return pd.DataFrame(columns=["ชื่อ-สกุล", "วันที่", "เวลาเข้า", "เวลาออก", "หมายเหตุ"])

    df_out = pd.DataFrame(rows_out)
    df_out["วันที่"] = pd.to_datetime(df_out["วันที่"], errors="coerce").dt.normalize()
    df_out["เดือน"] = df_out["วันที่"].dt.strftime("%Y-%m")
    df_out = df_out.dropna(subset=["วันที่"])
    df_out = df_out[df_out["ชื่อ-สกุล"] != ""].reset_index(drop=True)
    logger.info(f"read_attendance_report: {len(df_out)} rows loaded")
    return df_out

@st.cache_data(ttl=300)
def list_all_files_in_folder(parent_id: str = FOLDER_ID) -> List[dict]:
    """ดึงรายชื่อไฟล์ทั้งหมดใน Drive folder (เฉพาะ .xlsx)"""
    try:
        res = service.files().list(
            q=(
                f"'{parent_id}' in parents and trashed=false "
                f"and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'"
            ),
            fields="files(id, name, modifiedTime)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            orderBy="modifiedTime desc",
        ).execute()
        return res.get("files", [])
    except Exception as e:
        logger.error(f"list_all_files_in_folder: {e}")
        return []

# columns ที่ต้องการจาก travel — ใช้เพื่อ normalize
TRAVEL_REQUIRED_COLS = ["ชื่อ-สกุล", "วันที่เริ่ม", "วันที่สิ้นสุด", "เรื่อง/กิจกรรม"]

# ชื่อไฟล์ที่รู้แน่ว่าไม่ใช่ข้อมูลไปราชการ — ข้ามไปเลย
_NON_TRAVEL_FILES = {
    FILE_ATTEND, FILE_LEAVE, FILE_STAFF, FILE_NOTIFY,
    FILE_HOLIDAYS, FILE_MANUAL_SCAN,
    # ชื่อ pattern ที่เป็น backup
}

@st.cache_data(ttl=300)
def load_all_travel() -> pd.DataFrame:
    """
    โหลดข้อมูลไปราชการจาก:
      1. ทุกไฟล์ .xlsx ใน root folder (เดิม)
      2. Backup/BAK_travel_report/BAK_travel_report.xlsx (ใหม่)
         — รวมข้อมูลที่อาจถูกเขียนไว้ใน Backup โดยผิดพลาด
    """
    frames: List[pd.DataFrame] = []
    files = list_all_files_in_folder()

    for f in files:
        fname = f.get("name", "")

        # ข้ามไฟล์ที่รู้ว่าไม่ใช่ travel
        if fname in _NON_TRAVEL_FILES:
            continue
        # ข้าม backup files ใน root (แต่จะดึงจาก Backup/ folder แยกต่างหากด้านล่าง)
        if fname.startswith("BAK_"):
            continue

        try:
            df_raw = read_excel_from_drive(fname)
            if df_raw.empty:
                continue

            # ตรวจว่ามีคอลัมน์ที่จำเป็นหรือไม่
            has_name  = "ชื่อ-สกุล" in df_raw.columns or "ชื่อพนักงาน" in df_raw.columns or "ชื่อ" in df_raw.columns
            has_start = "วันที่เริ่ม"   in df_raw.columns
            has_end   = "วันที่สิ้นสุด" in df_raw.columns

            if not (has_name and has_start and has_end):
                continue

            df_norm = df_raw.copy()

            # Normalize ชื่อ column
            for alt in ["ชื่อพนักงาน", "ชื่อ", "fullname"]:
                if alt in df_norm.columns and "ชื่อ-สกุล" not in df_norm.columns:
                    df_norm.rename(columns={alt: "ชื่อ-สกุล"}, inplace=True)

            # Normalize dates
            df_norm["วันที่เริ่ม"]   = pd.to_datetime(df_norm["วันที่เริ่ม"],   errors="coerce").dt.normalize()
            df_norm["วันที่สิ้นสุด"] = pd.to_datetime(df_norm["วันที่สิ้นสุด"], errors="coerce").dt.normalize()

            # ตรวจ activity_log
            if fname == FILE_NOTIFY or (
                "ประเภท" in df_norm.columns and "รายละเอียด" in df_norm.columns
            ):
                df_travel_from_log = _extract_travel_from_activity_log(df_norm)
                if not df_travel_from_log.empty:
                    df_travel_from_log["_source_file"] = fname
                    frames.append(df_travel_from_log)
                continue

            df_norm["_source_file"] = fname
            if "เรื่อง/กิจกรรม" not in df_norm.columns:
                df_norm["เรื่อง/กิจกรรม"] = fname.replace(".xlsx", "")

            df_norm = df_norm.dropna(subset=["ชื่อ-สกุล", "วันที่เริ่ม", "วันที่สิ้นสุด"])
            df_norm["ชื่อ-สกุล"] = df_norm["ชื่อ-สกุล"].astype(str).str.strip()
            df_norm = df_norm[df_norm["ชื่อ-สกุล"].str.lower() != "nan"]

            if not df_norm.empty:
                frames.append(df_norm[TRAVEL_REQUIRED_COLS + ["_source_file"]])

        except Exception as e:
            logger.warning(f"load_all_travel: skip {fname} — {e}")
            continue

    # ── อ่าน Backup/BAK_travel_report/BAK_travel_report.xlsx ─────────
    try:
        backup_root = get_or_create_folder(BACKUP_FOLDER_NAME, FOLDER_ID)
        if backup_root:
            bak_name      = f"BAK_{FILE_TRAVEL}"
            bak_subfolder = get_or_create_folder(bak_name, backup_root)
            if bak_subfolder:
                bak_fid = get_file_id(bak_name, bak_subfolder)
                if bak_fid:
                    df_bak = _read_file_by_id(bak_fid)
                    if not df_bak.empty:
                        # normalize เหมือน main
                        for alt in ["ชื่อพนักงาน", "ชื่อ"]:
                            if alt in df_bak.columns and "ชื่อ-สกุล" not in df_bak.columns:
                                df_bak.rename(columns={alt: "ชื่อ-สกุล"}, inplace=True)
                        df_bak["วันที่เริ่ม"]   = pd.to_datetime(df_bak.get("วันที่เริ่ม"),   errors="coerce").dt.normalize()
                        df_bak["วันที่สิ้นสุด"] = pd.to_datetime(df_bak.get("วันที่สิ้นสุด"), errors="coerce").dt.normalize()
                        if "เรื่อง/กิจกรรม" not in df_bak.columns:
                            df_bak["เรื่อง/กิจกรรม"] = "ไปราชการ"
                        df_bak = df_bak.dropna(subset=["ชื่อ-สกุล", "วันที่เริ่ม", "วันที่สิ้นสุด"])
                        df_bak["ชื่อ-สกุล"] = df_bak["ชื่อ-สกุล"].astype(str).str.strip()
                        df_bak = df_bak[df_bak["ชื่อ-สกุล"].str.lower() != "nan"]
                        df_bak["_source_file"] = f"[Backup] {FILE_TRAVEL}"
                        valid_cols = [c for c in TRAVEL_REQUIRED_COLS + ["_source_file"] if c in df_bak.columns]
                        if not df_bak.empty:
                            frames.append(df_bak[valid_cols])
                            logger.info(f"load_all_travel: backup '{bak_name}' → {len(df_bak)} rows")
    except Exception as e:
        logger.warning(f"load_all_travel: ไม่สามารถอ่าน Backup travel: {e}")

    if not frames:
        return pd.DataFrame(columns=TRAVEL_REQUIRED_COLS + ["_source_file"])

    df_all = pd.concat(frames, ignore_index=True)

    # dedup: main > backup > อื่นๆ
    def _rank_src(src: str) -> int:
        if src == FILE_TRAVEL:         return 0  # main file — highest priority
        if src.startswith("[Backup]"): return 1  # backup
        return 2                                  # other files

    df_all["_rank"] = df_all["_source_file"].apply(_rank_src)
    df_all = (
        df_all.sort_values(["ชื่อ-สกุล", "วันที่เริ่ม", "_rank"])
        .drop_duplicates(subset=["ชื่อ-สกุล", "วันที่เริ่ม", "วันที่สิ้นสุด"], keep="first")
        .drop(columns=["_rank"])
        .reset_index(drop=True)
    )
    return df_all


def _extract_travel_from_activity_log(df_log: pd.DataFrame) -> pd.DataFrame:
    """
    แปลงข้อมูลจาก activity_log.xlsx ที่มี schema:
      [Timestamp, ประเภท, รายละเอียด, ผู้เกี่ยวข้อง]
    เฉพาะแถว ประเภท == "ไปราชการ" → แปลงเป็น travel schema
    รูปแบบ รายละเอียด: "<project> @ <location>"
    รูปแบบ ผู้เกี่ยวข้อง: "นาย ก, นาย ข, และอีก N คน"

    หมายเหตุ: activity_log ไม่มี วันที่เริ่ม/สิ้นสุด → ใช้ Timestamp เป็น proxy (วันเดียว)
    """
    if df_log.empty or "ประเภท" not in df_log.columns:
        return pd.DataFrame()

    df_tr = df_log[df_log["ประเภท"] == "ไปราชการ"].copy()
    if df_tr.empty:
        return pd.DataFrame()

    rows = []
    for _, row in df_tr.iterrows():
        try:
            ts = pd.to_datetime(row.get("Timestamp"), errors="coerce")
            if pd.isna(ts):
                continue
            d = ts.normalize()
            detail   = str(row.get("รายละเอียด", ""))
            involved = str(row.get("ผู้เกี่ยวข้อง", ""))
            project  = detail.split("@")[0].strip() if "@" in detail else detail

            # แยกชื่อ (คั่นด้วย ",")
            names = [n.strip() for n in involved.split(",") if n.strip() and "และอีก" not in n]
            for name in names:
                rows.append({
                    "ชื่อ-สกุล":       name,
                    "วันที่เริ่ม":      d,
                    "วันที่สิ้นสุด":    d,   # log ไม่มีวันสิ้นสุด → ใช้วันเดียวกัน
                    "เรื่อง/กิจกรรม":  project,
                })
        except Exception:
            continue

    return pd.DataFrame(rows) if rows else pd.DataFrame()
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

def write_excel_to_drive(filename: str, df: pd.DataFrame,
                          known_file_id: Optional[str] = None) -> bool:
    """
    เขียน DataFrame กลับไปที่ไฟล์เดิมใน Drive โดย update in-place
    ไม่สร้างไฟล์ใหม่ถ้าไฟล์มีอยู่แล้ว

    known_file_id: ถ้าส่งมาจะใช้ค่านี้โดยตรง ไม่ต้อง list ใหม่
                   (ป้องกัน race condition ระหว่าง read → write)
    """
    try:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            df.to_excel(w, index=False)
        buf.seek(0)
        media = MediaIoBaseUpload(buf, mimetype=EXCEL_MIME, resumable=False)

        # ── หา file ID ───────────────────────────────────────────────
        fid = known_file_id or get_file_id(filename)

        if fid:
            # UPDATE in-place — ห้าม create ใหม่
            service.files().update(
                fileId=fid,
                media_body=media,
                supportsAllDrives=True,
            ).execute()
            logger.info(f"write_excel_to_drive: updated '{filename}' id={fid}")
        else:
            # ไม่มีไฟล์นี้ใน Drive เลย → สร้างใหม่ครั้งเดียว
            new_file = service.files().create(
                body={"name": filename, "parents": [FOLDER_ID]},
                media_body=media,
                supportsAllDrives=True,
                fields="id",
            ).execute()
            logger.info(f"write_excel_to_drive: created '{filename}' id={new_file.get('id')}")

        st.cache_data.clear()
        return True

    except Exception as e:
        logger.error(f"write_excel_to_drive({filename}): {e}")
        st.error(f"บันทึกไฟล์ล้มเหลว: {e}")
        return False

def backup_excel(filename: str, df: pd.DataFrame) -> None:
    """
    สำรองไฟล์ก่อนแก้ไข โครงสร้าง:
      📁 Backup/
        └── 📁 BAK_{filename}/
              └── BAK_{filename}.xlsx  ← overwrite ทุกครั้ง (1 ไฟล์คงที่)
    """
    if df.empty:
        return
    try:
        fid = get_file_id(filename)
        if not fid:
            return

        bak_name     = f"BAK_{filename}"

        # 1. หรือสร้าง Backup/ folder
        backup_root  = get_or_create_folder(BACKUP_FOLDER_NAME, FOLDER_ID)
        if not backup_root:
            logger.warning("backup_excel: ไม่สามารถสร้าง Backup/ folder ได้")
            return

        # 2. หรือสร้าง Backup/BAK_{filename}/ subfolder
        bak_subfolder = get_or_create_folder(bak_name, backup_root)
        if not bak_subfolder:
            logger.warning(f"backup_excel: ไม่สามารถสร้าง {bak_name}/ subfolder ได้")
            return

        # 3. ลบ BAK เดิมใน subfolder (ถ้ามี) แล้ว copy ใหม่ทับ
        existing_bak_id = get_file_id(bak_name, bak_subfolder)
        if existing_bak_id:
            try:
                service.files().delete(
                    fileId=existing_bak_id, supportsAllDrives=True
                ).execute()
            except Exception:
                pass

        service.files().copy(
            fileId=fid,
            body={"name": bak_name, "parents": [bak_subfolder]},
            supportsAllDrives=True,
        ).execute()
        logger.info(f"Backup saved: {BACKUP_FOLDER_NAME}/{bak_name}/{bak_name}")

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
    """
    Strip + normalize whitespace จาก column ชื่อ
    FIX 1: ป้องกัน duplicate columns
    FIX 2: normalize internal whitespace (เช่น double space)
    """
    if df.empty or col not in df.columns:
        return df
    if df.columns.duplicated().any():
        df = df.loc[:, ~df.columns.duplicated()].copy()
    series = df[col]
    if isinstance(series, pd.DataFrame):
        series = series.iloc[:, 0]
    # strip + collapse internal whitespace
    df[col] = (
        series.astype(str)
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
    )
    return df

def preprocess_dataframes(df_leave, df_travel, df_att):
    """
    FIX: for loop เดิมทำ `df = clean_names(df, ...)` แต่ไม่ได้ assign กลับ
    ทำให้ df_leave / df_travel / df_att ที่ return ออกไปยังไม่ถูก clean
    แก้โดย assign ตรงๆ ทีละตัวแทน
    """
    # Rename columns ใน df_att — ทำเฉพาะกรณีที่ df_att ยังไม่ผ่าน read_attendance_report
    # (เช่น ถูกส่งมาจากที่อื่น) เพื่อความ backward compat
    if not df_att.empty:
        for old, new in COLUMN_MAPPING.items():
            if old in df_att.columns:
                if new in df_att.columns:
                    df_att = df_att.drop(columns=[new])
                df_att = df_att.rename(columns={old: new})
        if df_att.columns.duplicated().any():
            df_att = df_att.loc[:, ~df_att.columns.duplicated()].copy()
    # Normalize dates
    for col in ["วันที่เริ่ม", "วันที่สิ้นสุด"]:
        df_leave  = normalize_date_col(df_leave,  col)
        df_travel = normalize_date_col(df_travel, col)
    df_att = normalize_date_col(df_att, "วันที่")
    # Clean names — assign ผลกลับทีละตัว (ไม่ใช้ for loop)
    df_leave  = clean_names(df_leave,  "ชื่อ-สกุล")
    df_travel = clean_names(df_travel, "ชื่อ-สกุล")
    df_att    = clean_names(df_att,    "ชื่อ-สกุล")
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

def parse_time(val) -> Optional[dt.time]:
    """
    แปลง value หลายรูปแบบเป็น time object — คืน None ถ้าแปลงไม่ได้
    รองรับ: dt.time, dt.datetime, pd.Timedelta, str ("HH:MM", "HH:MM:SS"),
            float (Excel serial fraction), "6:33:00 PM" (12h format)
    """
    if val is None or val == "":
        return None
    if isinstance(val, float):
        if np.isnan(val):
            return None
        # Excel เก็บเวลาเป็น fraction ของวัน (0.375 = 09:00)
        try:
            total_sec = int(round(val * 86400))
            return dt.time(total_sec // 3600, (total_sec % 3600) // 60, total_sec % 60)
        except Exception:
            return None
    if isinstance(val, dt.time):
        return val
    if isinstance(val, dt.datetime):
        return val.time()
    # Timedelta (pandas อ่านเวลาจาก Excel บางครั้งได้เป็น timedelta)
    if isinstance(val, (pd.Timedelta, dt.timedelta)):
        try:
            total_sec = int(val.total_seconds())
            if total_sec < 0:
                return None
            return dt.time(total_sec // 3600 % 24, (total_sec % 3600) // 60, total_sec % 60)
        except Exception:
            return None
    # String formats
    s = str(val).strip()
    if not s or s.lower() in ("nat", "none", "nan", ""):
        return None
    # "0 days HH:MM:SS" pattern จาก pd.Timedelta ที่ถูก str() แล้ว
    # หรือ "HH:MM:SS AM/PM" format
    import re as _re
    m = _re.search(r"(\d+):(\d{2}):?(\d{2})?(?:\s*(AM|PM))?", s, _re.IGNORECASE)
    if m:
        h  = int(m.group(1))
        mn = int(m.group(2))
        sc = int(m.group(3)) if m.group(3) else 0
        meridiem = (m.group(4) or "").upper()
        # handle "X days" prefix
        d_match = _re.search(r"(\d+)\s+day", s, _re.IGNORECASE)
        if d_match:
            h += int(d_match.group(1)) * 24
        # handle AM/PM
        if meridiem == "PM" and h < 12:
            h += 12
        elif meridiem == "AM" and h == 12:
            h = 0
        try:
            return dt.time(h % 24, mn, sc)
        except Exception:
            pass
    try:
        return pd.to_datetime(s).time()
    except Exception:
        pass
    return None

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

def load_holidays_with_id() -> Tuple[pd.DataFrame, Optional[str]]:
    """โหลดวันหยุด Admin พร้อม file_id — ใช้คู่กับ write_excel_to_drive(known_file_id=)"""
    df, fid = read_excel_with_backup(FILE_HOLIDAYS, dedup_cols=["วันที่","ชื่อวันหยุด"])
    if not df.empty:
        df["วันที่"] = pd.to_datetime(df["วันที่"], errors="coerce")
        df = df.dropna(subset=["วันที่"])
        for col in HOLIDAY_COLS:
            if col not in df.columns:
                df[col] = ""
    return df, fid

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
# [MS] Manual Scan Helpers
# ===========================
MANUAL_SCAN_COLS = ["ชื่อ-สกุล", "วันที่", "เวลาเข้า", "เวลาออก", "หมายเหตุ"]

def _parse_manual_scan_detail(detail: str, person: str) -> Optional[dict]:
    """
    แปลงบรรทัด รายละเอียด ของ activity_log ประเภท "คีย์สแกนนิ้ว"
    รูปแบบ: "Admin คีย์ YYYY-MM-DD เข้า HH:MM ออก HH:MM"
    คืน dict หรือ None ถ้าแปลงไม่ได้
    """
    import re as _re
    # ลอง match pattern หลัก
    m = _re.search(
        r"(\d{4}-\d{2}-\d{2})\s+เข้า\s+(\d{1,2}:\d{2})\s+ออก\s+(\d{1,2}:\d{2})",
        str(detail),
    )
    if not m:
        # ลอง pattern วันที่แบบ DD/MM/YYYY
        m2 = _re.search(
            r"(\d{1,2}/\d{1,2}/\d{4})\s+เข้า\s+(\d{1,2}:\d{2})\s+ออก\s+(\d{1,2}:\d{2})",
            str(detail),
        )
        if not m2:
            return None
        date_str, t_in, t_out = m2.group(1), m2.group(2), m2.group(3)
        try:
            d = pd.to_datetime(date_str, dayfirst=True).normalize()
        except Exception:
            return None
    else:
        date_str, t_in, t_out = m.group(1), m.group(2), m.group(3)
        try:
            d = pd.to_datetime(date_str).normalize()
        except Exception:
            return None

    return {
        "ชื่อ-สกุล": person.strip(),
        "วันที่":     d,
        "เวลาเข้า":   t_in,
        "เวลาออก":    t_out,
        "หมายเหตุ":   f"Activity Log — {detail[:60]}",
    }

def _parse_delete_scan_detail(detail: str, person: str) -> Optional[tuple]:
    """
    แปลงบรรทัด รายละเอียด ของ activity_log ประเภท "ลบสแกนนิ้ว"
    รูปแบบ: "Admin ลบรายการ DD/MM/YYYY — เข้า HH:MM ออก HH:MM"
    คืน (person, date) หรือ None
    """
    import re as _re
    m = _re.search(r"(\d{1,2}/\d{1,2}/\d{4})", str(detail))
    if not m:
        m2 = _re.search(r"(\d{4}-\d{2}-\d{2})", str(detail))
        if not m2:
            return None
        date_str = m2.group(1)
        try:
            d = pd.to_datetime(date_str).normalize()
        except Exception:
            return None
    else:
        date_str = m.group(1)
        try:
            d = pd.to_datetime(date_str, dayfirst=True).normalize()
        except Exception:
            return None
    return (person.strip(), d)

@st.cache_data(ttl=300)
def load_manual_scans() -> pd.DataFrame:
    """
    โหลดข้อมูลลืมสแกนนิ้วจาก 2 แหล่ง แล้วรวมกัน:
      1. manual_scan.xlsx  — บันทึกตรงจากฟอร์มในแอป
      2. activity_log.xlsx — ประเภท 'คีย์สแกนนิ้ว' (หักรายการ 'ลบสแกนนิ้ว' ออก)
    """
    frames: List[pd.DataFrame] = []

    # ── 1. manual_scan.xlsx ───────────────────────────────────
    df_ms = read_excel_from_drive(FILE_MANUAL_SCAN)
    if not df_ms.empty:
        df_ms["วันที่"]    = pd.to_datetime(df_ms["วันที่"],    errors="coerce").dt.normalize()
        df_ms["ชื่อ-สกุล"] = df_ms["ชื่อ-สกุล"].astype(str).str.strip()
        for col in MANUAL_SCAN_COLS:
            if col not in df_ms.columns:
                df_ms[col] = ""
        frames.append(df_ms[MANUAL_SCAN_COLS])

    # ── 2. activity_log.xlsx ──────────────────────────────────
    df_log = read_excel_from_drive(FILE_NOTIFY)
    if not df_log.empty and "ประเภท" in df_log.columns:

        # รวบรวมรายการที่ถูกลบไว้ก่อน (key = "ชื่อ|วันที่")
        deleted_keys: set = set()
        df_del = df_log[df_log["ประเภท"].astype(str).str.strip() == "ลบสแกนนิ้ว"]
        for _, row in df_del.iterrows():
            result = _parse_delete_scan_detail(
                str(row.get("รายละเอียด", "")),
                str(row.get("ผู้เกี่ยวข้อง", "")),
            )
            if result:
                person, d = result
                deleted_keys.add(f"{person}|{d}")

        # parse รายการ คีย์สแกนนิ้ว
        df_key = df_log[df_log["ประเภท"].astype(str).str.strip() == "คีย์สแกนนิ้ว"]
        log_rows: List[dict] = []
        for _, row in df_key.iterrows():
            rec = _parse_manual_scan_detail(
                str(row.get("รายละเอียด", "")),
                str(row.get("ผู้เกี่ยวข้อง", "")),
            )
            if rec is None:
                continue
            key = f"{rec['ชื่อ-สกุล']}|{rec['วันที่']}"
            # ข้ามถ้าถูกลบแล้ว
            if key in deleted_keys:
                continue
            log_rows.append(rec)

        if log_rows:
            df_from_log = pd.DataFrame(log_rows)
            for col in MANUAL_SCAN_COLS:
                if col not in df_from_log.columns:
                    df_from_log[col] = ""
            frames.append(df_from_log[MANUAL_SCAN_COLS])

    if not frames:
        return pd.DataFrame(columns=MANUAL_SCAN_COLS)

    df_all = pd.concat(frames, ignore_index=True)
    df_all["วันที่"]    = pd.to_datetime(df_all["วันที่"], errors="coerce").dt.normalize()
    df_all["ชื่อ-สกุล"] = df_all["ชื่อ-สกุล"].astype(str).str.strip()
    df_all = df_all.dropna(subset=["วันที่"])
    df_all = df_all[df_all["ชื่อ-สกุล"].str.lower() != "nan"]

    # dedup: ถ้าคน+วันที่ซ้ำ เก็บอัน manual_scan.xlsx ก่อน (index ต่ำกว่า)
    df_all = df_all.drop_duplicates(subset=["ชื่อ-สกุล", "วันที่"], keep="first")
    df_all = df_all.sort_values(["ชื่อ-สกุล", "วันที่"]).reset_index(drop=True)
    return df_all

def merge_attendance_with_manual(df_att: pd.DataFrame, df_manual: pd.DataFrame) -> pd.DataFrame:
    """
    รวม attendance (post-preprocess) + manual scans
    เรียกหลัง preprocess_dataframes เสมอ เพื่อให้ df_att มี "ชื่อ-สกุล" ถูกต้อง
    """
    if df_manual.empty:
        return df_att

    if df_att.empty:
        df_manual_out = df_manual.copy()
        df_manual_out["_source"] = "manual"
        return df_manual_out

    df_att_work    = df_att.copy()
    df_manual_work = df_manual.copy()

    # หา name column ใน df_att (หลัง preprocess ควรเป็น "ชื่อ-สกุล" แล้ว)
    att_name_col = next(
        (c for c in ["ชื่อ-สกุล", "ชื่อพนักงาน", "ชื่อ"] if c in df_att_work.columns),
        None,
    )
    if att_name_col is None:
        # ไม่มี name column เลย — return att เดิม
        return df_att_work

    # ถ้าชื่อ column ไม่ใช่ "ชื่อ-สกุล" ให้ rename ก่อน
    if att_name_col != "ชื่อ-สกุล":
        df_att_work = df_att_work.rename(columns={att_name_col: "ชื่อ-สกุล"})

    df_att_work["วันที่"]    = pd.to_datetime(df_att_work["วันที่"],    errors="coerce").dt.normalize()
    df_manual_work["วันที่"] = pd.to_datetime(df_manual_work["วันที่"], errors="coerce").dt.normalize()

    # normalize ชื่อ (strip + collapse whitespace)
    df_att_work["ชื่อ-สกุล"]    = df_att_work["ชื่อ-สกุล"].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
    df_manual_work["ชื่อ-สกุล"] = df_manual_work["ชื่อ-สกุล"].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)

    # dedup key: "ชื่อ|YYYY-MM-DD"
    att_keys = set(
        df_att_work["ชื่อ-สกุล"].astype(str) + "|" +
        df_att_work["วันที่"].astype(str)
    )

    df_manual_new = df_manual_work[
        ~(df_manual_work["ชื่อ-สกุล"].astype(str) + "|" +
          df_manual_work["วันที่"].astype(str)).isin(att_keys)
    ].copy()

    if df_manual_new.empty:
        return df_att_work

    df_manual_new["_source"] = "manual"
    df_att_work["_source"]   = "scan"

    df_merged = pd.concat([df_att_work, df_manual_new], ignore_index=True)
    df_merged = df_merged.sort_values(["ชื่อ-สกุล", "วันที่"]).reset_index(drop=True)
    return df_merged
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
        df_log, _notify_fid = read_excel_with_backup(FILE_NOTIFY)
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
        write_excel_to_drive(FILE_NOTIFY, df_log, known_file_id=_notify_fid)
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
        df_leave  = read_excel_with_backup(FILE_LEAVE, dedup_cols=["ชื่อ-สกุล","วันที่เริ่ม","ประเภทการลา"])[0]
        df_travel = read_excel_with_backup(FILE_TRAVEL, dedup_cols=["ชื่อ-สกุล","วันที่เริ่ม","เรื่อง/กิจกรรม"])[0]
        df_att    = read_attendance_report()
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
        df_leave   = read_excel_with_backup(FILE_LEAVE, dedup_cols=["ชื่อ-สกุล","วันที่เริ่ม","ประเภทการลา"])[0]
        df_travel  = read_excel_with_backup(FILE_TRAVEL, dedup_cols=["ชื่อ-สกุล","วันที่เริ่ม","เรื่อง/กิจกรรม"])[0]
        df_att_raw = read_attendance_report()
        df_staff   = read_excel_from_drive(FILE_STAFF)
        df_leave, df_travel, df_att_raw = preprocess_dataframes(df_leave, df_travel, df_att_raw)
        df_manual  = load_manual_scans()
        df_att     = merge_attendance_with_manual(df_att_raw, df_manual)
        df_travel_all = load_all_travel()
        _, df_travel_pp, _ = preprocess_dataframes(df_leave, df_travel_all, pd.DataFrame())

    # ── CSS เฉพาะ Dashboard ──────────────────────────────────────────
    st.markdown("""
    <style>
    .kpi-card {
        background:linear-gradient(135deg,#1e293b 0%,#0f172a 100%);
        border-radius:14px; padding:20px 22px; border-left:4px solid;
        color:#f1f5f9; margin-bottom:4px;
    }
    .kpi-label { font-size:0.8rem; color:#94a3b8; font-weight:500; letter-spacing:.5px; }
    .kpi-value { font-size:2rem; font-weight:800; margin:4px 0; line-height:1; }
    .kpi-sub   { font-size:0.72rem; color:#64748b; margin-top:4px; }
    .insight-card {
        background:#1e293b; border-radius:10px; padding:14px 16px;
        border-left:3px solid #3b82f6; margin:6px 0; color:#e2e8f0;
    }
    .insight-num  { color:#60a5fa; font-weight:700; font-size:0.85rem; }
    .insight-title{ color:#cbd5e1; font-weight:600; font-size:0.85rem; }
    .insight-body { color:#94a3b8; font-size:0.82rem; margin-top:2px; }
    .section-card {
        background:#1e293b; border-radius:12px; padding:18px 20px;
        border:1px solid #334155; margin-bottom:12px;
    }
    .sec-title { color:#e2e8f0; font-weight:700; font-size:1rem; margin-bottom:12px; }
    </style>
    """, unsafe_allow_html=True)

    # ── คำนวณตัวเลขสำคัญ ──────────────────────────────────────────────
    LATE_CUT = dt.time(8, 31)

    def _att_status(row):
        if pd.to_datetime(row["วันที่"], errors="coerce").weekday() >= 5:
            return "วันหยุด"
        t_in  = parse_time(row.get("เวลาเข้า", ""))
        t_out = parse_time(row.get("เวลาออก",  ""))
        if not t_in and not t_out: return "ขาดงาน"
        if (t_in and not t_out) or (not t_in and t_out) or (t_in == t_out): return "ลืมสแกน"
        if t_in >= LATE_CUT: return "มาสาย"
        return "มาปกติ"

    if not df_att.empty:
        df_att["วันที่"] = pd.to_datetime(df_att["วันที่"], errors="coerce")
        df_att["เดือน"]  = df_att["วันที่"].dt.strftime("%Y-%m")
        df_att["สถานะสแกน"] = df_att.apply(_att_status, axis=1)
        df_work = df_att[~df_att["สถานะสแกน"].isin(["วันหยุด"])]
        total_work = len(df_work)
        n_ok     = len(df_work[df_work["สถานะสแกน"] == "มาปกติ"])
        n_late   = len(df_work[df_work["สถานะสแกน"] == "มาสาย"])
        n_absent = len(df_work[df_work["สถานะสแกน"] == "ขาดงาน"])
        n_forgot = len(df_work[df_work["สถานะสแกน"] == "ลืมสแกน"])
        pct_ok     = n_ok     / total_work * 100 if total_work else 0
        pct_late   = n_late   / total_work * 100 if total_work else 0
        pct_absent = n_absent / total_work * 100 if total_work else 0
    else:
        total_work = n_ok = n_late = n_absent = n_forgot = 0
        pct_ok = pct_late = pct_absent = 0

    # ── KPI CARDS ─────────────────────────────────────────────────────
    kc1, kc2, kc3, kc4 = st.columns(4)
    kpi_data = [
        (kc1, "🗓️ วันทำการรวม",    f"{total_work:,}", f"จากข้อมูลสแกน {len(df_att):,} รายการ", "#3b82f6"),
        (kc2, "✅ อัตรามาปกติ",    f"{pct_ok:.1f}%",   f"มาปกติ {n_ok:,} วัน",   "#22c55e"),
        (kc3, "⏰ อัตรามาสาย",     f"{pct_late:.1f}%", f"มาสาย {n_late:,} วัน",  "#f59e0b"),
        (kc4, "❌ อัตราขาดงาน",    f"{pct_absent:.1f}%",f"ขาด {n_absent:,} วัน",  "#ef4444"),
    ]
    for col, label, val, sub, border_col in kpi_data:
        with col:
            st.markdown(f"""
            <div class="kpi-card" style="border-color:{border_col}">
                <div class="kpi-label">{label}</div>
                <div class="kpi-value" style="color:{border_col}">{val}</div>
                <div class="kpi-sub">{sub}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    tab_summary, tab_trend, tab_charts, tab_insight, tab_export = st.tabs([
        "📋 ตารางสรุปรายบุคคล",
        "📈 แนวโน้มรายเดือน",
        "📊 กราฟวิเคราะห์",
        "🔍 7 ข้อวิเคราะห์",
        "📥 Export รายงาน",
    ])

    # ═══════════════════════════════════════════════════
    # TAB 1: ตารางสรุปรายบุคคล + Conditional Formatting
    # ═══════════════════════════════════════════════════
    with tab_summary:
        st.markdown('<div class="sec-title">📋 ตารางสรุปรายบุคคล — อัตรามาปกติ (%) พร้อม Conditional Formatting</div>', unsafe_allow_html=True)

        # filter
        sf1, sf2, sf3 = st.columns([1, 2, 2])
        with sf1:
            all_months_s = sorted(df_att["เดือน"].dropna().unique().tolist()) if not df_att.empty else []
            sel_s_month  = st.selectbox("เดือน", ["ทั้งหมด"] + all_months_s, key="s_month")
        with sf2:
            sel_s_grp = st.selectbox("กลุ่มงาน", ["ทุกกลุ่ม"] + STAFF_GROUPS, key="s_grp")
        with sf3:
            search_s = st.text_input("🔍 ค้นหาชื่อ", placeholder="...", key="s_search")

        df_s = df_work.copy() if not df_att.empty else pd.DataFrame()
        if not df_s.empty:
            if sel_s_month != "ทั้งหมด":
                df_s = df_s[df_s["เดือน"] == sel_s_month]
            if sel_s_grp != "ทุกกลุ่ม" and not df_staff.empty and "กลุ่มงาน" in df_staff.columns:
                grp_set = set(df_staff[df_staff["กลุ่มงาน"] == sel_s_grp]["ชื่อ-สกุล"].astype(str).str.strip())
                df_s = df_s[df_s["ชื่อ-สกุล"].isin(grp_set)]
            if search_s.strip():
                df_s = df_s[df_s["ชื่อ-สกุล"].str.contains(search_s.strip(), na=False)]

        if df_s.empty:
            st.info("ไม่มีข้อมูล")
        else:
            summary_rows = []
            for name, grp_df in df_s.groupby("ชื่อ-สกุล"):
                dept = ""
                if not df_staff.empty and "กลุ่มงาน" in df_staff.columns:
                    m = df_staff[df_staff["ชื่อ-สกุล"] == name]
                    if not m.empty: dept = str(m.iloc[0].get("กลุ่มงาน",""))
                total = len(grp_df)
                ok     = len(grp_df[grp_df["สถานะสแกน"] == "มาปกติ"])
                late   = len(grp_df[grp_df["สถานะสแกน"] == "มาสาย"])
                absent = len(grp_df[grp_df["สถานะสแกน"] == "ขาดงาน"])
                forgot = len(grp_df[grp_df["สถานะสแกน"] == "ลืมสแกน"])
                pct    = round(ok / total * 100, 1) if total > 0 else 0
                summary_rows.append({"ชื่อ-สกุล": name, "กลุ่มงาน": dept,
                                     "มาปกติ": ok, "มาสาย": late, "ขาดงาน": absent,
                                     "ลืมสแกน": forgot, "วันรวม": total,
                                     "% มาปกติ": pct})
            df_sum = pd.DataFrame(summary_rows).sort_values("% มาปกติ", ascending=False).reset_index(drop=True)

            # Conditional formatting function
            def color_pct(val):
                if not isinstance(val, (int, float)): return ""
                if val >= 80: return "background-color:#166534;color:#dcfce7;font-weight:700"
                if val >= 60: return "background-color:#854d0e;color:#fef9c3;font-weight:700"
                return "background-color:#991b1b;color:#fee2e2;font-weight:700"

            def color_absent(val):
                if not isinstance(val, (int, float)) or val == 0: return ""
                if val >= 5:  return "background-color:#7f1d1d;color:#fca5a5"
                if val >= 2:  return "background-color:#b91c1c;color:#fecaca"
                return "background-color:#dc2626;color:#fee2e2"

            styled = (
                df_sum.style
                .applymap(color_pct,    subset=["% มาปกติ"])
                .applymap(color_absent, subset=["ขาดงาน"])
                .format({"% มาปกติ": "{:.1f}%"})
                .set_properties(**{"font-size":"12px"})
            )
            st.dataframe(styled, use_container_width=True, height=450)
            st.caption(f"แสดง {len(df_sum)} คน  |  🟢 ≥80%  🟡 60-79%  🔴 <60%")

            # Export
            buf_s = io.BytesIO()
            with pd.ExcelWriter(buf_s, engine="xlsxwriter") as w:
                df_sum.to_excel(w, index=False, sheet_name="รายบุคคล")
            st.download_button("📥 Export ตาราง", buf_s.getvalue(),
                               "PersonSummary.xlsx", mime=EXCEL_MIME)

    # ═══════════════════════════════════════════════════
    # TAB 2: แนวโน้มรายเดือน
    # ═══════════════════════════════════════════════════
    with tab_trend:
        st.markdown('<div class="sec-title">📈 แนวโน้มรายเดือน</div>', unsafe_allow_html=True)

        if df_att.empty:
            st.info("ไม่มีข้อมูล")
        else:
            df_monthly = (
                df_work.groupby("เดือน")["สถานะสแกน"]
                .value_counts().unstack(fill_value=0).reset_index()
            )
            for col in ["มาปกติ","มาสาย","ขาดงาน","ลืมสแกน"]:
                if col not in df_monthly.columns: df_monthly[col] = 0
            df_monthly["วันรวม"] = df_monthly[["มาปกติ","มาสาย","ขาดงาน","ลืมสแกน"]].sum(axis=1)
            df_monthly["% มาปกติ"] = (df_monthly["มาปกติ"] / df_monthly["วันรวม"].replace(0,1) * 100).round(1)

            # mini KPI row per month
            for _, row_m in df_monthly.sort_values("เดือน").iterrows():
                p = row_m.get("% มาปกติ", 0)
                col_p = "#22c55e" if p >= 80 else ("#f59e0b" if p >= 60 else "#ef4444")
                st.markdown(f"""
                <div style="display:flex;align-items:center;gap:16px;background:#1e293b;
                     border-radius:8px;padding:10px 16px;margin:4px 0;border:1px solid #334155;">
                  <span style="color:#94a3b8;font-size:0.85rem;min-width:80px">{row_m['เดือน']}</span>
                  <span style="color:{col_p};font-weight:800;font-size:1.1rem;min-width:64px">{p:.1f}%</span>
                  <span style="color:#94a3b8;font-size:0.78rem">
                    <span style="color:#22c55e">●</span> ปกติ {int(row_m.get('มาปกติ',0))}
                    &nbsp;<span style="color:#f59e0b">●</span> สาย {int(row_m.get('มาสาย',0))}
                    &nbsp;<span style="color:#ef4444">●</span> ขาด {int(row_m.get('ขาดงาน',0))}
                    &nbsp;<span style="color:#8b5cf6">●</span> ลืมสแกน {int(row_m.get('ลืมสแกน',0))}
                    &nbsp;รวม {int(row_m.get('วันรวม',0))} วัน
                  </span>
                  <div style="flex:1;background:#0f172a;border-radius:4px;height:8px;overflow:hidden">
                    <div style="width:{min(p,100):.0f}%;background:{col_p};height:100%"></div>
                  </div>
                </div>""", unsafe_allow_html=True)

            st.divider()
            st.dataframe(df_monthly[["เดือน","มาปกติ","มาสาย","ขาดงาน","ลืมสแกน","วันรวม","% มาปกติ"]]
                         .sort_values("เดือน")
                         .style.format({"% มาปกติ":"{:.1f}%"})
                         .applymap(lambda v: f"background-color:#166534;color:#dcfce7" if isinstance(v,float) and v>=80
                                   else (f"background-color:#991b1b;color:#fee2e2" if isinstance(v,float) and v<60 else ""),
                                   subset=["% มาปกติ"]),
                         use_container_width=True)

    # ═══════════════════════════════════════════════════
    # TAB 3: กราฟ 3 ตัว
    # ═══════════════════════════════════════════════════
    with tab_charts:
        st.markdown('<div class="sec-title">📊 กราฟวิเคราะห์</div>', unsafe_allow_html=True)

        if df_att.empty:
            st.info("ไม่มีข้อมูล")
        else:
            df_monthly_c = (
                df_work.groupby("เดือน")["สถานะสแกน"]
                .value_counts().unstack(fill_value=0).reset_index()
            )
            for col in ["มาปกติ","มาสาย","ขาดงาน","ลืมสแกน"]:
                if col not in df_monthly_c.columns: df_monthly_c[col] = 0
            df_monthly_c["วันรวม"] = df_monthly_c[["มาปกติ","มาสาย","ขาดงาน","ลืมสแกน"]].sum(axis=1)
            df_monthly_c["% มาปกติ"] = (df_monthly_c["มาปกติ"] / df_monthly_c["วันรวม"].replace(0,1) * 100).round(1)
            df_monthly_c = df_monthly_c.sort_values("เดือน")

            # ── Chart 1: Line Trend (% มาปกติ) ──────────────────────
            st.subheader("1️⃣ Line Trend — อัตรามาปกติ (%)")
            line_data = df_monthly_c[["เดือน","% มาปกติ"]].copy()
            line_data.columns = ["เดือน","อัตรามาปกติ (%)"]
            line_ch = (
                alt.Chart(line_data)
                .mark_line(point=alt.OverlayMarkDef(filled=True, size=80), strokeWidth=2.5)
                .encode(
                    x=alt.X("เดือน:O", title="เดือน"),
                    y=alt.Y("อัตรามาปกติ (%):Q", title="%", scale=alt.Scale(domain=[0,100])),
                    color=alt.value("#3b82f6"),
                    tooltip=["เดือน","อัตรามาปกติ (%)"],
                )
                + alt.Chart(pd.DataFrame({"y":[80]}))
                .mark_rule(strokeDash=[4,4], color="#22c55e", opacity=0.6)
                .encode(y="y:Q")
            )
            st.altair_chart(line_ch, use_container_width=True)
            st.caption("เส้นประเขียว = เกณฑ์ 80%")

            st.divider()

            # ── Chart 2: Stacked Bar ──────────────────────────────────
            st.subheader("2️⃣ Stacked Bar — จำนวนวันแยกสถานะรายเดือน")
            df_melt = df_monthly_c[["เดือน","มาปกติ","มาสาย","ขาดงาน","ลืมสแกน"]].melt(
                id_vars="เดือน", var_name="สถานะ", value_name="วัน"
            )
            color_scale = alt.Scale(
                domain=["มาปกติ","มาสาย","ขาดงาน","ลืมสแกน"],
                range=["#22c55e","#f59e0b","#ef4444","#8b5cf6"]
            )
            bar_ch = (
                alt.Chart(df_melt)
                .mark_bar()
                .encode(
                    x=alt.X("เดือน:O", title="เดือน"),
                    y=alt.Y("วัน:Q", stack="zero", title="จำนวนวัน"),
                    color=alt.Color("สถานะ:N", scale=color_scale,
                                    legend=alt.Legend(orient="bottom")),
                    tooltip=["เดือน","สถานะ","วัน"],
                )
                .properties(height=300)
            )
            st.altair_chart(bar_ch, use_container_width=True)

            st.divider()

            # ── Chart 3: Rate Line (มาสาย + ขาดงาน) ─────────────────
            st.subheader("3️⃣ Rate Line — แนวโน้มมาสาย & ขาดงาน")
            df_rate = df_monthly_c[["เดือน","มาสาย","ขาดงาน","วันรวม"]].copy()
            df_rate["% มาสาย"]  = (df_rate["มาสาย"]   / df_rate["วันรวม"].replace(0,1) * 100).round(1)
            df_rate["% ขาดงาน"] = (df_rate["ขาดงาน"]  / df_rate["วันรวม"].replace(0,1) * 100).round(1)
            df_rate_melt = df_rate[["เดือน","% มาสาย","% ขาดงาน"]].melt(
                id_vars="เดือน", var_name="ประเภท", value_name="%"
            )
            rate_ch = (
                alt.Chart(df_rate_melt)
                .mark_line(point=alt.OverlayMarkDef(filled=True, size=60), strokeWidth=2)
                .encode(
                    x=alt.X("เดือน:O"),
                    y=alt.Y("%:Q", title="%"),
                    color=alt.Color("ประเภท:N",
                        scale=alt.Scale(domain=["% มาสาย","% ขาดงาน"],
                                        range=["#f59e0b","#ef4444"]),
                        legend=alt.Legend(orient="bottom")),
                    tooltip=["เดือน","ประเภท","%"],
                )
                .properties(height=280)
            )
            st.altair_chart(rate_ch, use_container_width=True)

    # ═══════════════════════════════════════════════════
    # TAB 4: 7 ข้อวิเคราะห์
    # ═══════════════════════════════════════════════════
    with tab_insight:
        st.markdown('<div class="sec-title">🔍 7 ข้อวิเคราะห์สำคัญ</div>', unsafe_allow_html=True)

        if df_att.empty:
            st.info("ไม่มีข้อมูล")
        else:
            insights = []

            # 1. อัตรามาปกติโดยรวม
            status_txt = "✅ ผ่านเกณฑ์ 80%" if pct_ok >= 80 else "⚠️ ต่ำกว่าเกณฑ์ 80%"
            insights.append(("อัตรามาปกติโดยรวม",
                f"อยู่ที่ **{pct_ok:.1f}%** — {status_txt}  "
                f"(มาปกติ {n_ok:,} วัน จากวันทำการรวม {total_work:,} วัน)"))

            # 2. เดือนที่อัตรามาปกติสูงสุด-ต่ำสุด
            if not df_att.empty and "เดือน" in df_att.columns:
                df_m2 = df_work.groupby("เดือน").apply(
                    lambda g: round(len(g[g["สถานะสแกน"]=="มาปกติ"])/max(len(g),1)*100,1)
                ).reset_index(name="%")
                if not df_m2.empty:
                    best = df_m2.loc[df_m2["%"].idxmax()]
                    worst = df_m2.loc[df_m2["%"].idxmin()]
                    insights.append(("เดือนที่มีอัตรามาปกติสูงสุด-ต่ำสุด",
                        f"สูงสุด: **{best['เดือน']}** ({best['%']:.1f}%)  |  "
                        f"ต่ำสุด: **{worst['เดือน']}** ({worst['%']:.1f}%)"))

            # 3. กลุ่มงานที่มีการลามากสุด
            if not df_leave.empty and "กลุ่มงาน" in df_leave.columns and "จำนวนวันลา" in df_leave.columns:
                top_lv = df_leave.groupby("กลุ่มงาน")["จำนวนวันลา"].sum().idxmax()
                top_lv_val = df_leave.groupby("กลุ่มงาน")["จำนวนวันลา"].sum().max()
                insights.append(("กลุ่มงานที่มีวันลารวมมากที่สุด",
                    f"**{top_lv}** มีวันลารวม {int(top_lv_val)} วัน"))

            # 4. วันขาดงานรวม
            insights.append(("จำนวนวันขาดงานรวมทั้งหมด",
                f"**{n_absent:,} วัน** คิดเป็น **{pct_absent:.1f}%** ของวันทำการรวม  "
                f"{'⚠️ สูงกว่า 5% ควรติดตาม' if pct_absent > 5 else '✅ อยู่ในเกณฑ์ปกติ'}"))

            # 5. ไปราชการสูงสุด
            if not df_travel_all.empty and "จำนวนวัน" in df_travel_all.columns:
                max_travel = df_travel_all.groupby("ชื่อ-สกุล")["จำนวนวัน"].sum().idxmax() if not df_travel_all.empty else "-"
                max_travel_val = df_travel_all.groupby("ชื่อ-สกุล")["จำนวนวัน"].sum().max() if not df_travel_all.empty else 0
                insights.append(("ผู้ที่ไปราชการสะสมมากสุด",
                    f"**{max_travel}** รวม {int(max_travel_val)} วัน"))
            else:
                insights.append(("ข้อมูลไปราชการ",
                    f"มีทั้งหมด **{len(df_travel_all):,} รายการ** จาก {len(df_travel_all['ชื่อ-สกุล'].unique()) if not df_travel_all.empty else 0} คน"))

            # 6. ลืมสแกน
            pct_forgot = n_forgot / total_work * 100 if total_work else 0
            insights.append(("สัดส่วนลืมสแกนนิ้ว",
                f"รวม **{n_forgot:,} วัน** คิดเป็น **{pct_forgot:.1f}%**  "
                f"{'⚠️ ควรปรับปรุงระบบหรือแจ้งเตือนบุคลากร' if pct_forgot > 3 else '✅ อยู่ในระดับที่ยอมรับได้'}"))

            # 7. แนวโน้มล่าสุด
            if not df_att.empty:
                recent_months = sorted(df_work["เดือน"].dropna().unique().tolist())
                if len(recent_months) >= 2:
                    def _pct_month(mo):
                        g = df_work[df_work["เดือน"]==mo]
                        return round(len(g[g["สถานะสแกน"]=="มาปกติ"])/max(len(g),1)*100,1)
                    last_m  = recent_months[-1]
                    prev_m  = recent_months[-2]
                    last_p  = _pct_month(last_m)
                    prev_p  = _pct_month(prev_m)
                    diff    = last_p - prev_p
                    arrow   = "📈" if diff > 0 else ("📉" if diff < 0 else "➡️")
                    insights.append(("แนวโน้มอัตรามาปกติ (เดือนล่าสุด vs เดือนก่อน)",
                        f"{arrow} **{last_m}** ({last_p:.1f}%) vs **{prev_m}** ({prev_p:.1f}%)  "
                        f"{'ดีขึ้น +' if diff>0 else 'ลดลง '}{abs(diff):.1f}%  "
                        f"{'✅ อยู่ในเกณฑ์ดี' if last_p>=80 else '⚠️ ต้องติดตาม'}"))

            # render cards
            border_colors = ["#3b82f6","#22c55e","#f59e0b","#ef4444","#8b5cf6","#ec4899","#06b6d4"]
            for i, (title, body) in enumerate(insights):
                bc = border_colors[i % len(border_colors)]
                st.markdown(f"""
                <div class="insight-card" style="border-color:{bc}">
                    <span class="insight-num">{i+1}.</span>
                    <span class="insight-title"> {title}</span>
                    <div class="insight-body">{body}</div>
                </div>""", unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════
    # TAB 5: Export
    # ═══════════════════════════════════════════════════
    with tab_export:
        st.subheader("📥 Export รายงาน")
        today = dt.date.today()
        month_opts_e = pd.date_range(
            start=f"{today.year-2}-01-01", end=f"{today.year+1}-12-31", freq="MS"
        ).strftime("%Y-%m").tolist()
        cur_ym = today.strftime("%Y-%m")
        default_i_e = month_opts_e.index(cur_ym) if cur_ym in month_opts_e else 0
        export_month = st.selectbox("เลือกเดือน", month_opts_e, index=default_i_e)

        if st.button("📊 สร้างรายงาน Excel", use_container_width=True, type="primary"):
            with st.spinner("กำลังสร้างรายงาน..."):
                m_start = pd.to_datetime(export_month + "-01")
                m_end   = m_start + pd.offsets.MonthEnd(0)
                df_lm = df_leave[(df_leave["วันที่เริ่ม"] >= m_start) & (df_leave["วันที่เริ่ม"] <= m_end)] if not df_leave.empty else pd.DataFrame()
                df_tm = df_travel[(df_travel["วันที่เริ่ม"] >= m_start) & (df_travel["วันที่เริ่ม"] <= m_end)] if not df_travel.empty else pd.DataFrame()
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    summary_data = {
                        "รายการ": ["การลา (ครั้ง)","วันลารวม","ไปราชการ (ครั้ง)","วันราชการรวม"],
                        "จำนวน": [
                            len(df_lm),
                            int(df_lm["จำนวนวันลา"].sum()) if not df_lm.empty and "จำนวนวันลา" in df_lm.columns else 0,
                            len(df_tm),
                            int(df_tm["จำนวนวัน"].sum()) if not df_tm.empty and "จำนวนวัน" in df_tm.columns else 0,
                        ],
                    }
                    pd.DataFrame(summary_data).to_excel(writer, sheet_name="📋 สรุป", index=False)
                    if not df_lm.empty: df_lm.to_excel(writer, sheet_name="การลา", index=False)
                    if not df_tm.empty: df_tm.to_excel(writer, sheet_name="ไปราชการ", index=False)
                    if not df_lm.empty and "ประเภทการลา" in df_lm.columns:
                        df_lm.groupby("ประเภทการลา")["จำนวนวันลา"].sum().reset_index().to_excel(
                            writer, sheet_name="ลาแยกประเภท", index=False)
                st.download_button("⬇️ ดาวน์โหลดรายงาน", output.getvalue(),
                                   f"HR_Report_{export_month}.xlsx", mime=EXCEL_MIME,
                                   use_container_width=True)

# ============================================================
# 📅 ตรวจสอบการปฏิบัติงาน
# ============================================================
elif menu == "📅 ตรวจสอบการปฏิบัติงาน":
    st.markdown('<div class="section-header">📅 สรุปการมาปฏิบัติงานรายวัน</div>', unsafe_allow_html=True)

    with st.spinner("กำลังโหลดข้อมูล..."):
        df_att    = read_attendance_report()
        df_leave  = read_excel_with_backup(FILE_LEAVE, dedup_cols=["ชื่อ-สกุล","วันที่เริ่ม","ประเภทการลา"])[0]
        df_staff  = read_excel_from_drive(FILE_STAFF)
        df_travel_all = load_all_travel()

        # ── ขั้นตอนที่ 1: preprocess ก่อนเสมอ (rename + normalize dates + clean names)
        # ต้องทำก่อน merge เพื่อให้ชื่อ column ถูกต้อง
        df_leave, df_travel_pp, df_att = preprocess_dataframes(df_leave, df_travel_all, df_att)

        # ── ขั้นตอนที่ 2: merge manual scans (ตอนนี้ df_att มี "ชื่อ-สกุล" ถูกต้องแล้ว)
        df_manual = load_manual_scans()
        df_att    = merge_attendance_with_manual(df_att, df_manual)

        all_names = get_active_staff(df_staff) or get_all_names_fallback(
            df_leave, df_travel_all, df_att
        )

    # แสดงแหล่งข้อมูลที่โหลดมา
    travel_sources = df_travel_all["_source_file"].unique().tolist() if not df_travel_all.empty and "_source_file" in df_travel_all.columns else [FILE_TRAVEL]
    n_travel = len(df_travel_all)
    st.caption(
        f"📂 ข้อมูลไปราชการ: **{n_travel} รายการ** จาก **{len(travel_sources)} ไฟล์** "
        f"({', '.join(travel_sources)})"
    )

    if df_att.empty:
        months_att = set()
    else:
        df_att["วันที่"] = pd.to_datetime(df_att["วันที่"], errors="coerce").dt.normalize()
        df_att["เดือน"] = df_att["วันที่"].dt.strftime("%Y-%m")
        months_att = set(df_att["เดือน"].dropna().unique().tolist())

    # รวมเดือนจาก travel (ทุกไฟล์) และ leave ด้วย
    months_travel: set = set()
    if not df_travel_all.empty:
        for col in ["วันที่เริ่ม", "วันที่สิ้นสุด"]:
            if col in df_travel_all.columns:
                months_travel.update(
                    df_travel_all[col].dropna()
                    .pipe(lambda s: pd.to_datetime(s, errors="coerce"))
                    .dt.to_period("M").astype(str).unique().tolist()
                )

    months_leave: set = set()
    if not df_leave.empty and "วันที่เริ่ม" in df_leave.columns:
        months_leave = set(
            df_leave["วันที่เริ่ม"].dropna()
            .pipe(lambda s: pd.to_datetime(s, errors="coerce"))
            .dt.strftime("%Y-%m")
            .dropna()
            .unique()
            .tolist()
        )

    months = sorted(months_att | months_travel | months_leave) or [dt.datetime.now().strftime("%Y-%m")]

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
    # normalize names ก่อนเปรียบเทียบ
    names_to_process = [
        str(n).strip().replace("  ", " ") for n in (selected_names or all_names)
    ]
    if sel_att_group != "ทุกกลุ่ม" and not df_staff_att.empty and "กลุ่มงาน" in df_staff_att.columns:
        grp_set = set(
            df_staff_att[df_staff_att["กลุ่มงาน"] == sel_att_group]["ชื่อ-สกุล"]
            .astype(str).str.strip().str.replace(r"\s+", " ", regex=True).tolist()
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
        # ตรวจ name_col อีกครั้งใน df_months_att (อาจต่างจาก df_att ถ้า merge เพิ่ม columns)
        actual_name_col = next(
            (c for c in ["ชื่อ-สกุล", "ชื่อพนักงาน", "ชื่อ"] if c in df_months_att.columns),
            name_col,
        )
        if actual_name_col != "ชื่อ-สกุล" and actual_name_col in df_months_att.columns:
            df_months_att = df_months_att.rename(columns={actual_name_col: "ชื่อ-สกุล"})

        # normalize ชื่อ: strip + collapse whitespace
        df_months_att["ชื่อ-สกุล"] = (
            df_months_att["ชื่อ-สกุล"]
            .astype(str).str.strip()
            .str.replace(r"\s+", " ", regex=True)
        )
        # normalize วันที่
        df_months_att["วันที่"] = pd.to_datetime(
            df_months_att["วันที่"], errors="coerce"
        ).dt.normalize()
        df_months_att["_date"] = df_months_att["วันที่"].dt.date
        # ใช้ "ชื่อ-สกุล" เสมอหลัง normalize
        name_col = "ชื่อ-สกุล"

    WORK_START = dt.time(8, 30)
    WORK_END   = dt.time(16, 30)

    # ── โหลดวันหยุดพิเศษ ─────────────────────────────────────
    sel_years = list({int(ym[:4]) for ym in selected_months})
    holiday_dates_set: set = set()
    holiday_df_lookup = pd.DataFrame()
    for yr in sel_years:
        holiday_dates_set.update(get_holiday_dates(yr))
        hdf = load_holidays_all(yr)
        holiday_df_lookup = pd.concat([holiday_df_lookup, hdf], ignore_index=True)

    # ── Pre-index travel รายคน (จากทุกไฟล์) ─────────────────
    travel_index: Dict[str, List[tuple]] = {}
    if not df_travel_pp.empty and "ชื่อ-สกุล" in df_travel_pp.columns:
        df_tr = df_travel_pp.copy()
        df_tr["วันที่เริ่ม"]   = pd.to_datetime(df_tr.get("วันที่เริ่ม"),   errors="coerce").dt.normalize()
        df_tr["วันที่สิ้นสุด"] = pd.to_datetime(df_tr.get("วันที่สิ้นสุด"), errors="coerce").dt.normalize()
        df_tr["ชื่อ-สกุล"]    = df_tr["ชื่อ-สกุล"].astype(str).str.strip()
        df_tr = df_tr.dropna(subset=["วันที่เริ่ม", "วันที่สิ้นสุด"])

        for _, row in df_tr.iterrows():
            s    = row["วันที่เริ่ม"].date()
            e    = row["วันที่สิ้นสุด"].date()
            proj = str(row.get("เรื่อง/กิจกรรม", "ไปราชการ")).strip() or "ไปราชการ"
            src  = str(row.get("_source_file", FILE_TRAVEL))

            # ── รวมชื่อคนหลัก + ผู้ร่วมเดินทางทั้งหมด ──────────────
            names_in_trip: List[str] = []

            # 1. คนหลัก
            main_name = row["ชื่อ-สกุล"]
            if main_name and main_name.lower() != "nan":
                names_in_trip.append(main_name)

            # 2. ผู้ร่วมเดินทาง — คั่นด้วย "," หรือขึ้นบรรทัดใหม่
            companions_raw = str(row.get("ผู้ร่วมเดินทาง", "")).strip()
            if companions_raw and companions_raw not in ("-", "nan", ""):
                # ทำความสะอาด: ตัดเลข "1.", "2.", "3." ที่อาจแอบมา
                import re as _re
                companions_raw = _re.sub(r"\d+\.\s*", "", companions_raw)
                for comp in companions_raw.replace("\n", ",").split(","):
                    comp = comp.strip()
                    # กรองเอาเฉพาะชื่อที่ดูเหมือนชื่อบุคคล (มีตัวอักษรไทย ≥ 3 ตัว)
                    if comp and len(comp) >= 3 and comp.lower() != "nan":
                        names_in_trip.append(comp)

            # 3. บันทึกลง index ทุกคนในทริปนี้
            for person in set(names_in_trip):  # set() ป้องกันซ้ำ
                travel_index.setdefault(person, []).append((s, e, proj, src))

    # ── Pre-index leave รายคน ─────────────────────────────────
    leave_index: Dict[str, List[tuple]] = {}
    if not df_leave.empty and "ชื่อ-สกุล" in df_leave.columns:
        df_lv = df_leave.copy()
        df_lv["วันที่เริ่ม"]   = pd.to_datetime(df_lv.get("วันที่เริ่ม"),   errors="coerce").dt.normalize()
        df_lv["วันที่สิ้นสุด"] = pd.to_datetime(df_lv.get("วันที่สิ้นสุด"), errors="coerce").dt.normalize()
        df_lv["ชื่อ-สกุล"]    = df_lv["ชื่อ-สกุล"].astype(str).str.strip()
        df_lv = df_lv.dropna(subset=["วันที่เริ่ม", "วันที่สิ้นสุด"])
        for _, row in df_lv.iterrows():
            name_l = row["ชื่อ-สกุล"]
            if not name_l or name_l.lower() == "nan":
                continue
            s     = row["วันที่เริ่ม"].date()
            e     = row["วันที่สิ้นสุด"].date()
            ltype = str(row.get("ประเภทการลา", "ลา")).strip() or "ลา"
            leave_index.setdefault(name_l, []).append((s, e, ltype))

    # ── คำนวณข้อมูลรายวัน ────────────────────────────────────
    records = []
    prog = st.progress(0, text="กำลังประมวลผล...")
    for i, name in enumerate(names_to_process):
        prog.progress((i + 1) / len(names_to_process), text=f"กำลังประมวลผล {name}...")
        for d in all_dates:
            d_date = d.date()   # ใช้ date object เพื่อ comparison ที่ consistent
            rec = {
                "ชื่อพนักงาน": name,
                "วันที่":       d_date,
                "เดือน":        d.strftime("%Y-%m"),
                "เวลาเข้า":     "",
                "เวลาออก":      "",
                "สถานะ":        "",
            }
            # ── lookup สแกน (ใช้ _date column ที่ pre-compute แล้ว ไม่เรียก .dt.date ใน loop)
            att = (
                df_months_att[
                    (df_months_att[name_col] == name)
                    & (df_months_att["_date"] == d_date)
                ]
                if not df_months_att.empty else pd.DataFrame()
            )

            # ── วันหยุดพิเศษ
            is_special_hday = d_date in holiday_dates_set
            special_hday_name = get_holiday_name(d_date, holiday_df_lookup) if is_special_hday else ""

            # ── ตรวจการลา (ใช้ pre-index)
            in_leave, leave_type = False, ""
            for ls, le, ltype in leave_index.get(name, []):
                if ls <= d_date <= le:
                    in_leave, leave_type = True, ltype
                    break

            # ── ตรวจไปราชการ (ใช้ pre-index จากทุกไฟล์) + เก็บชื่อโครงการ
            in_travel, travel_project = False, ""
            for ts, te, proj, _src in travel_index.get(name, []):
                if ts <= d_date <= te:
                    in_travel, travel_project = True, proj
                    break

            if in_leave:
                rec["สถานะ"] = f"ลา ({leave_type})"
            elif in_travel:
                rec["สถานะ"] = f"ไปราชการ ({travel_project})" if travel_project and travel_project != "ไปราชการ" else "ไปราชการ"
            elif d.weekday() >= 5:
                rec["สถานะ"] = "วันหยุด"
            elif is_special_hday:
                # วันหยุดพิเศษ — แสดงชื่อ
                rec["สถานะ"] = f"วันหยุด ({special_hday_name})"
            elif not att.empty:
                row = att.iloc[0]
                rec["เวลาเข้า"] = row.get("เวลาเข้า", "")
                rec["เวลาออก"]  = row.get("เวลาออก",  "")
                is_manual = str(row.get("_source", "")).strip() == "manual"

                t_in  = parse_time(rec["เวลาเข้า"])
                t_out = parse_time(rec["เวลาออก"])

                LATE_CUTOFF = dt.time(8, 31)   # เกินนี้ = มาสาย

                # ── กฎ 1: ไม่มีทั้งเข้าและออก
                if not t_in and not t_out:
                    base_status = "ขาดงาน"

                # ── กฎ 2 & 3: มีอย่างใดอย่างหนึ่ง
                elif (t_in and not t_out) or (not t_in and t_out):
                    base_status = "ลืมสแกน"

                # ── กฎ 4: เข้า-ออกเวลาเดียวกัน
                elif t_in == t_out:
                    base_status = "ลืมสแกน"

                # ── กฎ 5: เข้างาน 08:31 ขึ้นไป
                elif t_in >= LATE_CUTOFF:
                    base_status = "มาสาย"

                # ── ปกติ: เข้าทัน ≤ 08:30
                else:
                    base_status = "มาปกติ"

                rec["สถานะ"] = f"{base_status} (HR คีย์แทน)" if is_manual else base_status
            else:
                rec["สถานะ"] = "วันหยุด" if d.weekday() >= 5 else "ขาดงาน"
            records.append(rec)
    prog.empty()

    df_daily = pd.DataFrame(records).sort_values(["ชื่อพนักงาน","วันที่"])

    def simplify_status(s):
        """ย่อ status สำหรับ pivot table"""
        if not isinstance(s, str):
            return s
        if s.startswith("ลา"):
            return "ลา"
        if s.startswith("วันหยุด"):
            return "วันหยุด"
        if s.startswith("ไปราชการ"):
            return "ไปราชการ"
        clean = s.replace(" (HR คีย์แทน)", "").strip()
        return clean
    df_daily["สถานะย่อ"] = df_daily["สถานะ"].apply(simplify_status)

    STATUS_COLORS = {
        "มาปกติ":    "background-color:#d4edda",
        "มาสาย":     "background-color:#ffeeba",
        "ลืมสแกน":  "background-color:#ede9fe",
        "ขาดงาน":   "background-color:#f5c6cb",
        "ลา":        "background-color:#d1ecf1",
        "ไปราชการ":  "background-color:#fff3cd",
        "วันหยุด":   "background-color:#e2e3e5",
        "HR คีย์แทน": "background-color:#d1fae5",
    }
    def color_status(val):
        s = str(val)
        # ตรวจ suffix HR คีย์แทน ก่อน
        if "HR คีย์แทน" in s:
            return STATUS_COLORS["HR คีย์แทน"]
        for k, v in STATUS_COLORS.items():
            if k in s:
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
        required_cols = ["มาปกติ","มาสาย","ลืมสแกน","ขาดงาน","ลา","ไปราชการ","วันหยุด"]

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
            all_work  = summary[[c for c in ["มาปกติ","มาสาย","ลืมสแกน","ขาดงาน"] if c in summary.columns]].sum(axis=1)
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
            default=["มาปกติ","มาสาย","ลืมสแกน","ขาดงาน","ลา","ไปราชการ"],
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
        df_leave  = read_excel_with_backup(FILE_LEAVE, dedup_cols=["ชื่อ-สกุล","วันที่เริ่ม","ประเภทการลา"])[0]
        df_travel = read_excel_with_backup(FILE_TRAVEL, dedup_cols=["ชื่อ-สกุล","วันที่เริ่ม","เรื่อง/กิจกรรม"])[0]
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
        # ── อ่านพร้อม file_id เพื่อใช้ update in-place ────────────
        df_travel, _travel_fid = read_excel_with_backup(FILE_TRAVEL, dedup_cols=["ชื่อ-สกุล","วันที่เริ่ม","เรื่อง/กิจกรรม"])
        df_leave  = read_excel_with_backup(FILE_LEAVE, dedup_cols=["ชื่อ-สกุล","วันที่เริ่ม","ประเภทการลา"])[0]
        df_att    = read_attendance_report()
        df_staff  = read_excel_from_drive(FILE_STAFF)
        df_leave, df_travel, df_att = preprocess_dataframes(df_leave, df_travel, df_att)
        ALL_NAMES = get_active_staff(df_staff) or get_all_names_fallback(df_leave, df_travel, df_att)

    st.info(f"📂 ข้อมูลไปราชการปัจจุบัน: **{len(df_travel)} รายการ**  "
            f"{'(file ID: ' + _travel_fid[:8] + '...)' if _travel_fid else '⚠️ ยังไม่มีไฟล์ใน Drive'}")

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
                            fid_att = get_or_create_folder(ATTACHMENT_FOLDER_NAME, FOLDER_ID)
                            if fid_att:
                                fn = f"TRAVEL_{dt.datetime.now().strftime('%Y%m%d_%H%M')}_{len(final_staff)}pax.pdf"
                                link = upload_pdf_to_drive(uploaded_pdf, fn, fid_att)

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

                        # ส่ง known_file_id เพื่อให้ update ไฟล์เดิม ไม่สร้างใหม่
                        if write_excel_to_drive(FILE_TRAVEL, df_upd, known_file_id=_travel_fid):
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
        df_leave, _leave_fid = read_excel_with_backup(FILE_LEAVE, dedup_cols=["ชื่อ-สกุล","วันที่เริ่ม","ประเภทการลา"])
        df_travel = read_excel_with_backup(FILE_TRAVEL, dedup_cols=["ชื่อ-สกุล","วันที่เริ่ม","เรื่อง/กิจกรรม"])[0]
        df_att    = read_attendance_report()
        df_staff  = read_excel_from_drive(FILE_STAFF)
        df_leave, df_travel, df_att = preprocess_dataframes(df_leave, df_travel, df_att)
        ALL_NAMES = get_active_staff(df_staff) or get_all_names_fallback(df_leave, df_travel, df_att)

    st.info(f"📂 ข้อมูลการลาปัจจุบัน: **{len(df_leave)} รายการ**  "
            f"{'(file ID: ' + _leave_fid[:8] + '...)' if _leave_fid else '⚠️ ยังไม่มีไฟล์ใน Drive'}")

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

                        if write_excel_to_drive(FILE_LEAVE, df_upd, known_file_id=_leave_fid):
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
        df_leave = read_excel_with_backup(FILE_LEAVE, dedup_cols=["ชื่อ-สกุล","วันที่เริ่ม","ประเภทการลา"])[0]
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
        df_staff, _staff_fid = read_excel_with_backup(FILE_STAFF, dedup_cols=["ชื่อ-สกุล"])

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
                    if write_excel_to_drive(FILE_STAFF, df_staff, known_file_id=_staff_fid):
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
                        if write_excel_to_drive(FILE_STAFF, df_staff, known_file_id=_staff_fid):
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
            df_leave,  _fid_leave  = read_excel_with_backup(FILE_LEAVE, dedup_cols=["ชื่อ-สกุล","วันที่เริ่ม","ประเภทการลา"])
            df_travel, _fid_travel = read_excel_with_backup(FILE_TRAVEL, dedup_cols=["ชื่อ-สกุล","วันที่เริ่ม","เรื่อง/กิจกรรม"])
            df_att    = read_attendance_report()
            df_staff,  _fid_staff  = read_excel_with_backup(FILE_STAFF, dedup_cols=["ชื่อ-สกุล"])

        _fid_map = {
            FILE_LEAVE:  _fid_leave,
            FILE_TRAVEL: _fid_travel,
            FILE_STAFF:  _fid_staff,
            FILE_ATTEND: None,   # attendance ไม่ได้แก้จาก admin upload
        }

        tab1, tab2, tab3, tab4, tab5, tab6, tab_hol = st.tabs([
            "📂 ไฟล์ลา", "📂 ไฟล์ราชการ",
            "📂 ไฟล์สแกนนิ้ว", "📂 ไฟล์บุคลากร",
            "🔧 ตั้งค่าระบบ", "👆 คีย์ลืมสแกนนิ้ว", "🎌 วันหยุดพิเศษ",
        ])

        def admin_file_panel(df, filename, tab_obj):
            with tab_obj:
                st.subheader(f"ไฟล์: {filename}")
                st.caption(f"File ID ปัจจุบัน: `{_fid_map.get(filename,'—')}`")
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
                            # ใช้ known_file_id ที่โหลดมาตอนแรก → update in-place
                            known_fid = _fid_map.get(filename)
                            if write_excel_to_drive(filename, new_df, known_file_id=known_fid):
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

            st.divider()

            # ─── 🧹 ล้างไฟล์ Drive ────────────────────────────────
            st.subheader("🧹 ล้างไฟล์ซ้ำใน Drive")
            st.info(
                "สแกนเฉพาะ **root folder** (`Leave_App_Data/`) — "
                "ลบไฟล์ข้อมูลที่ชื่อซ้ำกัน เก็บไว้เฉพาะไฟล์ล่าสุด 1 ไฟล์ต่อชื่อ  \n"
                "📁 โฟลเดอร์ `Backup/` จะ **ไม่ถูกแตะเลย**"
            )

            with st.spinner("กำลังสแกนไฟล์ใน Drive..."):
                all_drive_files = list_all_files_in_folder()   # scan root only

            # แบ่งประเภทไฟล์ใน root
            import re
            dup_map:  dict = {}   # filename → [(id, modifiedTime), ...]
            misc_bak: list = []   # BAK_* ที่ยังหลุดอยู่ใน root (ไม่ควรมีแล้ว)

            for f in all_drive_files:
                fname = f["name"]
                if re.match(r"^BAK_", fname):
                    misc_bak.append(f)   # BAK ใน root = หลงเหลือจากระบบเก่า
                else:
                    dup_map.setdefault(fname, []).append(f)

            # หา duplicate (ชื่อเดียวกัน > 1 ไฟล์)
            dup_to_delete = []
            for fname, flist in dup_map.items():
                if len(flist) > 1:
                    # flist เรียง modifiedTime desc แล้ว — เก็บอันแรก
                    for dup_f in flist[1:]:
                        dup_to_delete.append({"name": fname, "id": dup_f["id"],
                                              "วันที่แก้ไข": dup_f.get("modifiedTime","")[:10]})

            total_to_delete = len(misc_bak) + len(dup_to_delete)

            # แสดงโครงสร้าง Backup folder
            backup_root_id = get_or_create_folder(BACKUP_FOLDER_NAME, FOLDER_ID)
            st.markdown("**📁 โครงสร้าง Backup folder ปัจจุบัน:**")
            bk_col1, bk_col2, bk_col3 = st.columns(3)
            data_files = [FILE_ATTEND, FILE_LEAVE, FILE_TRAVEL]
            for i, dfile in enumerate(data_files):
                bak_sub_id  = get_or_create_folder(f"BAK_{dfile}", backup_root_id) if backup_root_id else None
                bak_file_id = get_file_id(f"BAK_{dfile}", bak_sub_id) if bak_sub_id else None
                col = [bk_col1, bk_col2, bk_col3][i]
                col.markdown(
                    f"{'🟢' if bak_file_id else '🔴'} `BAK_{dfile}`  \n"
                    f"{'✅ มี backup' if bak_file_id else '⚠️ ยังไม่มี'}"
                )

            st.divider()

            # Preview + ปุ่มลบ
            cl1, cl2 = st.columns(2)
            cl1.metric("♻️ ไฟล์ซ้ำใน root (duplicate)", f"{len(dup_to_delete)} ไฟล์")
            cl2.metric("⚠️ BAK หลุดอยู่ใน root",        f"{len(misc_bak)} ไฟล์")

            if total_to_delete == 0:
                st.success("✅ Drive root สะอาดแล้ว ไม่มีไฟล์ที่ต้องล้าง")
            else:
                with st.expander(f"📋 ดูรายการที่จะถูกลบ ({total_to_delete} ไฟล์)", expanded=True):
                    preview_rows = (
                        [{"ชื่อไฟล์": d["name"], "วันที่แก้ไข": d.get("วันที่แก้ไข",""), "ประเภท": "ไฟล์ซ้ำ (ลบ duplicate เก่า)"} for d in dup_to_delete]
                        + [{"ชื่อไฟล์": f["name"], "วันที่แก้ไข": f.get("modifiedTime","")[:10], "ประเภท": "BAK หลุดใน root"} for f in misc_bak]
                    )
                    st.dataframe(pd.DataFrame(preview_rows), use_container_width=True)

                st.warning(f"⚠️ จะลบ {total_to_delete} ไฟล์ออกจาก root อย่างถาวร (ไฟล์ใน Backup/ จะไม่ถูกแตะ)")

                if st.button(
                    f"🧹 ล้างไฟล์ซ้ำ ({total_to_delete} ไฟล์)",
                    key="btn_drive_cleanup",
                    type="primary",
                    use_container_width=True,
                ):
                    success_count, fail_count = 0, 0
                    all_to_delete = (
                        [(d["id"], d["name"]) for d in dup_to_delete]
                        + [(f["id"], f["name"]) for f in misc_bak]
                    )
                    cleanup_prog = st.progress(0, text="กำลังลบไฟล์...")
                    for idx, (fid_del, fname_del) in enumerate(all_to_delete):
                        cleanup_prog.progress(
                            (idx + 1) / len(all_to_delete),
                            text=f"กำลังลบ {fname_del}...",
                        )
                        try:
                            service.files().delete(fileId=fid_del, supportsAllDrives=True).execute()
                            success_count += 1
                            logger.info(f"Cleanup deleted: {fname_del}")
                        except Exception as del_err:
                            fail_count += 1
                            logger.warning(f"Cleanup failed: {fname_del} — {del_err}")

                    cleanup_prog.empty()
                    st.cache_data.clear()
                    log_activity("ล้างไฟล์ Drive", f"ลบ {success_count} ไฟล์ซ้ำใน root", "Admin")

                    if fail_count == 0:
                        st.toast(f"✅ ล้างสำเร็จ {success_count} ไฟล์", icon="🧹")
                        st.success(f"✅ เสร็จสิ้น — ลบไป {success_count} ไฟล์")
                    else:
                        st.warning(f"ลบสำเร็จ {success_count} | ล้มเหลว {fail_count} (ดู log)")
                    time.sleep(1)
                    st.rerun()

        # ============================================================
        # 👆 Tab 6 — คีย์ลืมสแกนนิ้ว  (บันทึกลง manual_scan.xlsx แยกต่างหาก)
        # ============================================================
        with tab6:
            st.subheader("👆 บันทึกเวลาทำการสำหรับผู้ที่ลืมสแกนนิ้ว")
            st.info(
                "✅ ข้อมูลที่คีย์จะถูกบันทึกลง **`manual_scan.xlsx`** แยกต่างหาก "
                "— ไม่แตะ `attendance_report.xlsx` เลย  "
                "ระบบตรวจสอบการปฏิบัติงานจะ **merge** ข้อมูลทั้งสองไฟล์ให้อัตโนมัติ"
            )

            # ─── โหลดข้อมูล ────────────────────────────────────────
            all_staff_names = get_active_staff(df_staff) or get_all_names_fallback(
                df_leave, df_travel, df_att
            )
            df_manual_tab = load_manual_scans()
            _manual_fid   = get_file_id(FILE_MANUAL_SCAN)   # lock file_id ตอนโหลด

            st.markdown("---")

            # ─── ฟอร์มคีย์ข้อมูล ────────────────────────────────────
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
                    ms_time_in  = st.time_input("เวลาเข้างาน *",  value=dt.time(8, 30),  step=60)
                    ms_time_out = st.time_input("เวลาออกงาน *",   value=dt.time(16, 30), step=60)

                ms_note = st.text_input(
                    "หมายเหตุเพิ่มเติม",
                    value="Admin คีย์แทน — ลืมสแกนนิ้ว",
                )
                ms_submit = st.form_submit_button(
                    "💾 บันทึกข้อมูลสแกนนิ้ว",
                    use_container_width=True,
                    type="primary",
                )

                if ms_submit:
                    scan_errors: List[str] = []
                    if not ms_name:
                        scan_errors.append("❌ กรุณาเลือกชื่อ")
                    if ms_time_in >= ms_time_out:
                        scan_errors.append("❌ เวลาเข้างานต้องน้อยกว่าเวลาออกงาน")

                    ms_date_ts = pd.to_datetime(ms_date).normalize()

                    # ตรวจซ้ำใน manual_scan.xlsx
                    if not df_manual_tab.empty:
                        dup_manual = df_manual_tab[
                            (df_manual_tab["ชื่อ-สกุล"] == ms_name)
                            & (pd.to_datetime(df_manual_tab["วันที่"], errors="coerce").dt.normalize() == ms_date_ts)
                        ]
                        if not dup_manual.empty:
                            scan_errors.append(
                                f"⚠️ มีข้อมูลที่ Admin คีย์แล้วสำหรับ {ms_name} "
                                f"วันที่ {ms_date.strftime('%d/%m/%Y')} "
                                f"(เข้า {dup_manual.iloc[0].get('เวลาเข้า','?')} "
                                f"ออก {dup_manual.iloc[0].get('เวลาออก','?')})"
                            )

                    # ตรวจซ้ำใน attendance_report.xlsx
                    df_att_check = read_attendance_report()
                    if not df_att_check.empty and "วันที่" in df_att_check.columns:
                        df_att_check["วันที่"] = pd.to_datetime(df_att_check["วันที่"], errors="coerce").dt.normalize()
                        name_col_chk = next(
                            (c for c in ["ชื่อ-สกุล","ชื่อพนักงาน","ชื่อ"] if c in df_att_check.columns),
                            "ชื่อ-สกุล",
                        )
                        if name_col_chk in df_att_check.columns:
                            dup_att = df_att_check[
                                (df_att_check[name_col_chk].astype(str).str.strip() == ms_name)
                                & (df_att_check["วันที่"] == ms_date_ts)
                            ]
                            if not dup_att.empty:
                                scan_errors.append(
                                    f"⚠️ มีข้อมูลสแกนจริงของ {ms_name} วันที่ {ms_date.strftime('%d/%m/%Y')} "
                                    f"อยู่ใน attendance_report.xlsx แล้ว "
                                    f"(เข้า {dup_att.iloc[0].get('เวลาเข้า','?')} "
                                    f"ออก {dup_att.iloc[0].get('เวลาออก','?')}) — ไม่จำเป็นต้องคีย์ซ้ำ"
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
                                new_row = {
                                    "ชื่อ-สกุล": ms_name,
                                    "วันที่":     pd.to_datetime(ms_date),
                                    "เวลาเข้า":   ms_time_in.strftime("%H:%M"),
                                    "เวลาออก":    ms_time_out.strftime("%H:%M"),
                                    "หมายเหตุ":   note_full,
                                }
                                df_manual_upd = pd.concat(
                                    [df_manual_tab, pd.DataFrame([new_row])],
                                    ignore_index=True,
                                ).sort_values(["ชื่อ-สกุล", "วันที่"]).reset_index(drop=True)

                                # บันทึกลง manual_scan.xlsx เท่านั้น
                                if write_excel_to_drive(FILE_MANUAL_SCAN, df_manual_upd, known_file_id=_manual_fid):
                                    log_activity(
                                        "คีย์สแกนนิ้ว",
                                        f"Admin คีย์ {ms_date} เข้า {ms_time_in.strftime('%H:%M')} "
                                        f"ออก {ms_time_out.strftime('%H:%M')}",
                                        ms_name,
                                    )
                                    df_manual_tab = df_manual_upd
                                    status.update(label="✅ บันทึกสำเร็จ!", state="complete")
                                    st.toast(
                                        f"✅ บันทึกสแกนนิ้วของ {ms_name} "
                                        f"วันที่ {ms_date.strftime('%d/%m/%Y')} สำเร็จ",
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

            # ─── ตรวจสอบข้อมูลที่คีย์ไว้ ────────────────────────────
            st.subheader("🔍 ตรวจสอบข้อมูลที่ Admin คีย์ไว้")

            col_s1, col_s2, col_s3 = st.columns([2, 1, 1])
            with col_s1:
                search_ms_name = st.selectbox(
                    "เลือกชื่อบุคลากร", all_staff_names, key="search_ms_name"
                )
            with col_s2:
                if not df_manual_tab.empty:
                    ms_month_opts = sorted(
                        pd.to_datetime(df_manual_tab["วันที่"], errors="coerce")
                        .dt.strftime("%Y-%m").dropna().unique().tolist()
                    )
                else:
                    ms_month_opts = [dt.date.today().strftime("%Y-%m")]
                search_ms_month = st.selectbox(
                    "เดือน", ms_month_opts,
                    index=len(ms_month_opts) - 1,
                    key="search_ms_month",
                )
            with col_s3:
                st.write("")
                st.write("")
                ms_show_all = st.checkbox("แสดงทุกเดือน", key="ms_show_all")

            if df_manual_tab.empty:
                st.info("ยังไม่มีข้อมูลที่ Admin คีย์ไว้ในระบบ")
            else:
                df_ms_view = df_manual_tab[
                    df_manual_tab["ชื่อ-สกุล"] == search_ms_name
                ].copy()
                df_ms_view["วันที่"] = pd.to_datetime(df_ms_view["วันที่"], errors="coerce")

                if not ms_show_all:
                    df_ms_view = df_ms_view[
                        df_ms_view["วันที่"].dt.strftime("%Y-%m") == search_ms_month
                    ]
                df_ms_view = df_ms_view.sort_values("วันที่", ascending=False)

                if df_ms_view.empty:
                    st.info(f"ไม่พบข้อมูลของ {search_ms_name} ในเดือน {search_ms_month}")
                else:
                    st.caption(f"พบ {len(df_ms_view)} รายการ")
                    st.dataframe(
                        df_ms_view[["วันที่","เวลาเข้า","เวลาออก","หมายเหตุ"]],
                        use_container_width=True,
                        height=300,
                    )

                    st.divider()
                    st.subheader("🗑️ ลบรายการที่คีย์ผิด")

                    if df_ms_view.empty:
                        st.info("ไม่มีรายการในช่วงที่เลือก")
                    else:
                        df_ms_view["label"] = df_ms_view.apply(
                            lambda r: (
                                f"{r['วันที่'].strftime('%d/%m/%Y') if pd.notna(r['วันที่']) else '?'} "
                                f"— เข้า {r.get('เวลาเข้า','?')} ออก {r.get('เวลาออก','?')}"
                            ),
                            axis=1,
                        )
                        del_ms_label = st.selectbox(
                            "เลือกรายการที่ต้องการลบ",
                            df_ms_view["label"].tolist(),
                            key="del_ms_select",
                        )
                        if st.button("🗑️ ลบรายการนี้", key="btn_del_ms", type="primary"):
                            idx_drop = df_ms_view[df_ms_view["label"] == del_ms_label].index.tolist()
                            if idx_drop:
                                df_manual_new = df_manual_tab.drop(index=idx_drop).reset_index(drop=True)
                                if write_excel_to_drive(FILE_MANUAL_SCAN, df_manual_new, known_file_id=_manual_fid):
                                    log_activity("ลบสแกนนิ้ว", del_ms_label, search_ms_name)
                                    st.toast("✅ ลบรายการสำเร็จ", icon="🗑️")
                                    time.sleep(1)
                                    st.rerun()

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
            df_hol_custom, _hol_fid = load_holidays_with_id()
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

                        if write_excel_to_drive(FILE_HOLIDAYS, df_hol_new, known_file_id=_hol_fid):
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
                            if write_excel_to_drive(FILE_HOLIDAYS, df_hol_after, known_file_id=_hol_fid):
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
