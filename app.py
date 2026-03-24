# ====================================================
# 📋 ระบบติดตามการลาและไปราชการ สคร.9
# ✨ v3.0 — Refactored & Optimized Edition
# ====================================================

import io
import os
import time
import logging
import datetime as dt
import requests
import re
import math
import threading
import gc
from typing import Dict, List, Optional, Tuple

# ลด malloc heap fragmentation ป้องกัน "double linked list corrupted"
os.environ.setdefault("MALLOC_TRIM_THRESHOLD_", "100000")

import numpy as np
import pandas as pd
import altair as alt
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from googleapiclient.errors import HttpError
import ssl

# ===========================
# 🔧 Logging
# ===========================
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s", handlers=[logging.StreamHandler()])
logger = logging.getLogger(__name__)

# ===========================
# 📱 Custom CSS
# ===========================
CUSTOM_CSS = """
<style>
html, body, [class*="css"] { font-family: 'Sarabun', sans-serif; }
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0f172a 0%, #1e293b 100%); color: white;
}
section[data-testid="stSidebar"] * { color: white !important; }
section[data-testid="stSidebar"] .stRadio > label { 
    background: rgba(255,255,255,0.05); border-radius: 8px;
    padding: 6px 12px; margin: 2px 0; display: block; transition: background 0.2s;
}
section[data-testid="stSidebar"] .stRadio > label:hover { background: rgba(255,255,255,0.15); }
div[data-testid="metric-container"] {
    background: white; border-radius: 12px; padding: 16px;
    border: 1px solid #e2e8f0; box-shadow: 0 1px 3px rgba(0,0,0,0.08);
}
.badge-green  { background:#dcfce7; color:#166534; padding:2px 10px; border-radius:999px; font-size:0.78rem; font-weight:600; }
.badge-yellow { background:#fef9c3; color:#854d0e; padding:2px 10px; border-radius:999px; font-size:0.78rem; font-weight:600; }
.badge-red    { background:#fee2e2; color:#991b1b; padding:2px 10px; border-radius:999px; font-size:0.78rem; font-weight:600; }
.badge-blue   { background:#dbeafe; color:#1e40af; padding:2px 10px; border-radius:999px; font-size:0.78rem; font-weight:600; }
.badge-gray   { background:#f1f5f9; color:#475569; padding:2px 10px; border-radius:999px; font-size:0.78rem; font-weight:600; }
.section-header {
    background: linear-gradient(90deg, #0ea5e9, #6366f1);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    font-size: 1.4rem; font-weight: 700; margin-bottom: 1rem;
}
.activity-item { padding: 10px 14px; border-left: 3px solid #6366f1; background: #f8fafc; border-radius: 0 8px 8px 0; margin-bottom: 8px; font-size: 0.87rem; }
.quota-bar-wrap { background:#e2e8f0; border-radius:999px; height:10px; margin:4px 0; }
.quota-bar-fill { height:10px; border-radius:999px; transition: width 0.4s; }
@media (max-width: 768px) { div[data-testid="metric-container"] { margin-bottom: 8px; } .block-container { padding: 1rem !important; } }
</style>
<link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap" rel="stylesheet">
"""

# ===========================
# 🔐 App Init
# ===========================
st.set_page_config(page_title="สคร.9 — HR Tracking v3", page_icon="📋", layout="wide", initial_sidebar_state="expanded")

# Silence Streamlit deprecation warnings ที่ไม่กระทบการทำงาน
import warnings
warnings.filterwarnings("ignore", message=".*use_container_width.*")
warnings.filterwarnings("ignore", message=".*dayfirst.*")
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

if "gcp_service_account" not in st.secrets:
    st.error("❌ ไม่พบ gcp_service_account ใน secrets.toml"); st.stop()

# ===========================
# ⚙️ Configuration
# ===========================
LEAVE_QUOTA: Dict[str, int] = {
    "ลาป่วย": 90, "ลากิจส่วนตัว": 45, "ลาพักผ่อน": 10,
    "ลาคลอดบุตร": 98, "ลาอุปสมบท": 120, "ลาช่วยเหลือภริยาที่คลอดบุตร": 15,
}
STAFF_GROUPS: List[str] = [
    "กลุ่มบริหารทั่วไป","กลุ่มบริหารทั่วไป (งานธุรการ)","กลุ่มบริหารทั่วไป (งานการเงินและบัญชี)",
    "กลุ่มบริหารทั่วไป (งานการเจ้าหน้าที่)","กลุ่มบริหารทั่วไป (งานพัสดุและยานพาหนะ (งานพัสดุ))",
    "กลุ่มบริหารทั่วไป (งานพัสดุและยานพาหนะ (งานยานพาหนะ))","กลุ่มบริหารทั่วไป (งานพัสดุและยานพาหนะ (งานอาคารสถานที่))",
    "กลุ่มยุทธศาสตร์และแผนงาน","กลุ่มระบาดวิทยาและตอบโต้ภาวะฉุกเฉินทางสาธารณสุข",
    "กลุ่มโรคติดต่อ","กลุ่มโรคไม่ติดต่อ","กลุ่มโรคติดต่อเรื้อรัง","กลุ่มโรคติดต่อนำโดยแมลง",
    "กลุ่มโรคติดต่อนำโดยแมลง (ศตม. 9.1 จ.ชัยภูมิ)","กลุ่มโรคติดต่อนำโดยแมลง (ศตม. 9.2 จ.บุรีรัมย์)",
    "กลุ่มโรคติดต่อนำโดยแมลง (ศตม. 9.3 จ.สุรินทร์)","กลุ่มโรคติดต่อนำโดยแมลง (ศตม. 9.4 อ.ปากช่อง)",
    "กลุ่มโรคจากการประกอบอาชีพและสิ่งแวดล้อม","กลุ่มห้องปฏิบัติการทางการแพทย์ด้านควบคุมโรค",
    "กลุ่มสื่อสารความเสี่ยงโรคและภัยสุขภาพ","กลุ่มพัฒนานวัตกรรมและวิจัย","กลุ่มพัฒนาองค์กร",
    "ศูนย์ฝึกอบรมนักระบาดวิทยาภาคสนาม","ศูนย์บริการเวชศาสตร์ป้องกัน",
    "งานกฎหมาย","งานเภสัชกรรม","ด่านควบคุมโรคติดต่อระหว่างประเทศ","อื่นๆ",
]
LEAVE_TYPES: List[str] = list(LEAVE_QUOTA.keys())
COLUMN_MAPPING: Dict[str, str] = {"ชื่อพนักงาน": "ชื่อ-สกุล","ชื่อ": "ชื่อ-สกุล","fullname": "ชื่อ-สกุล"}
FILE_ATTEND="attendance_report.xlsx"; FILE_LEAVE="leave_report.xlsx"; FILE_TRAVEL="travel_report.xlsx"
FILE_STAFF="staff_master.xlsx"; FILE_NOTIFY="activity_log.xlsx"; FILE_HOLIDAYS="special_holidays.xlsx"; FILE_MANUAL_SCAN="manual_scan.xlsx"
FOLDER_ID="1YFJZvs59ahRHmlRrKcQwepWJz6A-4B7d"; ATTACHMENT_FOLDER_NAME="Attachments_Leave_App"; BACKUP_FOLDER_NAME="Backup"
STAFF_MASTER_COLS=["ชื่อ-สกุล","กลุ่มงาน","ตำแหน่ง","ประเภทบุคลากร","วันเริ่มงาน","สถานะ"]
MANUAL_SCAN_COLS=["ชื่อ-สกุล","วันที่","เวลาเข้า","เวลาออก","หมายเหตุ"]
ACTIVITY_LOG_COLS=["Timestamp","ประเภท","รายละเอียด","ผู้เกี่ยวข้อง"]
HOLIDAY_COLS=["วันที่","ชื่อวันหยุด","ประเภท","หมายเหตุ"]
TRAVEL_REQUIRED_COLS=["ชื่อ-สกุล","วันที่เริ่ม","วันที่สิ้นสุด","เรื่อง/กิจกรรม"]
_NON_TRAVEL_FILES={FILE_ATTEND,FILE_LEAVE,FILE_STAFF,FILE_NOTIFY,FILE_HOLIDAYS,FILE_MANUAL_SCAN}

# ===========================
# 🔒 Drive Thread-Safety
# ===========================
# httplib2 ไม่ Thread-safe — ใช้ thread-local แยก service ต่อ thread
# ป้องกัน "malloc: double linked list corrupted" จาก shared connection
_thread_local  = threading.local()
_DRIVE_LOCK    = threading.Lock()
_DRIVE_LOCK_TIMEOUT = 15

# Reconnect cooldown — ป้องกัน reconnect storm
_LAST_RECONNECT_TIME: float = 0.0
_RECONNECT_COOLDOWN = 10.0  # วินาที

# Circuit breaker global — ถ้า Drive down ชั่วคราว ไม่ loop ซ้ำ
_DRIVE_CIRCUIT_OPEN = False
_DRIVE_CIRCUIT_RESET_AT: float = 0.0
_DRIVE_CIRCUIT_TIMEOUT = 30.0  # เปิด circuit 30 วิ แล้วลองใหม่

# ===========================
# ☁️ Google Drive Service
# ===========================
def _build_drive_service():
    """สร้าง Drive service ใหม่ 1 ตัวต่อ 1 thread"""
    import httplib2
    import google_auth_httplib2

    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/drive"],
    )
    # google_auth_httplib2.AuthorizedHttp รองรับ google-auth ใหม่
    # (ไม่ใช้ creds.authorize() ซึ่งเป็น oauth2client เก่า)
    authorized_http = google_auth_httplib2.AuthorizedHttp(
        creds, http=httplib2.Http(timeout=20)
    )
    svc = build("drive", "v3", http=authorized_http, cache_discovery=False)
    logger.info("Drive connected (thread=%s)", threading.current_thread().name)
    return svc

def get_drive_service():
    """
    คืน Drive service แบบ thread-local
    - แต่ละ thread มี connection แยกกัน → ไม่ชนกัน
    - circuit breaker: fail 3 ครั้ง → error แทน crash
    """
    svc = getattr(_thread_local, "service", None)
    if svc is not None:
        return svc

    fail_count = getattr(_thread_local, "fail_count", 0)
    if fail_count >= 3:
        # main thread แสดง error ใน UI, background thread แค่ raise
        if threading.current_thread() is threading.main_thread():
            st.error("❌ เชื่อมต่อ Google Drive ไม่สำเร็จหลายครั้ง กรุณา Refresh หน้าเว็บ")
            st.stop()
        raise RuntimeError("Drive: circuit breaker open")

    try:
        _thread_local.service = _build_drive_service()
        _thread_local.fail_count = 0
        # reset circuit breaker เมื่อ connect สำเร็จ
        _DRIVE_CIRCUIT_OPEN = False
        return _thread_local.service
    except Exception as e:
        _thread_local.fail_count = fail_count + 1
        logger.error("Drive init failed (%d/3): %s", fail_count + 1, e)
        if threading.current_thread() is threading.main_thread():
            st.error(f"❌ เชื่อมต่อ Google Drive ไม่สำเร็จ: {e}")
            st.stop()
        raise

def _drop_drive_service() -> None:
    """ทิ้ง Drive service ของ thread นี้ — สร้างใหม่รอบต่อไป"""
    global _LAST_RECONNECT_TIME
    _thread_local.service = None

    # cooldown ป้องกัน reconnect storm
    now = time.time()
    if now - _LAST_RECONNECT_TIME < _RECONNECT_COOLDOWN:
        wait = _RECONNECT_COOLDOWN - (now - _LAST_RECONNECT_TIME)
        logger.warning("Drive reconnect cooldown: wait %.1fs", wait)
        time.sleep(wait)
    _LAST_RECONNECT_TIME = time.time()
    logger.warning("Drive service dropped — will reconnect on next call")

def _drive_execute(request, retries: int = 2):
    """
    Execute Drive API request พร้อม retry
    - ใช้ thread-local service (ไม่แชร์ข้าม thread)
    - Lock เฉพาะตอน reconnect ป้องกัน race condition
    - Circuit breaker: ถ้า Drive down ชั่วคราว ไม่ loop ซ้ำ
    """
    global _DRIVE_CIRCUIT_OPEN, _DRIVE_CIRCUIT_RESET_AT
    # ตรวจ circuit breaker
    if _DRIVE_CIRCUIT_OPEN:
        now = time.time()
        if now < _DRIVE_CIRCUIT_RESET_AT:
            raise RuntimeError(f"Drive circuit open — retry in {_DRIVE_CIRCUIT_RESET_AT - now:.0f}s")
        # ครบเวลาแล้ว → ลองเปิดใหม่
        _DRIVE_CIRCUIT_OPEN = False
        logger.info("Drive circuit breaker: half-open (trying again)")
    _TE = (
        BrokenPipeError, ConnectionResetError, ConnectionAbortedError,
        ConnectionRefusedError, OSError, ssl.SSLError, TimeoutError,
    )
    last_exc = None
    is_callable = callable(request)

    for attempt in range(retries):
        try:
            req = request() if is_callable else request
            return req.execute()
        except HttpError as e:
            status = e.resp.status if hasattr(e, "resp") else 0
            if status in (429, 500, 502, 503, 504):
                wait = (2 ** attempt) + 0.5
                logger.warning("Drive HTTP %d — retry %d/%d in %.1fs", status, attempt+1, retries, wait)
                time.sleep(wait)
                last_exc = e
                continue
            raise
        except _TE as e:
            logger.warning("Drive transport error (%s) — reconnect & retry %d/%d", type(e).__name__, attempt+1, retries)
            with _DRIVE_LOCK:          # lock เฉพาะ drop+reconnect
                _drop_drive_service()
            time.sleep(2 ** attempt)
            last_exc = e
            continue
        except Exception as e:
            if any(k in str(e).lower() for k in ("ssl", "record layer", "handshake", "eof")):
                logger.warning("Drive SSL error — reconnect & retry %d/%d: %s", attempt+1, retries, e)
                with _DRIVE_LOCK:
                    _drop_drive_service()
                time.sleep(2 ** attempt)
                last_exc = e
                continue
            raise

    # เปิด circuit breaker เมื่อ retry หมด
    _DRIVE_CIRCUIT_OPEN     = True
    _DRIVE_CIRCUIT_RESET_AT = time.time() + _DRIVE_CIRCUIT_TIMEOUT
    logger.error("Drive circuit opened — will reset in %.0fs", _DRIVE_CIRCUIT_TIMEOUT)
    raise last_exc or RuntimeError("Drive API: max retries exceeded")

def get_file_id(filename: str, parent_id: str = FOLDER_ID) -> Optional[str]:
    try:
        res = _drive_execute(lambda: get_drive_service().files().list(q=f"name='{filename}' and '{parent_id}' in parents and trashed=false", fields="files(id,modifiedTime)", orderBy="modifiedTime desc", supportsAllDrives=True, includeItemsFromAllDrives=True))
        files = res.get("files", [])
        if not files: return None
        keep_id = files[0]["id"]
        for dup in files[1:]:
            try: _drive_execute(lambda: get_drive_service().files().delete(fileId=dup["id"], supportsAllDrives=True))
            except Exception: pass
        return keep_id
    except Exception as e: logger.error(f"get_file_id({filename}): {e}"); return None

def get_or_create_folder(folder_name: str, parent_id: str) -> Optional[str]:
    try:
        svc = get_drive_service()
        res = _drive_execute(lambda: get_drive_service().files().list(q=f"name='{folder_name}' and '{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false", fields="files(id)", supportsAllDrives=True, includeItemsFromAllDrives=True))
        folders = res.get("files", [])
        if folders: return folders[0]["id"]
        new = _drive_execute(lambda: get_drive_service().files().create(body={"name":folder_name,"parents":[parent_id],"mimeType":"application/vnd.google-apps.folder"}, supportsAllDrives=True, fields="id"))
        return new.get("id")
    except Exception as e: logger.error(f"get_or_create_folder: {e}"); return None

@st.cache_data(ttl=900, show_spinner=False)
def _read_file_by_id(file_id: str) -> pd.DataFrame:
    try:
        svc = get_drive_service()
        req = svc.files().get_media(fileId=file_id, supportsAllDrives=True)
        fh = io.BytesIO(); dl = MediaIoBaseDownload(fh, req); done = False
        while not done: _, done = dl.next_chunk()
        fh.seek(0); return pd.read_excel(fh, engine="openpyxl")
    except Exception as e: logger.warning(f"_read_file_by_id({file_id}): {e}"); return pd.DataFrame()

@st.cache_data(ttl=900)
def read_excel_from_drive(filename: str) -> pd.DataFrame:
    fid = get_file_id(filename)
    if not fid: return pd.DataFrame()
    return _read_file_by_id(fid)

def read_excel_with_id(filename: str) -> Tuple[pd.DataFrame, Optional[str]]:
    fid = get_file_id(filename)
    if not fid: return pd.DataFrame(), None
    return _read_file_by_id(fid), fid

def read_excel_with_backup(filename: str, dedup_cols: Optional[List[str]] = None) -> Tuple[pd.DataFrame, Optional[str]]:
    frames: List[pd.DataFrame] = []
    df_main, main_fid = read_excel_with_id(filename)
    if not df_main.empty: df_main["_src"]="main"; frames.append(df_main)
    bak_name = f"BAK_{filename}"
    try:
        backup_root = get_or_create_folder(BACKUP_FOLDER_NAME, FOLDER_ID)
        if backup_root:
            bak_sub = get_or_create_folder(bak_name, backup_root)
            if bak_sub:
                bak_fid = get_file_id(bak_name, bak_sub)
                if bak_fid:
                    df_bak = _read_file_by_id(bak_fid)
                    if not df_bak.empty: df_bak["_src"]="backup"; frames.append(df_bak)
    except Exception as e: logger.warning(f"Backup read failed '{filename}': {e}")
    if not frames: return pd.DataFrame(), main_fid
    df_all = pd.concat(frames, ignore_index=True)
    if dedup_cols:
        df_all["_src_order"] = df_all["_src"].map({"main":0,"backup":1})
        df_all = df_all.sort_values("_src_order").drop_duplicates(subset=dedup_cols, keep="first").drop(columns=["_src_order"], errors="ignore")
    return df_all.drop(columns=["_src"], errors="ignore").reset_index(drop=True), main_fid

def write_excel_to_drive(filename: str, df: pd.DataFrame, known_file_id: Optional[str] = None) -> bool:
    try:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w: df.to_excel(w, index=False)
        buf.seek(0); media = MediaIoBaseUpload(buf, mimetype=EXCEL_MIME, resumable=False)
        svc = get_drive_service(); fid = known_file_id or get_file_id(filename)
        if fid: _drive_execute(lambda: get_drive_service().files().update(fileId=fid, media_body=media, supportsAllDrives=True))
        else: _drive_execute(lambda: get_drive_service().files().create(body={"name":filename,"parents":[FOLDER_ID]}, media_body=media, supportsAllDrives=True, fields="id"))
        # ⚡ ล้างเฉพาะ @st.cache_data ของไฟล์นั้น ไม่ล้างทั้งหมด
        read_excel_from_drive.clear(filename)
        _invalidate_cache()  # บังคับโหลด session cache ใหม่รอบต่อไป
        return True
    except Exception as e:
        logger.error("write_excel_to_drive(%s): %s", filename, e)
        st.error(f"บันทึกไฟล์ล้มเหลว: {e}")
        return False

def backup_excel(filename: str, df: pd.DataFrame) -> None:
    """
    สำรองไฟล์ — รันหลัง write เสร็จแล้ว (synchronous แต่ silent)
    ไม่ใช้ background thread เพราะ thread แยกใช้ httplib2 connection
    ร่วมกับ main thread ทำให้เกิด heap corruption
    """
    if df.empty:
        return
    try:
        fid = get_file_id(filename)
        if not fid:
            return
        bak_name    = f"BAK_{filename}"
        backup_root = get_or_create_folder(BACKUP_FOLDER_NAME, FOLDER_ID)
        if not backup_root:
            return
        bak_sub = get_or_create_folder(bak_name, backup_root)
        if not bak_sub:
            return
        existing = get_file_id(bak_name, bak_sub)
        if existing:
            try:
                _drive_execute(lambda: get_drive_service().files().delete(
                    fileId=existing, supportsAllDrives=True))
            except Exception:
                pass
        _drive_execute(lambda: get_drive_service().files().copy(
            fileId=fid,
            body={"name": bak_name, "parents": [bak_sub]},
            supportsAllDrives=True,
        ))
        logger.info("backup_excel: %s → BAK สำเร็จ", filename)
    except Exception as e:
        logger.warning("backup_excel(%s): %s", filename, e)

def upload_pdf_to_drive(uploaded_file, new_filename: str, folder_id: str) -> str:
    try:
        svc = get_drive_service(); meta = {"name":new_filename,"parents":[folder_id]}
        media = MediaIoBaseUpload(io.BytesIO(uploaded_file.getvalue()), mimetype="application/pdf", resumable=True)
        created = _drive_execute(lambda: get_drive_service().files().create(body=meta, media_body=media, supportsAllDrives=True, fields="id,webViewLink"))
        return created.get("webViewLink", "-")
    except Exception as e: logger.error(f"upload_pdf: {e}"); return "-"

@st.cache_data(ttl=900)
def list_all_files_in_folder(parent_id: str = FOLDER_ID) -> List[dict]:
    try:
        res = _drive_execute(lambda: get_drive_service().files().list(q=f"'{parent_id}' in parents and trashed=false and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'", fields="files(id,name,modifiedTime)", supportsAllDrives=True, includeItemsFromAllDrives=True, orderBy="modifiedTime desc"))
        return res.get("files", [])
    except Exception as e: logger.error(f"list_all_files: {e}"); return []

# ===========================
# 🛠️ Data Processing
# ===========================
def _normalize_name(val) -> str:
    if val is None: return ""
    s = str(val).strip()
    if s.lower() in ("nan","none",""): return ""
    return re.sub(r"\s+", " ", s)

def _normalize_date(val) -> Optional[dt.date]:
    if val is None: return None
    if isinstance(val, dt.datetime): return val.date()
    if isinstance(val, dt.date): return val
    try:
        s = str(val).strip()
        # ISO format (YYYY-MM-DD) → dayfirst=False ป้องกัน UserWarning
        if re.match(r'^\d{4}-\d{2}-\d{2}', s):
            ts = pd.to_datetime(s[:19], errors="coerce")
        else:
            ts = pd.to_datetime(s, dayfirst=True, errors="coerce")
        return None if pd.isna(ts) else ts.date()
    except Exception: return None

def _normalize_time_value(val) -> str:
    if val is None: return ""
    if isinstance(val, float):
        if math.isnan(val): return ""
        total_sec = int(round(val * 86400)); h, m = (total_sec//3600)%24, (total_sec%3600)//60
        return f"{h:02d}:{m:02d}"
    if isinstance(val, (pd.Timedelta, dt.timedelta)):
        total_sec = int(val.total_seconds())
        if total_sec < 0: return ""
        h, m = (total_sec//3600)%24, (total_sec%3600)//60; return f"{h:02d}:{m:02d}"
    if isinstance(val, (dt.datetime, dt.time)): return val.strftime("%H:%M")
    s = str(val).strip()
    if not s or s.lower() in ("nan","none","nat",""): return ""
    m_re = re.search(r"(\d+):(\d{2})(?::(\d{2}))?(?:\s*(AM|PM))?", s, re.IGNORECASE)
    if m_re:
        h, mn = int(m_re.group(1)), int(m_re.group(2))
        meridiem = (m_re.group(4) or "").upper()
        d_m = re.search(r"(\d+)\s+day", s, re.IGNORECASE)
        if d_m: h += int(d_m.group(1)) * 24
        if meridiem == "PM" and h < 12: h += 12
        elif meridiem == "AM" and h == 12: h = 0
        return f"{h%24:02d}:{mn:02d}"
    return ""

def _parse_date_flex(val) -> Optional[pd.Timestamp]:
    if val is None or (isinstance(val, float) and pd.isna(val)): return pd.NaT
    if isinstance(val, pd.Timestamp): return val
    if isinstance(val, (dt.datetime, dt.date)): return pd.Timestamp(val)
    if isinstance(val, (int, float)):
        try: return pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(val))
        except Exception: return pd.NaT
    val_str = str(val).strip()
    if not val_str or val_str.lower() in ("nat","nan","none",""): return pd.NaT
    if re.match(r"^\d{4}-\d{2}-\d{2}", val_str):
        try: return pd.Timestamp(val_str[:19])
        except Exception: pass
    val_clean = re.sub(r"(\d{1,2})\.(\d{2})(\s*$)", r"\1:\2", val_str)
    parts = val_clean.split(" ")[0].split("T")[0].split("/")
    if len(parts) != 3:
        try: return pd.to_datetime(val_str, dayfirst=True, errors="coerce")
        except Exception: return pd.NaT
    try: a, b, c = int(parts[0]), int(parts[1]), int(parts[2])
    except ValueError: return pd.NaT
    year = c-543 if c>2400 else (2000+c if c<100 else c)
    if a>12: day,month=a,b
    elif b>12: month,day=a,b
    elif c>2400: day,month=a,b
    else: month,day=a,b
    try: return pd.Timestamp(year=year, month=month, day=day)
    except Exception:
        try: return pd.Timestamp(year=year, month=day, day=month)
        except Exception: return pd.NaT

def normalize_date_col(df: pd.DataFrame, col: str) -> pd.DataFrame:
    if df.empty or col not in df.columns: return df
    series = df[col]
    if pd.api.types.is_datetime64_any_dtype(series): df[col]=series.dt.normalize(); return df
    df[col] = series.apply(_parse_date_flex)
    df[col] = pd.to_datetime(df[col], errors="coerce").dt.normalize()
    return df

def clean_names(df: pd.DataFrame, col: str) -> pd.DataFrame:
    if df.empty or col not in df.columns: return df
    if df.columns.duplicated().any(): df = df.loc[:,~df.columns.duplicated()].copy()
    series = df[col]
    if isinstance(series, pd.DataFrame): series = series.iloc[:,0]
    df[col] = series.astype(str).str.strip().str.replace(r"\s+"," ",regex=True)
    return df

def preprocess_dataframes(df_leave, df_travel, df_att):
    if not df_att.empty:
        for old,new in COLUMN_MAPPING.items():
            if old in df_att.columns:
                if new in df_att.columns: df_att=df_att.drop(columns=[new])
                df_att=df_att.rename(columns={old:new})
        if df_att.columns.duplicated().any(): df_att=df_att.loc[:,~df_att.columns.duplicated()].copy()
    for col in ["วันที่เริ่ม","วันที่สิ้นสุด"]:
        df_leave=normalize_date_col(df_leave,col); df_travel=normalize_date_col(df_travel,col)
    df_att=normalize_date_col(df_att,"วันที่")
    df_leave=clean_names(df_leave,"ชื่อ-สกุล"); df_travel=clean_names(df_travel,"ชื่อ-สกุล"); df_att=clean_names(df_att,"ชื่อ-สกุล")
    return df_leave, df_travel, df_att

def count_weekdays(start_date, end_date, extra_holidays: Optional[List[dt.date]] = None) -> int:
    if not start_date or not end_date: return 0
    if isinstance(start_date, dt.datetime): start_date=start_date.date()
    if isinstance(end_date, dt.datetime): end_date=end_date.date()
    base = int(np.busday_count(start_date, end_date+dt.timedelta(days=1)))
    if extra_holidays:
        overlap = sum(1 for h in extra_holidays if start_date<=h<=end_date and h.weekday()<5)
        base = max(0, base-overlap)
    return base

def parse_time(val) -> Optional[dt.time]:
    if val is None or val == "": return None
    if isinstance(val, float):
        if np.isnan(val): return None
        try:
            total_sec = int(round(val*86400))
            return dt.time(total_sec//3600, (total_sec%3600)//60, total_sec%60)
        except Exception: return None
    if isinstance(val, dt.time): return val
    if isinstance(val, dt.datetime): return val.time()
    if isinstance(val, (pd.Timedelta, dt.timedelta)):
        try:
            total_sec = int(val.total_seconds())
            if total_sec < 0: return None
            return dt.time(total_sec//3600%24, (total_sec%3600)//60, total_sec%60)
        except Exception: return None
    s = str(val).strip()
    if not s or s.lower() in ("nat","none","nan",""): return None
    m = re.search(r"(\d+):(\d{2}):?(\d{2})?(?:\s*(AM|PM))?", s, re.IGNORECASE)
    if m:
        h,mn,sc = int(m.group(1)),int(m.group(2)),int(m.group(3)) if m.group(3) else 0
        meridiem = (m.group(4) or "").upper()
        d_match = re.search(r"(\d+)\s+day", s, re.IGNORECASE)
        if d_match: h += int(d_match.group(1))*24
        if meridiem=="PM" and h<12: h+=12
        elif meridiem=="AM" and h==12: h=0
        try: return dt.time(h%24, mn, sc)
        except Exception: pass
    try: return pd.to_datetime(s).time()
    except Exception: return None

# ===========================
# 📅 Attendance Report
# ===========================
@st.cache_data(ttl=900)
def read_attendance_report() -> pd.DataFrame:
    """
    อ่านไฟล์ attendance_report.xlsx อย่างละเอียด รองรับหลายรูปแบบ:

    รูปแบบ A — แต่ละแถวคือ 1 การสแกน (ชื่อ | วันที่ | เวลาเข้า | เวลาออก)
    รูปแบบ B — แต่ละแถวมีชื่อซ้ำหลายวัน (ชื่อ | วันที่ | เวลา | เวลา)
    รูปแบบ C — ไฟล์เครื่องสแกน ZKTeco/Fingertec: No | ชื่อ | Department | Date | Time | ...
    รูปแบบ D — ชื่อ column ภาษาอังกฤษ: Name/Employee | Date | Check In | Check Out
    """
    fid = get_file_id(FILE_ATTEND)
    if not fid:
        logger.warning("read_attendance_report: ไม่พบไฟล์ %s ใน Drive", FILE_ATTEND)
        return pd.DataFrame()

    try:
        req  = get_drive_service().files().get_media(fileId=fid, supportsAllDrives=True)
        fh   = io.BytesIO()
        dl   = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = dl.next_chunk()
        fh.seek(0)
        # อ่าน dtype=str ทั้งหมดเพื่อป้องกัน pandas auto-cast วันที่/เวลาผิด
        df_raw = pd.read_excel(fh, engine="openpyxl", header=0, dtype=str)
    except Exception as e:
        logger.error("read_attendance_report: %s", e)
        return pd.DataFrame()

    if df_raw.empty:
        logger.warning("read_attendance_report: ไฟล์ว่างเปล่า")
        return pd.DataFrame()

    # ── normalize column names ──────────────────────────────────────────
    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    raw_cols = df_raw.columns.tolist()
    logger.info("read_attendance_report: columns = %s", raw_cols)

    # ── fuzzy column matching ───────────────────────────────────────────
    # ชื่อพนักงาน
    NAME_CANDIDATES = [
        "ชื่อ-สกุล","ชื่อพนักงาน","ชื่อ","Name","Employee Name",
        "employee","name","fullname","FullName","EMPLOYEE","NAME",
        "ชื่อ - สกุล","ชื่อ-นามสกุล",
    ]
    # วันที่
    DATE_CANDIDATES = [
        "วันที่","date","Date","DATE","วันที่เข้างาน","Check Date",
        "checkdate","AttendDate","วัน/เดือน/ปี","Attendance Date",
    ]
    # เวลาเข้า
    IN_CANDIDATES = [
        "เวลาเข้า","เข้า","check_in","Check In","CheckIn","checkin",
        "เวลาเข้างาน","Time In","time_in","IN","In","เข้างาน",
        "First Check","First In","Scan In",
    ]
    # เวลาออก
    OUT_CANDIDATES = [
        "เวลาออก","ออก","check_out","Check Out","CheckOut","checkout",
        "เวลาออกงาน","Time Out","time_out","OUT","Out","ออกงาน",
        "Last Check","Last Out","Scan Out",
    ]
    # หมายเหตุ
    NOTE_CANDIDATES = ["หมายเหตุ","note","Note","NOTE","Remark","remark","REMARK"]

    def _find_col(candidates: list[str]) -> Optional[str]:
        """ค้นหา column จาก candidates list (exact → lower → contains)"""
        # exact match
        for c in candidates:
            if c in raw_cols:
                return c
        # case-insensitive
        raw_lower = {col.lower(): col for col in raw_cols}
        for c in candidates:
            if c.lower() in raw_lower:
                return raw_lower[c.lower()]
        # contains match (สำหรับชื่อ column ยาว เช่น "เวลาเข้างาน (HH:MM)")
        for c in candidates:
            for col in raw_cols:
                if c.lower() in col.lower():
                    return col
        return None

    COL_NAME = _find_col(NAME_CANDIDATES)
    COL_DATE = _find_col(DATE_CANDIDATES)
    COL_IN   = _find_col(IN_CANDIDATES)
    COL_OUT  = _find_col(OUT_CANDIDATES)
    COL_NOTE = _find_col(NOTE_CANDIDATES)

    logger.info(
        "read_attendance_report: mapping — ชื่อ=%s วันที่=%s เข้า=%s ออก=%s หมายเหตุ=%s",
        COL_NAME, COL_DATE, COL_IN, COL_OUT, COL_NOTE,
    )

    # ── ถ้าหา column หลักไม่เจอ ให้ลอง detect แบบ positional ──────────
    # บางไฟล์เครื่องสแกนมี header แปลก เช่น แถวแรกไม่ใช่ header จริง
    if COL_DATE is None or COL_NAME is None:
        logger.warning("read_attendance_report: ไม่พบ column มาตรฐาน — ลอง multi-header scan")
        # ลองอ่านซ้ำโดยข้าม 1-3 แถวแรก
        for skip in range(1, 5):
            try:
                fh.seek(0)
                df_try = pd.read_excel(fh, engine="openpyxl", header=skip, dtype=str)
                df_try.columns = [str(c).strip() for c in df_try.columns]
                if _find_col(DATE_CANDIDATES) or _find_col(NAME_CANDIDATES):
                    df_raw  = df_try
                    raw_cols = df_raw.columns.tolist()
                    COL_NAME = _find_col(NAME_CANDIDATES)
                    COL_DATE = _find_col(DATE_CANDIDATES)
                    COL_IN   = _find_col(IN_CANDIDATES)
                    COL_OUT  = _find_col(OUT_CANDIDATES)
                    COL_NOTE = _find_col(NOTE_CANDIDATES)
                    logger.info("read_attendance_report: ใช้ header row=%d → %s", skip, raw_cols[:6])
                    break
            except Exception:
                continue

    if COL_DATE is None:
        logger.error(
            "read_attendance_report: ไม่พบ column วันที่เลย (columns=%s)", raw_cols
        )
        return pd.DataFrame()

    # ── กรณีไม่มี column ชื่อ — ลองดู column แรกหรือ column ที่มีชื่อบุคคล ──
    if COL_NAME is None:
        # ลองหา column ที่ค่าเริ่มต้นด้วยคำนำหน้าชื่อ
        prefix_re = re.compile(r"^(นาย|นาง(?:สาว)?|Mr|Mrs|Ms|Miss)", re.IGNORECASE)
        for col in raw_cols:
            sample = df_raw[col].dropna().astype(str).head(20)
            if sample.str.match(prefix_re).sum() >= 3:
                COL_NAME = col
                logger.info("read_attendance_report: detect ชื่อจาก value pattern → '%s'", col)
                break
        if COL_NAME is None and raw_cols:
            COL_NAME = raw_cols[0]  # fallback: column แรก
            logger.warning("read_attendance_report: ใช้ column แรก '%s' เป็นชื่อ", COL_NAME)

    # ── build output rows ────────────────────────────────────────────────
    rows_out = []
    skipped  = 0

    for idx, row in df_raw.iterrows():
        # ชื่อ
        name = _normalize_name(row.get(COL_NAME, "")) if COL_NAME else ""
        if not name:
            skipped += 1
            continue

        # วันที่ — ลอง _normalize_date ก่อน แล้ว fallback _parse_date_flex
        raw_date = row.get(COL_DATE, "")
        date_val = _normalize_date(raw_date)
        if date_val is None:
            ts = _parse_date_flex(raw_date)
            date_val = ts.date() if ts is not None and not pd.isna(ts) else None
        if date_val is None:
            skipped += 1
            continue

        # เวลา
        time_in  = _normalize_time_value(row.get(COL_IN,  "")) if COL_IN  else ""
        time_out = _normalize_time_value(row.get(COL_OUT, "")) if COL_OUT else ""

        # กรณีเวลาเข้า=ออก เหมือนกัน (เครื่องสแกนบางรุ่น record ครั้งเดียว)
        # ไม่ต้องแก้ไขที่นี่ — logic ใน _att_status จะจัดการเอง

        note = str(row.get(COL_NOTE, "") or "").strip() if COL_NOTE else ""

        rows_out.append({
            "ชื่อ-สกุล": name,
            "วันที่":     pd.Timestamp(date_val),
            "เวลาเข้า":   time_in,
            "เวลาออก":    time_out,
            "หมายเหตุ":   note,
        })

    logger.info(
        "read_attendance_report: อ่านได้ %d แถว, ข้าม %d แถว (ชื่อ/วันที่ว่าง)",
        len(rows_out), skipped,
    )

    if not rows_out:
        return pd.DataFrame(columns=["ชื่อ-สกุล","วันที่","เวลาเข้า","เวลาออก","หมายเหตุ","เดือน"])

    df_out = pd.DataFrame(rows_out)
    df_out["วันที่"] = pd.to_datetime(df_out["วันที่"], errors="coerce").dt.normalize()
    df_out["เดือน"]  = df_out["วันที่"].dt.strftime("%Y-%m")
    df_out = df_out.dropna(subset=["วันที่"])
    df_out = df_out[df_out["ชื่อ-สกุล"] != ""].reset_index(drop=True)

    # ── dedup: ถ้า 1 คน 1 วัน มีหลายแถว ให้เอาเวลาเข้าแรกสุด + ออกหลังสุด ──
    # (เครื่องบางรุ่น record ทุกครั้งที่แตะ)
    df_out["_time_in_dt"]  = df_out["เวลาเข้า"].apply(parse_time)
    df_out["_time_out_dt"] = df_out["เวลาออก"].apply(parse_time)

    def _agg_scans(grp: pd.DataFrame) -> pd.Series:
        times_in  = grp["_time_in_dt"].dropna().tolist()
        times_out = grp["_time_out_dt"].dropna().tolist()
        t_in_str  = min(times_in).strftime("%H:%M")  if times_in  else ""
        t_out_str = max(times_out).strftime("%H:%M") if times_out else ""
        note_combined = " | ".join(filter(None, grp["หมายเหตุ"].unique().tolist()))
        return pd.Series({
            "เวลาเข้า": t_in_str,
            "เวลาออก":  t_out_str,
            "หมายเหตุ": note_combined,
            "เดือน":    grp["เดือน"].iloc[0],
        })

    n_before = len(df_out)
    df_out = (
        df_out
        .groupby(["ชื่อ-สกุล", "วันที่"], as_index=False)
        .apply(_agg_scans)
        .reset_index(drop=True)
    )
    n_after = len(df_out)
    if n_before != n_after:
        logger.info(
            "read_attendance_report: รวม multi-scan %d → %d แถว (dedup)",
            n_before, n_after,
        )

    df_out = df_out.sort_values(["ชื่อ-สกุล","วันที่"]).reset_index(drop=True)
    return df_out

# ===========================
# 🚗 Travel Data
# ===========================
def _extract_travel_from_activity_log(df_log: pd.DataFrame) -> pd.DataFrame:
    if df_log.empty or "ประเภท" not in df_log.columns: return pd.DataFrame()
    df_tr = df_log[df_log["ประเภท"]=="ไปราชการ"].copy()
    if df_tr.empty: return pd.DataFrame()
    rows=[]
    for _,row in df_tr.iterrows():
        try:
            ts=pd.to_datetime(row.get("Timestamp"),errors="coerce")
            if pd.isna(ts): continue
            d=ts.normalize(); detail=str(row.get("รายละเอียด",""))
            project=detail.split("@")[0].strip() if "@" in detail else detail
            names=[n.strip() for n in str(row.get("ผู้เกี่ยวข้อง","")).split(",") if n.strip() and "และอีก" not in n]
            for name in names: rows.append({"ชื่อ-สกุล":name,"วันที่เริ่ม":d,"วันที่สิ้นสุด":d,"เรื่อง/กิจกรรม":project})
        except Exception: continue
    return pd.DataFrame(rows) if rows else pd.DataFrame()

@st.cache_data(ttl=900)
def load_all_travel() -> pd.DataFrame:
    frames: List[pd.DataFrame]=[]
    for f in list_all_files_in_folder():
        fname=f.get("name","")
        if fname in _NON_TRAVEL_FILES or fname.startswith("BAK_"): continue
        try:
            df_raw=read_excel_from_drive(fname)
            if df_raw.empty: continue
            has_name=any(c in df_raw.columns for c in ["ชื่อ-สกุล","ชื่อพนักงาน","ชื่อ"])
            if not (has_name and "วันที่เริ่ม" in df_raw.columns and "วันที่สิ้นสุด" in df_raw.columns): continue
            df_norm=df_raw.copy()
            for alt in ["ชื่อพนักงาน","ชื่อ","fullname"]:
                if alt in df_norm.columns and "ชื่อ-สกุล" not in df_norm.columns: df_norm.rename(columns={alt:"ชื่อ-สกุล"},inplace=True)
            df_norm["วันที่เริ่ม"]=pd.to_datetime(df_norm["วันที่เริ่ม"],errors="coerce").dt.normalize()
            df_norm["วันที่สิ้นสุด"]=pd.to_datetime(df_norm["วันที่สิ้นสุด"],errors="coerce").dt.normalize()
            if fname==FILE_NOTIFY or ("ประเภท" in df_norm.columns and "รายละเอียด" in df_norm.columns):
                df_tl=_extract_travel_from_activity_log(df_norm)
                if not df_tl.empty: df_tl["_source_file"]=fname; frames.append(df_tl)
                continue
            df_norm["_source_file"]=fname
            if "เรื่อง/กิจกรรม" not in df_norm.columns: df_norm["เรื่อง/กิจกรรม"]=fname.replace(".xlsx","")
            df_norm=df_norm.dropna(subset=["ชื่อ-สกุล","วันที่เริ่ม","วันที่สิ้นสุด"])
            df_norm["ชื่อ-สกุล"]=df_norm["ชื่อ-สกุล"].astype(str).str.strip()
            df_norm=df_norm[df_norm["ชื่อ-สกุล"].str.lower()!="nan"]
            if not df_norm.empty: frames.append(df_norm[TRAVEL_REQUIRED_COLS+["_source_file"]])
        except Exception as e: logger.warning(f"load_all_travel skip {fname}: {e}")
    try:
        backup_root=get_or_create_folder(BACKUP_FOLDER_NAME,FOLDER_ID)
        if backup_root:
            bak_sub=get_or_create_folder(f"BAK_{FILE_TRAVEL}",backup_root)
            bak_fid=get_file_id(f"BAK_{FILE_TRAVEL}",bak_sub) if bak_sub else None
            if bak_fid:
                df_bak=_read_file_by_id(bak_fid)
                if not df_bak.empty:
                    for alt in ["ชื่อพนักงาน","ชื่อ"]:
                        if alt in df_bak.columns and "ชื่อ-สกุล" not in df_bak.columns: df_bak.rename(columns={alt:"ชื่อ-สกุล"},inplace=True)
                    df_bak["วันที่เริ่ม"]=pd.to_datetime(df_bak.get("วันที่เริ่ม"),errors="coerce").dt.normalize()
                    df_bak["วันที่สิ้นสุด"]=pd.to_datetime(df_bak.get("วันที่สิ้นสุด"),errors="coerce").dt.normalize()
                    if "เรื่อง/กิจกรรม" not in df_bak.columns: df_bak["เรื่อง/กิจกรรม"]="ไปราชการ"
                    df_bak=df_bak.dropna(subset=["ชื่อ-สกุล","วันที่เริ่ม","วันที่สิ้นสุด"])
                    df_bak["ชื่อ-สกุล"]=df_bak["ชื่อ-สกุล"].astype(str).str.strip()
                    df_bak=df_bak[df_bak["ชื่อ-สกุล"].str.lower()!="nan"]
                    df_bak["_source_file"]=f"[Backup] {FILE_TRAVEL}"
                    valid_cols=[c for c in TRAVEL_REQUIRED_COLS+["_source_file"] if c in df_bak.columns]
                    if not df_bak.empty: frames.append(df_bak[valid_cols])
    except Exception as e: logger.warning(f"Backup travel read failed: {e}")
    if not frames: return pd.DataFrame(columns=TRAVEL_REQUIRED_COLS+["_source_file"])
    df_all=pd.concat(frames,ignore_index=True)
    df_all["_rank"]=df_all["_source_file"].apply(lambda s:0 if s==FILE_TRAVEL else(1 if s.startswith("[Backup]") else 2))
    return df_all.sort_values(["ชื่อ-สกุล","วันที่เริ่ม","_rank"]).drop_duplicates(subset=["ชื่อ-สกุล","วันที่เริ่ม","วันที่สิ้นสุด"],keep="first").drop(columns=["_rank"]).reset_index(drop=True)

# ===========================
# 🏖️ Holidays
# ===========================
FIXED_THAI_HOLIDAYS: List[Tuple[int,int,str]] = [
    (1,1,"วันขึ้นปีใหม่"),(4,6,"วันจักรี"),(4,13,"วันสงกรานต์"),(4,14,"วันสงกรานต์"),(4,15,"วันสงกรานต์"),
    (5,1,"วันแรงงานแห่งชาติ"),(5,5,"วันฉัตรมงคล"),(6,3,"วันเฉลิมพระชนมพรรษา สมเด็จพระราชินี"),
    (7,28,"วันเฉลิมพระชนมพรรษา ร.10"),(8,12,"วันแม่แห่งชาติ"),(10,13,"วันคล้ายวันสวรรคต ร.9"),
    (10,23,"วันปิยมหาราช"),(12,5,"วันพ่อแห่งชาติ / วันชาติ"),(12,10,"วันรัฐธรรมนูญ"),(12,31,"วันสิ้นปี"),
]

def get_fixed_holidays_for_year(year: int) -> pd.DataFrame:
    rows=[]
    for month,day,name in FIXED_THAI_HOLIDAYS:
        try: rows.append({"วันที่":pd.Timestamp(dt.date(year,month,day)),"ชื่อวันหยุด":name,"ประเภท":"วันหยุดราชการ","หมายเหตุ":"กำหนดโดยระบบ"})
        except ValueError: pass
    return pd.DataFrame(rows)

@st.cache_data(ttl=900)
def load_holidays_raw() -> pd.DataFrame:
    df=read_excel_from_drive(FILE_HOLIDAYS)
    if not df.empty:
        df["วันที่"]=pd.to_datetime(df["วันที่"],errors="coerce"); df=df.dropna(subset=["วันที่"])
        for col in HOLIDAY_COLS:
            if col not in df.columns: df[col]=""
    return df

def load_holidays_with_id() -> Tuple[pd.DataFrame, Optional[str]]:
    df,fid=read_excel_with_backup(FILE_HOLIDAYS,dedup_cols=["วันที่","ชื่อวันหยุด"])
    if not df.empty:
        df["วันที่"]=pd.to_datetime(df["วันที่"],errors="coerce"); df=df.dropna(subset=["วันที่"])
        for col in HOLIDAY_COLS:
            if col not in df.columns: df[col]=""
    return df,fid

def load_holidays_all(year: Optional[int]=None) -> pd.DataFrame:
    df_custom=load_holidays_raw(); frames=[]
    if year: frames.append(get_fixed_holidays_for_year(year))
    if not df_custom.empty: frames.append(df_custom[df_custom["วันที่"].dt.year==year] if year else df_custom)
    if not frames: return pd.DataFrame(columns=HOLIDAY_COLS)
    return pd.concat(frames,ignore_index=True).drop_duplicates(subset=["วันที่"]).sort_values("วันที่").reset_index(drop=True)

def get_holiday_dates(year: Optional[int]=None) -> List[dt.date]:
    df_h=load_holidays_all(year)
    if df_h.empty: return []
    return pd.to_datetime(df_h["วันที่"],errors="coerce").dropna().dt.date.tolist()

# ===========================
# 👥 Staff & Scan
# ===========================
def get_active_staff(df_staff: pd.DataFrame) -> List[str]:
    if df_staff.empty or "ชื่อ-สกุล" not in df_staff.columns: return []
    df_active=df_staff[df_staff["สถานะ"]=="ปฏิบัติงาน"] if "สถานะ" in df_staff.columns else df_staff
    return sorted(df_active["ชื่อ-สกุล"].dropna().astype(str).str.strip().unique().tolist())

def get_all_names_fallback(df_leave,df_travel,df_att) -> List[str]:
    all_names=set()
    for df in [df_leave,df_travel,df_att]:
        if not df.empty and "ชื่อ-สกุล" in df.columns: all_names.update(df["ชื่อ-สกุล"].dropna().astype(str).str.strip().unique())
    return sorted([n for n in all_names if n.lower()!="nan"])

def _parse_manual_scan_detail(detail: str, person: str) -> Optional[dict]:
    m=re.search(r"(\d{4}-\d{2}-\d{2})\s+เข้า\s+(\d{1,2}:\d{2})\s+ออก\s+(\d{1,2}:\d{2})",str(detail))
    if not m:
        m2=re.search(r"(\d{1,2}/\d{1,2}/\d{4})\s+เข้า\s+(\d{1,2}:\d{2})\s+ออก\s+(\d{1,2}:\d{2})",str(detail))
        if not m2: return None
        date_str,t_in,t_out=m2.group(1),m2.group(2),m2.group(3)
        try: d=pd.to_datetime(date_str,dayfirst=True).normalize()
        except Exception: return None
    else:
        date_str,t_in,t_out=m.group(1),m.group(2),m.group(3)
        try: d=pd.to_datetime(date_str).normalize()
        except Exception: return None
    return {"ชื่อ-สกุล":person.strip(),"วันที่":d,"เวลาเข้า":t_in,"เวลาออก":t_out,"หมายเหตุ":f"Activity Log — {detail[:60]}"}

def _parse_delete_scan_detail(detail: str, person: str) -> Optional[tuple]:
    m=re.search(r"(\d{1,2}/\d{1,2}/\d{4})",str(detail))
    if not m:
        m2=re.search(r"(\d{4}-\d{2}-\d{2})",str(detail))
        if not m2: return None
        try: d=pd.to_datetime(m2.group(1)).normalize()
        except Exception: return None
    else:
        try: d=pd.to_datetime(m.group(1),dayfirst=True).normalize()
        except Exception: return None
    return (person.strip(),d)

@st.cache_data(ttl=900)
def load_manual_scans() -> pd.DataFrame:
    frames: List[pd.DataFrame]=[]
    df_ms=read_excel_from_drive(FILE_MANUAL_SCAN)
    if not df_ms.empty:
        df_ms["วันที่"]=pd.to_datetime(df_ms["วันที่"],errors="coerce").dt.normalize()
        df_ms["ชื่อ-สกุล"]=df_ms["ชื่อ-สกุล"].astype(str).str.strip()
        for col in MANUAL_SCAN_COLS:
            if col not in df_ms.columns: df_ms[col]=""
        frames.append(df_ms[MANUAL_SCAN_COLS])
    df_log=read_excel_from_drive(FILE_NOTIFY)
    if not df_log.empty and "ประเภท" in df_log.columns:
        deleted_keys=set()
        for _,row in df_log[df_log["ประเภท"].astype(str).str.strip()=="ลบสแกนนิ้ว"].iterrows():
            result=_parse_delete_scan_detail(str(row.get("รายละเอียด","")),str(row.get("ผู้เกี่ยวข้อง","")))
            if result: deleted_keys.add(f"{result[0]}|{result[1]}")
        log_rows=[]
        for _,row in df_log[df_log["ประเภท"].astype(str).str.strip()=="คีย์สแกนนิ้ว"].iterrows():
            rec=_parse_manual_scan_detail(str(row.get("รายละเอียด","")),str(row.get("ผู้เกี่ยวข้อง","")))
            if rec and f"{rec['ชื่อ-สกุล']}|{rec['วันที่']}" not in deleted_keys: log_rows.append(rec)
        if log_rows:
            df_fl=pd.DataFrame(log_rows)
            for col in MANUAL_SCAN_COLS:
                if col not in df_fl.columns: df_fl[col]=""
            frames.append(df_fl[MANUAL_SCAN_COLS])
    if not frames: return pd.DataFrame(columns=MANUAL_SCAN_COLS)
    df_all=pd.concat(frames,ignore_index=True)
    df_all["วันที่"]=pd.to_datetime(df_all["วันที่"],errors="coerce").dt.normalize()
    df_all["ชื่อ-สกุล"]=df_all["ชื่อ-สกุล"].astype(str).str.strip()
    df_all=df_all.dropna(subset=["วันที่"]); df_all=df_all[df_all["ชื่อ-สกุล"].str.lower()!="nan"]
    return df_all.drop_duplicates(subset=["ชื่อ-สกุล","วันที่"],keep="first").sort_values(["ชื่อ-สกุล","วันที่"]).reset_index(drop=True)

def merge_attendance_with_manual(df_att: pd.DataFrame, df_manual: pd.DataFrame) -> pd.DataFrame:
    if df_manual.empty: return df_att
    if df_att.empty: df_manual_out=df_manual.copy(); df_manual_out["_source"]="manual"; return df_manual_out
    df_att_work=df_att.copy(); df_manual_work=df_manual.copy()
    att_name_col=next((c for c in ["ชื่อ-สกุล","ชื่อพนักงาน","ชื่อ"] if c in df_att_work.columns),None)
    if att_name_col is None: return df_att_work
    if att_name_col!="ชื่อ-สกุล": df_att_work=df_att_work.rename(columns={att_name_col:"ชื่อ-สกุล"})
    df_att_work["วันที่"]=pd.to_datetime(df_att_work["วันที่"],errors="coerce").dt.normalize()
    df_manual_work["วันที่"]=pd.to_datetime(df_manual_work["วันที่"],errors="coerce").dt.normalize()
    df_att_work["ชื่อ-สกุล"]=df_att_work["ชื่อ-สกุล"].astype(str).str.strip().str.replace(r"\s+"," ",regex=True)
    df_manual_work["ชื่อ-สกุล"]=df_manual_work["ชื่อ-สกุล"].astype(str).str.strip().str.replace(r"\s+"," ",regex=True)
    att_keys=set(df_att_work["ชื่อ-สกุล"].astype(str)+"|"+df_att_work["วันที่"].astype(str))
    df_manual_new=df_manual_work[~(df_manual_work["ชื่อ-สกุล"].astype(str)+"|"+df_manual_work["วันที่"].astype(str)).isin(att_keys)].copy()
    if df_manual_new.empty: return df_att_work
    df_manual_new["_source"]="manual"; df_att_work["_source"]="scan"
    return pd.concat([df_att_work,df_manual_new],ignore_index=True).sort_values(["ชื่อ-สกุล","วันที่"]).reset_index(drop=True)

# ===========================
# 🔔 Notifications & Audit
# ===========================
def send_line_notify(message: str) -> bool:
    token=st.secrets.get("line_notify_token","")
    if not token: return False
    try:
        resp=requests.post("https://notify-api.line.me/api/notify",headers={"Authorization":f"Bearer {token}"},data={"message":message},timeout=5)
        return resp.status_code==200
    except Exception as e: logger.warning(f"LINE Notify failed: {e}"); return False

def format_leave_notify(record: dict) -> str:
    start_dt=record.get('วันที่เริ่ม',''); end_dt=record.get('วันที่สิ้นสุด','')
    s=start_dt.strftime('%d/%m/%Y') if hasattr(start_dt,'strftime') else str(start_dt)
    e=end_dt.strftime('%d/%m/%Y') if hasattr(end_dt,'strftime') else str(end_dt)
    return f"\n🔔 แจ้งการลา — สคร.9\n👤 {record.get('ชื่อ-สกุล','')} ({record.get('กลุ่มงาน','')})\n📌 {record.get('ประเภทการลา','')} {record.get('จำนวนวันลา','')} วัน\n📅 {s} ถึง {e}\n📝 {record.get('เหตุผล','')}"

def format_travel_notify(persons: List[str], project: str, location: str, d_start, d_end, days: int) -> str:
    names_str=", ".join(persons[:5])+(f" และอีก {len(persons)-5} คน" if len(persons)>5 else "")
    return f"\n✈️ แจ้งไปราชการ — สคร.9\n👥 {names_str}\n📌 {project}\n📍 {location}\n📅 {d_start.strftime('%d/%m/%Y')} ถึง {d_end.strftime('%d/%m/%Y')} ({days} วันทำการ)"

def log_activity(action_type: str, detail: str, persons: str = "") -> None:
    try:
        df_log,_notify_fid=read_excel_with_backup(FILE_NOTIFY)
        if df_log.empty: df_log=pd.DataFrame(columns=ACTIVITY_LOG_COLS)
        new_row={"Timestamp":dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"ประเภท":action_type,"รายละเอียด":str(detail).replace("\n"," ")[:500],"ผู้เกี่ยวข้อง":persons}
        df_log=pd.concat([df_log,pd.DataFrame([new_row])],ignore_index=True).tail(500).reset_index(drop=True)
        write_excel_to_drive(FILE_NOTIFY,df_log,known_file_id=_notify_fid)
    except Exception as e: logger.warning(f"log_activity failed: {e}")

# ===========================
# ✅ Validation & Quota
# ===========================
def validate_leave_data(name,start_date,end_date,reason,df_leave) -> List[str]:
    errors=[]
    if not name or not name.strip(): errors.append("❌ กรุณาเลือกชื่อ-สกุล")
    if start_date>end_date: errors.append("❌ วันที่เริ่มต้องน้อยกว่าหรือเท่ากับวันที่สิ้นสุด")
    if not reason or len(reason.strip())<5: errors.append("❌ กรุณาระบุเหตุผลอย่างน้อย 5 ตัวอักษร")
    if not df_leave.empty and name:
        s,e=pd.to_datetime(start_date),pd.to_datetime(end_date)
        overlap=df_leave[(df_leave["ชื่อ-สกุล"]==name)&(df_leave["วันที่เริ่ม"]<=e)&(df_leave["วันที่สิ้นสุด"]>=s)]
        if not overlap.empty: errors.append("❌ มีการลาซ้ำในช่วงเวลานี้แล้ว")
    return errors

def validate_travel_data(staff_list,project,location,start_date,end_date) -> List[str]:
    errors=[]
    if not staff_list: errors.append("❌ กรุณาเลือกผู้เดินทางอย่างน้อย 1 คน")
    if not project or len(project.strip())<3: errors.append("❌ กรุณาระบุชื่อโครงการ/กิจกรรม")
    if not location or len(location.strip())<3: errors.append("❌ กรุณาระบุสถานที่")
    if start_date>end_date: errors.append("❌ วันที่เริ่มต้องน้อยกว่าหรือเท่ากับวันที่สิ้นสุด")
    return errors

def get_leave_used(name:str,leave_type:str,df_leave:pd.DataFrame,year:int) -> int:
    if df_leave.empty or "ชื่อ-สกุล" not in df_leave.columns: return 0
    mask=(df_leave["ชื่อ-สกุล"]==name)&(df_leave["ประเภทการลา"]==leave_type)&(df_leave["วันที่เริ่ม"].dt.year==year)
    return int(df_leave.loc[mask,"จำนวนวันลา"].sum())

def quota_bar_html(used:int,quota:int) -> str:
    pct=min(used/quota,1.0) if quota>0 else 1.0
    color="#22c55e" if pct<0.8 else ("#f59e0b" if pct<1.0 else "#ef4444")
    return f'<div class="quota-bar-wrap"><div class="quota-bar-fill" style="width:{pct*100:.0f}%;background:{color};"></div></div>'

def check_leave_quota(name:str,leave_type:str,days_req:int,df_leave:pd.DataFrame,year:int) -> Optional[str]:
    quota=LEAVE_QUOTA.get(leave_type,9999); used=get_leave_used(name,leave_type,df_leave,year); remaining=quota-used
    if days_req>remaining: return f"❌ ลาเกินสิทธิ์! คงเหลือ {remaining} วัน (ขอ {days_req} วัน, ใช้ไปแล้ว {used}/{quota} วัน)"
    if (used+days_req)/quota>=0.8: return f"⚠️ เตือน: จะใช้สิทธิ์ลา{leave_type}ไปแล้ว {used+days_req}/{quota} วัน (ใกล้หมดสิทธิ์)"
    return None

def check_admin_password(password:str) -> bool:
    return password==st.secrets.get("admin_password","204486")

# ===========================
# 🚀 DataCache System
# ===========================
_CACHE_TTL_SEC = 300

def _cache_is_fresh() -> bool:
    ts=st.session_state.get("_data_loaded_at")
    return ts is not None and (dt.datetime.now()-ts).total_seconds()<_CACHE_TTL_SEC

def _load_all_data_to_cache(force: bool = False) -> None:
    """
    โหลดข้อมูลทั้งหมดลง session_state
    - ครั้งแรก: โหลดทุกไฟล์ แสดง progress รายไฟล์
    - ครั้งต่อไป (cache ยังสด): return ทันที ไม่ยิง Drive เลย
    - force=True: โหลดใหม่ทุกไฟล์
    """
    if not force and _cache_is_fresh():
        return

    # ถ้า force → ล้าง @st.cache_data ของทุกฟังก์ชันอ่านไฟล์
    if force:
        for fn in [read_excel_from_drive, read_attendance_report,
                   load_all_travel, load_manual_scans, _read_file_by_id]:
            try:
                fn.clear()
            except Exception:
                pass

    ph = st.empty()

    # ── 1. ไฟล์หลัก 3 ไฟล์ (เบา) ─────────────────────────────
    ph.caption("⏳ กำลังโหลด leave_report...")
    df_leave, _fid_leave = read_excel_with_backup(
        FILE_LEAVE, dedup_cols=["ชื่อ-สกุล","วันที่เริ่ม","ประเภทการลา"])

    ph.caption("⏳ กำลังโหลด travel_report...")
    df_travel, _fid_travel = read_excel_with_backup(
        FILE_TRAVEL, dedup_cols=["ชื่อ-สกุล","วันที่เริ่ม","เรื่อง/กิจกรรม"])

    ph.caption("⏳ กำลังโหลด staff_master...")
    df_staff, _fid_staff = read_excel_with_backup(
        FILE_STAFF, dedup_cols=["ชื่อ-สกุล"])

    # ── 2. ไฟล์หนัก (attendance + manual + travel_all) ───────
    ph.caption("⏳ กำลังโหลดข้อมูลสแกนนิ้ว...")
    df_att    = read_attendance_report()

    ph.caption("⏳ กำลังโหลดข้อมูลสแกนนิ้ว (manual)...")
    df_manual = load_manual_scans()

    ph.caption("⏳ กำลังโหลดข้อมูลไปราชการทั้งหมด...")
    df_travel_all = load_all_travel()

    # ── 3. Preprocess ─────────────────────────────────────────
    ph.caption("⏳ กำลังประมวลผลข้อมูล...")
    df_leave, df_travel, df_att = preprocess_dataframes(df_leave, df_travel, df_att)
    _, df_travel_all, _         = preprocess_dataframes(pd.DataFrame(), df_travel_all, pd.DataFrame())
    df_att = merge_attendance_with_manual(df_att, df_manual)

    # ── 4. บันทึกลง session_state ครบทุกตัว ──────────────────
    st.session_state.update({
        "cache_leave":       df_leave,
        "cache_travel":      df_travel,
        "cache_travel_all":  df_travel_all,
        "cache_att":         df_att,
        "cache_staff":       df_staff,
        "cache_manual":      df_manual,
        "_fid_leave":        _fid_leave,
        "_fid_travel":       _fid_travel,
        "_fid_staff":        _fid_staff,
        "_data_loaded_at":   dt.datetime.now(),
    })

    # ── dtype optimization: ลด RAM 30-50% ────────────────────
    def _optimize_dtypes(df: pd.DataFrame) -> pd.DataFrame:
        """แปลง string columns เป็น category เพื่อลด memory"""
        if df.empty: return df
        for col in df.select_dtypes(include=['object']).columns:
            # category เหมาะกับ column ที่มีค่าซ้ำมาก
            if df[col].nunique() / max(len(df), 1) < 0.5:  # category threshold
                try: df[col] = df[col].astype('category')
                except Exception: pass
        return df

    df_leave      = _optimize_dtypes(df_leave)
    df_travel     = _optimize_dtypes(df_travel)
    df_staff      = _optimize_dtypes(df_staff)
    df_travel_all = _optimize_dtypes(df_travel_all)

    # attendance ใหญ่มาก — optimize เฉพาะ string cols
    for col in ["ชื่อ-สกุล", "เดือน", "สถานะสแกน", "_source"]:
        if col in df_att.columns:
            try: df_att[col] = df_att[col].astype('category')
            except Exception: pass

    # ล้าง memory หลังโหลดข้อมูลขนาดใหญ่
    gc.collect()

    ph.empty()
    logger.info(
        "Cache loaded: leave=%d travel=%d att=%d staff=%d travel_all=%d",
        len(df_leave), len(df_travel), len(df_att), len(df_staff), len(df_travel_all),
    )

def _dc(key:str,default=None):
    val=st.session_state.get(key,default)
    return val if val is not None else (pd.DataFrame() if default is None else default)

def _invalidate_cache() -> None:
    st.session_state.pop("_data_loaded_at",None)

# ===========================
# 🖥️ Sidebar (init ก่อน)
# ===========================
with st.sidebar:
    st.markdown("## 🏥 สคร.9 HR System"); st.markdown("---")
    menu = st.radio("เมนูใช้งาน",[
        "🏠 หน้าหลัก","📊 Dashboard & รายงาน","📅 ตรวจสอบการปฏิบัติงาน","📅 ปฏิทินกลาง",
        "🧭 บันทึกไปราชการ","🕒 บันทึกการลา","📈 วันลาคงเหลือ","👤 จัดการบุคลากร","🔔 กิจกรรมล่าสุด","⚙️ ผู้ดูแลระบบ",
    ], label_visibility="collapsed")
    st.markdown("---")
    loaded_at=st.session_state.get("_data_loaded_at")
    if loaded_at:
        age_sec=int((dt.datetime.now()-loaded_at).total_seconds())
        st.caption(f"🗄️ Cache: {age_sec//60}:{age_sec%60:02d} นาที")
    if st.button("🔄 โหลดข้อมูลใหม่",use_container_width=True):
        _load_all_data_to_cache(force=True); st.rerun()
    st.caption(f"v3.0 | {dt.date.today().strftime('%d/%m/%Y')}")

# ✅ FIX: เรียก cache หลัง sidebar init ครบแล้ว
# ตรวจ session_state ก่อน — ป้องกัน health check timeout ตอน startup
if "cache_leave" not in st.session_state:
    with st.spinner("⏳ โหลดข้อมูลเริ่มต้นระบบ..."):
        _load_all_data_to_cache()
else:
    _load_all_data_to_cache()  # ถ้ามีแล้ว จะ return ทันทีถ้า cache ยังสด

# ===========================
# 🏠 หน้าหลัก
# ===========================
if menu == "🏠 หน้าหลัก":
    st.markdown('<div class="section-header">🏥 ระบบติดตามการลา ไปราชการ และการปฏิบัติงาน<br>สำนักงานป้องกันควบคุมโรคที่ 9</div>', unsafe_allow_html=True)
    df_leave,df_travel=_dc("cache_leave"),_dc("cache_travel")
    c1,c2,c3,c4=st.columns(4)
    this_month=dt.date.today().strftime("%Y-%m")
    leave_tm=len(df_leave[df_leave["วันที่เริ่ม"].dt.strftime("%Y-%m")==this_month]) if not df_leave.empty and "วันที่เริ่ม" in df_leave.columns else 0
    travel_tm=len(df_travel[df_travel["วันที่เริ่ม"].dt.strftime("%Y-%m")==this_month]) if not df_travel.empty and "วันที่เริ่ม" in df_travel.columns else 0
    c1.metric("📋 ลาเดือนนี้",f"{leave_tm} ครั้ง"); c2.metric("🚗 ราชการเดือนนี้",f"{travel_tm} ครั้ง")
    c3.metric("📋 ลารวมทั้งหมด",f"{len(df_leave)} ครั้ง"); c4.metric("🚗 ราชการรวมทั้งหมด",f"{len(df_travel)} ครั้ง")
    st.markdown("---")
    col_news,col_feat=st.columns([2,1])
    with col_news:
        st.subheader("🆕 อัปเดต v3.0 (Optimized)")
        st.markdown("""| ฟีเจอร์ | สถานะ |\n|--------|------|\n| ⚡ O(1) Dictionary Lookup | ✅ |\n| 🗄️ DataCache โหลดครั้งเดียว | ✅ |\n| 🔒 Thread-safe Drive Service | ✅ |\n| 📅 วันที่ทุกรูปแบบ (พ.ศ./ค.ศ.) | ✅ |""")
    with col_feat:
        st.subheader("⚙️ สถานะการเชื่อมต่อ")
        st.markdown(f"LINE Notify: {'🟢 เชื่อมต่อแล้ว' if st.secrets.get('line_notify_token','') else '🔴 ยังไม่ตั้งค่า'}")
        st.markdown("Google Drive: 🟢 เชื่อมต่อแล้ว")
        st.markdown(f"Staff Master: {'🟢 มีข้อมูล' if not _dc('cache_staff').empty else '🟡 ยังไม่มีข้อมูล'}")

# ===========================
# 📊 Dashboard
# ===========================
elif menu == "📊 Dashboard & รายงาน":
    st.markdown('<div class="section-header">📊 Dashboard & วิเคราะห์ข้อมูล</div>', unsafe_allow_html=True)
    df_att        = _dc("cache_att")
    df_leave      = _dc("cache_leave")
    df_staff      = _dc("cache_staff")
    df_travel_all = _dc("cache_travel_all")

    LATE_CUT = dt.time(8, 31)

    # ── คำนวณ KPI ──────────────────────────────────────────────
    def _att_status(row):
        if pd.to_datetime(row["วันที่"], errors="coerce").weekday() >= 5: return "วันหยุด"
        t_in  = parse_time(row.get("เวลาเข้า", ""))
        t_out = parse_time(row.get("เวลาออก",  ""))
        if not t_in and not t_out:                                      return "ขาดงาน"
        if (t_in and not t_out) or (not t_in and t_out) or (t_in == t_out): return "ลืมสแกน"
        if t_in >= LATE_CUT:                                            return "มาสาย"
        return "มาปกติ"

    if not df_att.empty:
        df_att = df_att.copy()
        df_att["วันที่"]       = pd.to_datetime(df_att["วันที่"], errors="coerce")
        df_att["เดือน"]        = df_att["วันที่"].dt.strftime("%Y-%m")
        df_att["สถานะสแกน"]   = df_att.apply(_att_status, axis=1)
        df_work    = df_att[~df_att["สถานะสแกน"].isin(["วันหยุด"])]
        total_work = len(df_work)
        n_ok    = len(df_work[df_work["สถานะสแกน"] == "มาปกติ"])
        n_late  = len(df_work[df_work["สถานะสแกน"] == "มาสาย"])
        n_absent= len(df_work[df_work["สถานะสแกน"] == "ขาดงาน"])
        n_forgot= len(df_work[df_work["สถานะสแกน"] == "ลืมสแกน"])
        pct_ok  = n_ok / total_work * 100 if total_work else 0
        pct_late= n_late / total_work * 100 if total_work else 0
    else:
        total_work = n_ok = n_late = n_absent = n_forgot = 0
        pct_ok = pct_late = 0.0
        df_work = pd.DataFrame()

    # ── KPI Cards ──────────────────────────────────────────────
    kc1, kc2, kc3, kc4 = st.columns(4)
    kc1.metric("🗓️ วันทำการรวม",  f"{total_work:,}")
    kc2.metric("✅ อัตรามาปกติ",   f"{pct_ok:.1f}%",   delta=f"{n_ok:,} วัน")
    kc3.metric("⏰ อัตรามาสาย",    f"{pct_late:.1f}%", delta=f"{n_late:,} วัน",  delta_color="inverse")
    kc4.metric("❌ อัตราขาดงาน",   f"{n_absent/total_work*100:.1f}%" if total_work else "0%",
               delta=f"{n_absent:,} วัน", delta_color="inverse")

    st.divider()

    # ── 5 Tabs ─────────────────────────────────────────────────
    tab_summary, tab_trend, tab_charts, tab_insight, tab_export = st.tabs([
        "📋 สรุปรายบุคคล", "📈 แนวโน้มรายเดือน", "📊 กราฟ", "🔍 วิเคราะห์", "📥 Export",
    ])

    # ── Tab 1: สรุปรายบุคคล ───────────────────────────────────
    with tab_summary:
        if df_work.empty:
            st.info("ไม่มีข้อมูลการสแกนนิ้ว")
        else:
            # filter เดือน
            months_avail = sorted(df_att["เดือน"].dropna().unique().tolist())
            sel_month = st.selectbox("เดือน", months_avail,
                                     index=len(months_avail)-1 if months_avail else 0,
                                     key="dash_month")
            df_m = df_work[df_work["เดือน"] == sel_month] if sel_month else df_work

            # สรุปรายบุคคล
            summary_rows = []
            for name, grp in df_m.groupby("ชื่อ-สกุล"):
                total = len(grp)
                ok    = len(grp[grp["สถานะสแกน"] == "มาปกติ"])
                late  = len(grp[grp["สถานะสแกน"] == "มาสาย"])
                absent= len(grp[grp["สถานะสแกน"] == "ขาดงาน"])
                forgot= len(grp[grp["สถานะสแกน"] == "ลืมสแกน"])
                pct   = ok / total * 100 if total else 0
                if   pct >= 80: badge = "🟢"
                elif pct >= 60: badge = "🟡"
                else:           badge = "🔴"
                summary_rows.append({
                    "ชื่อ-สกุล":  name,
                    "วันทำการ":   total,
                    "มาปกติ":     ok,
                    "มาสาย":      late,
                    "ขาดงาน":     absent,
                    "ลืมสแกน":    forgot,
                    "% มาปกติ":   round(pct, 1),
                    "สถานะ":       badge,
                })
            if summary_rows:
                df_sum = pd.DataFrame(summary_rows).sort_values("% มาปกติ", ascending=False)
                st.dataframe(df_sum, use_container_width=True, height=450)
                st.caption(f"🟢 ≥ 80%   🟡 60–79%   🔴 < 60%")

    # ── Tab 2: แนวโน้มรายเดือน ────────────────────────────────
    with tab_trend:
        if df_work.empty:
            st.info("ไม่มีข้อมูลสแกนนิ้ว")
        else:
            df_monthly = (df_work.groupby("เดือน")["สถานะสแกน"]
                          .value_counts().unstack(fill_value=0).reset_index())
            for col in ["มาปกติ", "มาสาย", "ขาดงาน", "ลืมสแกน"]:
                if col not in df_monthly.columns: df_monthly[col] = 0
            df_monthly["วันรวม"]   = df_monthly[["มาปกติ","มาสาย","ขาดงาน","ลืมสแกน"]].sum(axis=1)
            df_monthly["% มาปกติ"] = (df_monthly["มาปกติ"] / df_monthly["วันรวม"].replace(0, 1) * 100).round(1)
            df_monthly = df_monthly.sort_values("เดือน")

            # progress bar inline
            def _bar(pct):
                c = "#22c55e" if pct >= 80 else ("#f59e0b" if pct >= 60 else "#ef4444")
                return f'<div style="background:#e2e8f0;border-radius:4px;height:8px"><div style="width:{min(pct,100):.0f}%;background:{c};height:8px;border-radius:4px"></div></div>'

            st.dataframe(
                df_monthly[["เดือน","มาปกติ","มาสาย","ขาดงาน","ลืมสแกน","วันรวม","% มาปกติ"]],
                use_container_width=True, height=400,
            )

    # ── Tab 3: กราฟ ──────────────────────────────────────────
    with tab_charts:
        if df_work.empty:
            st.info("ไม่มีข้อมูล")
        else:
            df_monthly_c = (df_work.groupby("เดือน")["สถานะสแกน"]
                            .value_counts().unstack(fill_value=0).reset_index())
            for col in ["มาปกติ", "มาสาย", "ขาดงาน", "ลืมสแกน"]:
                if col not in df_monthly_c.columns: df_monthly_c[col] = 0
            df_monthly_c["วันรวม"]    = df_monthly_c[["มาปกติ","มาสาย","ขาดงาน","ลืมสแกน"]].sum(axis=1)
            df_monthly_c["% มาปกติ"]  = (df_monthly_c["มาปกติ"] / df_monthly_c["วันรวม"].replace(0, 1) * 100).round(1)
            df_monthly_c = df_monthly_c.sort_values("เดือน")

            col_c1, col_c2 = st.columns(2)

            # กราฟ Line: % มาปกติ รายเดือน + เส้นเกณฑ์ 80%
            with col_c1:
                st.subheader("📈 % มาปกติ รายเดือน")
                line = alt.Chart(df_monthly_c).mark_line(point=True, color="#6366f1", strokeWidth=2.5).encode(
                    x=alt.X("เดือน:O", title="เดือน"),
                    y=alt.Y("% มาปกติ:Q", title="% มาปกติ", scale=alt.Scale(domain=[0, 100])),
                    tooltip=["เดือน", "% มาปกติ", "มาปกติ", "วันรวม"],
                )
                rule = alt.Chart(pd.DataFrame({"y": [80]})).mark_rule(
                    color="red", strokeDash=[6, 3], strokeWidth=1.5
                ).encode(y="y:Q")
                st.altair_chart((line + rule).properties(height=280), use_container_width=True)

            # กราฟ Stacked Bar: สัดส่วนสถานะรายเดือน
            with col_c2:
                st.subheader("📊 สัดส่วนสถานะรายเดือน")
                df_melt = df_monthly_c.melt(
                    id_vars="เดือน",
                    value_vars=["มาปกติ", "มาสาย", "ขาดงาน", "ลืมสแกน"],
                    var_name="สถานะ", value_name="จำนวน",
                )
                bar = alt.Chart(df_melt).mark_bar().encode(
                    x=alt.X("เดือน:O", title="เดือน"),
                    y=alt.Y("จำนวน:Q", title="จำนวนวัน"),
                    color=alt.Color("สถานะ:N", scale=alt.Scale(
                        domain=["มาปกติ", "มาสาย", "ขาดงาน", "ลืมสแกน"],
                        range=["#22c55e", "#f59e0b", "#ef4444", "#a78bfa"],
                    )),
                    tooltip=["เดือน", "สถานะ", "จำนวน"],
                ).properties(height=280)
                st.altair_chart(bar, use_container_width=True)

            # กราฟ Bar: วันลาแยกตามกลุ่มงาน
            if not df_leave.empty and "กลุ่มงาน" in df_leave.columns:
                st.subheader("📋 วันลารวมแยกตามกลุ่มงาน (Top 10)")
                df_lc = df_leave.groupby("กลุ่มงาน")["จำนวนวันลา"].sum().nlargest(10).reset_index()
                st.altair_chart(
                    alt.Chart(df_lc).mark_bar(
                        cornerRadiusTopRight=4, cornerRadiusBottomRight=4
                    ).encode(
                        x=alt.X("จำนวนวันลา:Q", title="วันลารวม"),
                        y=alt.Y("กลุ่มงาน:N", sort="-x", title=""),
                        color=alt.value("#6366f1"),
                        tooltip=["กลุ่มงาน", "จำนวนวันลา"],
                    ).properties(height=320),
                    use_container_width=True,
                )

    # ── Tab 4: วิเคราะห์ ──────────────────────────────────────
    with tab_insight:
        st.subheader("🔍 ข้อวิเคราะห์จากข้อมูลจริง")
        if df_work.empty:
            st.info("ไม่มีข้อมูลเพียงพอสำหรับการวิเคราะห์")
        else:
            insights = []

            # 1. อัตรามาปกติรวม
            insights.append(f"📌 อัตรามาปกติรวมทั้งหมด **{pct_ok:.1f}%** จากทั้งหมด {total_work:,} วันทำการ"
                             + (" (✅ ผ่านเกณฑ์ 80%)" if pct_ok >= 80 else " (⚠️ ต่ำกว่าเกณฑ์ 80%)"))

            # 2. บุคลากรมาสายมากสุด
            if "ชื่อ-สกุล" in df_work.columns:
                late_by_name = df_work[df_work["สถานะสแกน"] == "มาสาย"].groupby("ชื่อ-สกุล").size().nlargest(3)
                if not late_by_name.empty:
                    top_late = ", ".join([f"{n} ({c} วัน)" for n, c in late_by_name.items()])
                    insights.append(f"⏰ บุคลากรมาสายสูงสุด 3 อันดับ: {top_late}")

            # 3. บุคลากรขาดงานมากสุด
            absent_by_name = df_work[df_work["สถานะสแกน"] == "ขาดงาน"].groupby("ชื่อ-สกุล").size().nlargest(3)
            if not absent_by_name.empty:
                top_abs = ", ".join([f"{n} ({c} วัน)" for n, c in absent_by_name.items()])
                insights.append(f"❌ บุคลากรขาดงานสูงสุด 3 อันดับ: {top_abs}")

            # 4. เดือนที่มาปกติน้อยสุด
            if "เดือน" in df_work.columns:
                m_ok = df_work[df_work["สถานะสแกน"] == "มาปกติ"].groupby("เดือน").size()
                m_total = df_work.groupby("เดือน").size()
                m_pct = (m_ok / m_total * 100).dropna()
                if not m_pct.empty:
                    worst_m = m_pct.idxmin()
                    insights.append(f"📅 เดือนที่มาปกติน้อยที่สุด: **{worst_m}** ({m_pct[worst_m]:.1f}%)")
                    best_m = m_pct.idxmax()
                    insights.append(f"📅 เดือนที่มาปกติมากที่สุด: **{best_m}** ({m_pct[best_m]:.1f}%)")

            # 5. ลืมสแกนนิ้ว
            if n_forgot > 0:
                insights.append(f"🟣 มีการลืมสแกนนิ้ว **{n_forgot:,} ครั้ง** ({n_forgot/total_work*100:.1f}% ของวันทำการ)")

            # 6. ประเภทลาที่ใช้มากสุด
            if not df_leave.empty and "ประเภทการลา" in df_leave.columns:
                top_leave = df_leave["ประเภทการลา"].value_counts().head(1)
                if not top_leave.empty:
                    insights.append(f"🗂️ ประเภทการลาที่ใช้มากที่สุด: **{top_leave.index[0]}** ({top_leave.iloc[0]:,} ครั้ง)")

            # 7. สัดส่วนขาดงาน warning
            if total_work > 0 and n_absent / total_work > 0.1:
                insights.append(f"🚨 สัดส่วนขาดงาน **{n_absent/total_work*100:.1f}%** สูงเกิน 10% ควรตรวจสอบ")

            for ins in insights:
                st.markdown(f"- {ins}")

    # ── Tab 5: Export ─────────────────────────────────────────
    with tab_export:
        today = dt.date.today()
        month_opts = pd.date_range(f"{today.year-2}-01-01", f"{today.year+1}-12-31",
                                   freq="MS").strftime("%Y-%m").tolist()
        export_month = st.selectbox(
            "เลือกเดือน", month_opts,
            index=month_opts.index(today.strftime("%Y-%m")) if today.strftime("%Y-%m") in month_opts else 0,
            key="export_month_sel",
        )
        if st.button("📊 สร้างรายงาน Excel", type="primary", key="btn_export"):
            m_start = pd.to_datetime(export_month + "-01")
            m_end   = m_start + pd.offsets.MonthEnd(0)
            df_lm   = df_leave[(df_leave["วันที่เริ่ม"] >= m_start) & (df_leave["วันที่เริ่ม"] <= m_end)] \
                      if not df_leave.empty else pd.DataFrame()
            df_wm   = df_work[df_work["เดือน"] == export_month] if not df_work.empty else pd.DataFrame()
            output  = io.BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                pd.DataFrame({
                    "รายการ": ["การลา (ครั้ง)", "วันลารวม", "วันทำการ", "มาปกติ", "มาสาย", "ขาดงาน"],
                    "จำนวน": [
                        len(df_lm),
                        int(df_lm["จำนวนวันลา"].sum()) if not df_lm.empty else 0,
                        len(df_wm),
                        len(df_wm[df_wm["สถานะสแกน"] == "มาปกติ"]) if not df_wm.empty else 0,
                        len(df_wm[df_wm["สถานะสแกน"] == "มาสาย"])  if not df_wm.empty else 0,
                        len(df_wm[df_wm["สถานะสแกน"] == "ขาดงาน"]) if not df_wm.empty else 0,
                    ],
                }).to_excel(writer, sheet_name="สรุป", index=False)
                if not df_lm.empty: df_lm.to_excel(writer, sheet_name="การลา", index=False)
                if not df_wm.empty: df_wm.to_excel(writer, sheet_name="การมาปฏิบัติงาน", index=False)
            st.download_button(
                "⬇️ ดาวน์โหลดรายงาน",
                output.getvalue(),
                f"HR_Report_{export_month}.xlsx",
                mime=EXCEL_MIME,
            )

# ===========================
# 📅 ตรวจสอบการปฏิบัติงาน
# ===========================
elif menu == "📅 ตรวจสอบการปฏิบัติงาน":
    st.markdown('<div class="section-header">📅 สรุปการมาปฏิบัติงานรายวัน</div>', unsafe_allow_html=True)
    df_att,df_leave,df_staff,df_travel_all=_dc("cache_att"),_dc("cache_leave"),_dc("cache_staff"),_dc("cache_travel_all")
    all_names=get_active_staff(df_staff) or get_all_names_fallback(df_leave,df_travel_all,df_att)
    if not df_att.empty:
        df_att["วันที่"]=pd.to_datetime(df_att["วันที่"],errors="coerce").dt.normalize()
        months_att=sorted(df_att["วันที่"].dt.strftime("%Y-%m").dropna().unique().tolist())
    else: months_att=[dt.datetime.now().strftime("%Y-%m")]
    selected_months=st.multiselect("📅 เลือกเดือน",months_att,default=[months_att[-1]] if months_att else [])
    selected_names=st.multiselect("👥 บุคลากร (ว่าง = ทุกคน)",all_names)
    names_to_process=selected_names or all_names
    if not selected_months or not names_to_process: st.warning("กรุณาเลือกเดือนและบุคลากร"); st.stop()
    all_dates=pd.DatetimeIndex([])
    for ym in selected_months:
        ms=pd.to_datetime(ym+"-01"); all_dates=all_dates.append(pd.date_range(ms,ms+pd.offsets.MonthEnd(0),freq="D"))
    df_months_att=df_att[df_att["วันที่"].dt.strftime("%Y-%m").isin(selected_months)].copy() if not df_att.empty else pd.DataFrame()
    # ⚡ O(1) Dictionary Lookup
    att_dict={}
    if not df_months_att.empty:
        name_col=next((c for c in ["ชื่อ-สกุล","ชื่อพนักงาน","ชื่อ"] if c in df_months_att.columns),"ชื่อ-สกุล")
        for _,row in df_months_att.iterrows():
            d_date=row["วันที่"].date() if isinstance(row["วันที่"],pd.Timestamp) else row["วันที่"]
            att_dict[(str(row[name_col]).strip(),d_date)]=row
    holiday_dates_set=set()
    for yr in {int(ym[:4]) for ym in selected_months}: holiday_dates_set.update(get_holiday_dates(yr))
    leave_index={}
    if not df_leave.empty:
        for _,row in df_leave.dropna(subset=["วันที่เริ่ม","วันที่สิ้นสุด"]).iterrows():
            leave_index.setdefault(str(row.get("ชื่อ-สกุล","")).strip(),[]).append((row["วันที่เริ่ม"].date(),row["วันที่สิ้นสุด"].date(),str(row.get("ประเภทการลา","ลา"))))
    travel_index={}
    if not df_travel_all.empty:
        for _,row in df_travel_all.dropna(subset=["วันที่เริ่ม","วันที่สิ้นสุด"]).iterrows():
            proj=str(row.get("เรื่อง/กิจกรรม","ไปราชการ")).strip()
            names=[str(row.get("ชื่อ-สกุล","")).strip()]
            for comp in str(row.get("ผู้ร่วมเดินทาง","")).replace("\n",",").split(","):
                comp=re.sub(r"\d+\.\s*","",comp).strip()
                if comp and len(comp)>=3 and comp.lower()!="nan": names.append(comp)
            for p in set(names): travel_index.setdefault(p,[]).append((row["วันที่เริ่ม"].date(),row["วันที่สิ้นสุด"].date(),proj))
    records=[]; prog=st.progress(0,text="กำลังประมวลผล..."); LATE_CUTOFF=dt.time(8,31)
    for i,name in enumerate(names_to_process):
        prog.progress((i+1)/len(names_to_process),text=f"กำลังประมวลผล {name}...")
        for d in all_dates:
            d_date=d.date()
            rec={"ชื่อพนักงาน":name,"วันที่":d_date,"เดือน":d.strftime("%Y-%m"),"เวลาเข้า":"","เวลาออก":"","สถานะ":""}
            in_leave,leave_type=False,""
            for ls,le,ltype in leave_index.get(name,[]):
                if ls<=d_date<=le: in_leave,leave_type=True,ltype; break
            in_travel,travel_proj=False,""
            for ts,te,proj in travel_index.get(name,[]):
                if ts<=d_date<=te: in_travel,travel_proj=True,proj; break
            att_row=att_dict.get((name,d_date))
            if in_leave: rec["สถานะ"]=f"ลา ({leave_type})"
            elif in_travel: rec["สถานะ"]=f"ไปราชการ ({travel_proj})" if travel_proj!="ไปราชการ" else "ไปราชการ"
            elif d.weekday()>=5: rec["สถานะ"]="วันหยุด"
            elif d_date in holiday_dates_set: rec["สถานะ"]="วันหยุดพิเศษ"
            elif att_row is not None:
                rec["เวลาเข้า"],rec["เวลาออก"]=att_row.get("เวลาเข้า",""),att_row.get("เวลาออก","")
                t_in,t_out=parse_time(rec["เวลาเข้า"]),parse_time(rec["เวลาออก"])
                is_manual=str(att_row.get("_source","")).strip()=="manual"
                if not t_in and not t_out: base_status="ขาดงาน"
                elif (t_in and not t_out) or (not t_in and t_out) or (t_in==t_out): base_status="ลืมสแกน"
                elif t_in>=LATE_CUTOFF: base_status="มาสาย"
                else: base_status="มาปกติ"
                rec["สถานะ"]=f"{base_status} (HR คีย์แทน)" if is_manual else base_status
            else: rec["สถานะ"]="ขาดงาน"
            records.append(rec)
    prog.empty()
    df_result=pd.DataFrame(records).sort_values(["ชื่อพนักงาน","วันที่"])
    STATUS_COLORS={"มาปกติ":"background-color:#dcfce7","มาสาย":"background-color:#fef9c3","ขาดงาน":"background-color:#fee2e2","ลืมสแกน":"background-color:#f3e8ff","วันหยุด":"background-color:#f1f5f9","วันหยุดพิเศษ":"background-color:#e0f2fe"}
    def color_status(val):
        for k,v in STATUS_COLORS.items():
            if str(val).startswith(k): return v
        return ""
    st.dataframe(df_result.style.applymap(color_status,subset=["สถานะ"]),use_container_width=True,height=500)

# ===========================
# 📅 ปฏิทินกลาง
# ===========================
elif menu == "📅 ปฏิทินกลาง":
    st.markdown('<div class="section-header">📅 ปฏิทินกลางหน่วยงาน</div>', unsafe_allow_html=True)
    df_leave,df_travel,df_staff=_dc("cache_leave"),_dc("cache_travel"),_dc("cache_staff")
    all_names=get_active_staff(df_staff) or get_all_names_fallback(df_leave,df_travel,pd.DataFrame())
    today=dt.date.today()
    month_opts=pd.date_range(f"{today.year-1}-01-01",f"{today.year+1}-12-31",freq="MS").strftime("%Y-%m").tolist()
    cal_month=st.selectbox("เดือน",month_opts,index=month_opts.index(today.strftime("%Y-%m")) if today.strftime("%Y-%m") in month_opts else 0)
    cal_group=st.selectbox("กลุ่มงาน",["ทุกกลุ่ม"]+STAFF_GROUPS)
    cal_names_sel=st.multiselect("เลือกบุคลากร (ว่าง=ทุกคน)",all_names)
    m_start=pd.to_datetime(cal_month+"-01"); m_end=m_start+pd.offsets.MonthEnd(0)
    date_range=pd.date_range(m_start,m_end,freq="D")
    names_to_show=cal_names_sel or all_names
    if cal_group!="ทุกกลุ่ม" and not df_staff.empty and "กลุ่มงาน" in df_staff.columns:
        grp_names=df_staff[df_staff["กลุ่มงาน"]==cal_group]["ชื่อ-สกุล"].tolist()
        names_to_show=[n for n in names_to_show if n in grp_names]
    cal_records=[]
    for name in names_to_show:
        for d in date_range:
            status="วันหยุด" if d.weekday()>=5 else "ปฏิบัติงาน"
            if not df_leave.empty:
                ml=df_leave[(df_leave["ชื่อ-สกุล"]==name)&(df_leave["วันที่เริ่ม"]<=d)&(df_leave["วันที่สิ้นสุด"]>=d)]
                if not ml.empty: status="ลา"
            if not df_travel.empty:
                mt=df_travel[(df_travel["ชื่อ-สกุล"]==name)&(df_travel["วันที่เริ่ม"]<=d)&(df_travel["วันที่สิ้นสุด"]>=d)]
                if not mt.empty: status="ไปราชการ"
            cal_records.append({"ชื่อ-สกุล":name,"วันที่":d.strftime("%d"),"สถานะ":status,"วันที่เต็ม":d})
    if cal_records:
        df_cal=pd.DataFrame(cal_records)
        heatmap=alt.Chart(df_cal).mark_rect(stroke="white",strokeWidth=1).encode(x=alt.X("วันที่:O",title="วันที่",sort=None),y=alt.Y("ชื่อ-สกุล:N",title=""),color=alt.Color("สถานะ:N",scale=alt.Scale(domain=["ปฏิบัติงาน","ลา","ไปราชการ","วันหยุด"],range=["#22c55e","#60a5fa","#f59e0b","#e2e8f0"]),legend=alt.Legend(orient="bottom")),tooltip=["ชื่อ-สกุล","วันที่เต็ม","สถานะ"]).properties(height=max(200,len(names_to_show)*22),title=f"ปฏิทินการปฏิบัติงาน — {cal_month}")
        st.altair_chart(heatmap, use_container_width=True)
    else: st.info("ไม่มีข้อมูล")

# ===========================
# 🧭 บันทึกไปราชการ
# ===========================
elif menu == "🧭 บันทึกไปราชการ":
    st.markdown('<div class="section-header">🧭 บันทึกการเดินทางไปราชการ</div>', unsafe_allow_html=True)
    df_travel=_dc("cache_travel"); _travel_fid=st.session_state.get("_fid_travel")
    df_staff=_dc("cache_staff"); ALL_NAMES=get_active_staff(df_staff) or get_all_names_fallback(_dc("cache_leave"),df_travel,_dc("cache_att"))
    st.info(f"📂 ข้อมูลไปราชการปัจจุบัน: **{len(df_travel)} รายการ**")
    with st.form("form_travel"):
        col1,col2=st.columns(2)
        with col1:
            group_job=st.selectbox("กลุ่มงาน",STAFF_GROUPS)
            project=st.text_input("ชื่อโครงการ/กิจกรรม *")
            location=st.text_input("สถานที่ *")
        with col2:
            d_start=st.date_input("วันที่เริ่ม *",value=dt.date.today())
            d_end=st.date_input("วันที่สิ้นสุด *",value=dt.date.today())
        st.markdown("**👥 รายชื่อผู้เดินทาง**")
        selected_staff=st.multiselect("เลือกจากระบบ",ALL_NAMES)
        extra_staff_text=st.text_area("เพิ่มชื่อที่ไม่มีในระบบ (คั่นด้วย ,)")
        uploaded_pdf=st.file_uploader("แนบเอกสาร (PDF)",type=["pdf"])
        submitted=st.form_submit_button("💾 บันทึกข้อมูล",use_container_width=True,type="primary")
        if submitted:
            final_staff=list(selected_staff)+[n.strip() for n in extra_staff_text.replace("\n",",").split(",") if n.strip()]
            final_staff=sorted(set(final_staff))
            errors=validate_travel_data(final_staff,project,location,d_start,d_end)
            if errors:
                for e in errors: st.error(e)
            else:
                # ✅ FIX: ใช้ flag แทน st.rerun() ใน with st.status
                _rerun_after=False
                with st.status("กำลังบันทึก...",expanded=True) as status:
                    try:
                        link="-"
                        if uploaded_pdf:
                            fid_att=get_or_create_folder(ATTACHMENT_FOLDER_NAME,FOLDER_ID)
                            if fid_att: link=upload_pdf_to_drive(uploaded_pdf,f"TRAVEL_{dt.datetime.now().strftime('%Y%m%d_%H%M')}.pdf",fid_att)
                        ts=dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"); days=count_weekdays(d_start,d_end)
                        new_rows=[{"Timestamp":ts,"กลุ่มงาน":group_job,"ชื่อ-สกุล":p,"เรื่อง/กิจกรรม":project,"สถานที่":location,"วันที่เริ่ม":pd.to_datetime(d_start),"วันที่สิ้นสุด":pd.to_datetime(d_end),"จำนวนวัน":days,"ไฟล์แนบ":link} for p in final_staff]
                        backup_excel(FILE_TRAVEL,df_travel)
                        df_upd=pd.concat([df_travel,pd.DataFrame(new_rows)],ignore_index=True)
                        if write_excel_to_drive(FILE_TRAVEL,df_upd,known_file_id=_travel_fid):
                            send_line_notify(format_travel_notify(final_staff,project,location,d_start,d_end,days))
                            log_activity("ไปราชการ",f"{project} @ {location}",", ".join(final_staff[:3]))
                            status.update(label=f"✅ บันทึกสำเร็จ ({len(final_staff)} ท่าน)",state="complete")
                            st.toast(f"✅ บันทึกไปราชการสำเร็จ {len(final_staff)} ท่าน",icon="✅")
                            _rerun_after=True
                        else: status.update(label="❌ บันทึกล้มเหลว",state="error")
                    except Exception as e: logger.error(f"travel form: {e}"); status.update(label=f"❌ {e}",state="error")
                if _rerun_after: time.sleep(1); st.rerun()

# ===========================
# 🕒 บันทึกการลา
# ===========================
elif menu == "🕒 บันทึกการลา":
    st.markdown('<div class="section-header">🕒 บันทึกการลา</div>', unsafe_allow_html=True)
    df_leave=_dc("cache_leave"); _leave_fid=st.session_state.get("_fid_leave"); df_staff=_dc("cache_staff")
    ALL_NAMES=get_active_staff(df_staff) or get_all_names_fallback(df_leave,_dc("cache_travel"),_dc("cache_att"))
    st.info(f"📂 ข้อมูลการลาปัจจุบัน: **{len(df_leave)} รายการ**")
    with st.form("form_leave"):
        col1,col2=st.columns(2)
        with col1:
            l_name=st.selectbox("ชื่อ-สกุล *",ALL_NAMES)
            l_group=st.selectbox("กลุ่มงาน",STAFF_GROUPS)
            l_type=st.selectbox("ประเภทการลา *",LEAVE_TYPES)
        with col2:
            l_start=st.date_input("วันที่เริ่มลา *",value=dt.date.today())
            l_end=st.date_input("ถึงวันที่ *",value=dt.date.today())
            l_reason=st.text_area("เหตุผลการลา *")
        l_file=st.file_uploader("แนบใบลา (PDF)",type=["pdf"])
        if st.form_submit_button("💾 บันทึกการลา",use_container_width=True,type="primary"):
            days_req=count_weekdays(l_start,l_end)
            errors=validate_leave_data(l_name,l_start,l_end,l_reason,df_leave)
            quota_msg=check_leave_quota(l_name,l_type,days_req,df_leave,l_start.year) if l_name else None
            if quota_msg and quota_msg.startswith("❌"): errors.append(quota_msg)
            if errors:
                for e in errors: st.error(e)
            else:
                if quota_msg: st.warning(quota_msg)
                # ✅ FIX: flag pattern
                _rerun_leave=False
                with st.status("กำลังบันทึก...",expanded=True) as status:
                    try:
                        link="-"
                        if l_file:
                            fid=get_or_create_folder(ATTACHMENT_FOLDER_NAME,FOLDER_ID)
                            if fid: link=upload_pdf_to_drive(l_file,f"LEAVE_{l_name}_{dt.datetime.now().strftime('%Y%m%d_%H%M')}.pdf",fid)
                        new_rec={"Timestamp":dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"ชื่อ-สกุล":l_name,"กลุ่มงาน":l_group,"ประเภทการลา":l_type,"วันที่เริ่ม":pd.to_datetime(l_start),"วันที่สิ้นสุด":pd.to_datetime(l_end),"จำนวนวันลา":days_req,"เหตุผล":l_reason,"ไฟล์แนบ":link}
                        backup_excel(FILE_LEAVE,df_leave)
                        if write_excel_to_drive(FILE_LEAVE,pd.concat([df_leave,pd.DataFrame([new_rec])],ignore_index=True),known_file_id=_leave_fid):
                            send_line_notify(format_leave_notify(new_rec))
                            log_activity("การลา",f"{l_type} {days_req} วัน — {l_reason[:30]}",l_name)
                            status.update(label="✅ บันทึกสำเร็จ",state="complete")
                            st.toast(f"✅ บันทึกการลาสำเร็จ ({l_type} {days_req} วัน)",icon="✅")
                            _rerun_leave=True
                        else: status.update(label="❌ บันทึกล้มเหลว",state="error")
                    except Exception as e: logger.error(f"leave form: {e}"); status.update(label=f"❌ {e}",state="error")
                if _rerun_leave: time.sleep(1); st.rerun()

# ===========================
# 📈 วันลาคงเหลือ
# ===========================
elif menu == "📈 วันลาคงเหลือ":
    st.markdown('<div class="section-header">📈 สิทธิ์วันลาคงเหลือ</div>', unsafe_allow_html=True)
    df_leave=_dc("cache_leave"); df_staff=_dc("cache_staff")
    all_names=get_active_staff(df_staff) or get_all_names_fallback(df_leave,pd.DataFrame(),pd.DataFrame())
    sel_year=st.selectbox("ปี (พ.ศ.)",list(range(dt.date.today().year+543,dt.date.today().year+540,-1)))
    year_ad=sel_year-543
    sel_person=st.selectbox("เลือกบุคลากร",["— ทุกคน —"]+all_names)
    names_to_show=all_names if sel_person=="— ทุกคน —" else [sel_person]
    quota_rows=[]
    for name in names_to_show:
        row={"ชื่อ-สกุล":name}
        for lt in LEAVE_TYPES:
            quota=LEAVE_QUOTA.get(lt,9999); used=get_leave_used(name,lt,df_leave,year_ad)
            row[f"{lt} (ใช้)"],row[f"{lt} (คงเหลือ)"]=used,max(0,quota-used)
        quota_rows.append(row)
    if quota_rows:
        df_quota=pd.DataFrame(quota_rows)
        st.dataframe(df_quota,use_container_width=True,height=400)
        buf=io.BytesIO()
        with pd.ExcelWriter(buf,engine="xlsxwriter") as w: df_quota.to_excel(w,index=False,sheet_name="วันลาคงเหลือ")
        st.download_button("📥 Export Excel",buf.getvalue(),"leave_quota.xlsx",mime=EXCEL_MIME)

# ===========================
# 👤 จัดการบุคลากร
# ===========================
elif menu == "👤 จัดการบุคลากร":
    st.markdown('<div class="section-header">👤 จัดการฐานข้อมูลบุคลากร</div>', unsafe_allow_html=True)
    df_staff=_dc("cache_staff"); _staff_fid=st.session_state.get("_fid_staff")
    if df_staff.empty: df_staff=pd.DataFrame(columns=STAFF_MASTER_COLS)
    tab_list,tab_add,tab_edit=st.tabs(["📋 รายชื่อทั้งหมด","➕ เพิ่มบุคลากร","✏️ แก้ไข"])
    with tab_list:
        col_s,col_f=st.columns([1,2])
        with col_s: filter_status=st.selectbox("สถานะ",["ทุกสถานะ","ปฏิบัติงาน","ลาออก","ยืมตัว"])
        with col_f: filter_group=st.selectbox("กลุ่มงาน",["ทุกกลุ่ม"]+STAFF_GROUPS)
        df_show=df_staff.copy()
        if filter_status!="ทุกสถานะ" and "สถานะ" in df_show.columns: df_show=df_show[df_show["สถานะ"]==filter_status]
        if filter_group!="ทุกกลุ่ม" and "กลุ่มงาน" in df_show.columns: df_show=df_show[df_show["กลุ่มงาน"]==filter_group]
        st.caption(f"แสดง {len(df_show)} รายการ")
        st.dataframe(df_show,use_container_width=True,height=420)
    with tab_add:
        with st.form("form_add_staff"):
            c1,c2=st.columns(2)
            with c1:
                s_name=st.text_input("ชื่อ-สกุล *"); s_group=st.selectbox("กลุ่มงาน",STAFF_GROUPS)
                s_pos=st.text_input("ตำแหน่ง")
            with c2:
                s_type=st.selectbox("ประเภทบุคลากร",["ข้าราชการ","พนักงานราชการ","ลูกจ้างประจำ","จ้างเหมา"])
                s_start=st.date_input("วันเริ่มงาน",value=dt.date.today()); s_status=st.selectbox("สถานะ",["ปฏิบัติงาน","ลาออก","ยืมตัว"])
            if st.form_submit_button("✅ เพิ่มบุคลากร",type="primary"):
                if not s_name.strip(): st.error("❌ กรุณาระบุชื่อ-สกุล")
                elif not df_staff.empty and s_name.strip() in df_staff["ชื่อ-สกุล"].values: st.error("❌ ชื่อนี้มีอยู่ในระบบแล้ว")
                else:
                    new_staff={"ชื่อ-สกุล":s_name.strip(),"กลุ่มงาน":s_group,"ตำแหน่ง":s_pos,"ประเภทบุคลากร":s_type,"วันเริ่มงาน":str(s_start),"สถานะ":s_status}
                    df_staff=pd.concat([df_staff,pd.DataFrame([new_staff])],ignore_index=True)
                    if write_excel_to_drive(FILE_STAFF,df_staff,known_file_id=_staff_fid):
                        log_activity("เพิ่มบุคลากร",f"เพิ่ม {s_name} ({s_group})",s_name)
                        st.toast(f"✅ เพิ่ม {s_name} สำเร็จ",icon="✅"); time.sleep(0.5); st.rerun()
    with tab_edit:
        if df_staff.empty: st.info("ยังไม่มีข้อมูลบุคลากร")
        else:
            edit_name=st.selectbox("เลือกบุคลากรที่ต้องการแก้ไข",df_staff["ชื่อ-สกุล"].tolist())
            row_idx=df_staff[df_staff["ชื่อ-สกุล"]==edit_name].index
            if len(row_idx)>0:
                idx=row_idx[0]
                with st.form("form_edit_staff"):
                    c1,c2=st.columns(2)
                    with c1:
                        e_group=st.selectbox("กลุ่มงาน",STAFF_GROUPS,index=STAFF_GROUPS.index(df_staff.at[idx,"กลุ่มงาน"]) if "กลุ่มงาน" in df_staff.columns and df_staff.at[idx,"กลุ่มงาน"] in STAFF_GROUPS else 0)
                        e_pos=st.text_input("ตำแหน่ง",value=str(df_staff.at[idx,"ตำแหน่ง"]) if "ตำแหน่ง" in df_staff.columns else "")
                    with c2:
                        e_type=st.selectbox("ประเภทบุคลากร",["ข้าราชการ","พนักงานราชการ","ลูกจ้างประจำ","จ้างเหมา"])
                        e_status_opts=["ปฏิบัติงาน","ลาออก","ยืมตัว"]
                        cur_status=str(df_staff.at[idx,"สถานะ"]) if "สถานะ" in df_staff.columns else "ปฏิบัติงาน"
                        e_status=st.selectbox("สถานะ",e_status_opts,index=e_status_opts.index(cur_status) if cur_status in e_status_opts else 0)
                    if st.form_submit_button("✅ บันทึกการแก้ไข",use_container_width=True):
                        df_staff.at[idx,"กลุ่มงาน"]=e_group; df_staff.at[idx,"ตำแหน่ง"]=e_pos
                        df_staff.at[idx,"ประเภทบุคลากร"]=e_type; df_staff.at[idx,"สถานะ"]=e_status
                        if write_excel_to_drive(FILE_STAFF,df_staff,known_file_id=_staff_fid):
                            log_activity("แก้ไขบุคลากร",f"อัปเดต {edit_name} สถานะ→{e_status}",edit_name)
                            st.toast(f"✅ อัปเดต {edit_name} สำเร็จ",icon="✅"); st.rerun()

# ===========================
# 🔔 กิจกรรมล่าสุด
# ===========================
elif menu == "🔔 กิจกรรมล่าสุด":
    st.markdown('<div class="section-header">🔔 กิจกรรมล่าสุดในระบบ</div>', unsafe_allow_html=True)
    df_log=read_excel_from_drive(FILE_NOTIFY)
    if df_log.empty: st.info("ยังไม่มีกิจกรรมในระบบ")
    else: st.dataframe(df_log.sort_values("Timestamp",ascending=False).head(50),use_container_width=True)

# ===========================
# ⚙️ ผู้ดูแลระบบ
# ===========================
elif menu == "⚙️ ผู้ดูแลระบบ":
    st.markdown('<div class="section-header">⚙️ ผู้ดูแลระบบ</div>', unsafe_allow_html=True)
    password=st.text_input("🔑 รหัสผ่าน Admin",type="password")
    if password and check_admin_password(password):
        st.success("✅ เข้าสู่ระบบสำเร็จ")
        df_leave=_dc("cache_leave"); df_travel=_dc("cache_travel"); df_att=_dc("cache_att"); df_staff=_dc("cache_staff")
        _fid_leave=st.session_state.get("_fid_leave"); _fid_travel=st.session_state.get("_fid_travel"); _fid_staff=st.session_state.get("_fid_staff")
        _fid_map={FILE_LEAVE:_fid_leave,FILE_TRAVEL:_fid_travel,FILE_STAFF:_fid_staff,FILE_ATTEND:None}
        tab1,tab2,tab3,tab4,tab5,tab6,tab_hol=st.tabs(["📂 ไฟล์ลา","📂 ไฟล์ราชการ","📂 ไฟล์สแกนนิ้ว","📂 ไฟล์บุคลากร","🔧 ตั้งค่า","👆 คีย์สแกน","🎌 วันหยุด"])
        def admin_file_panel(df,filename,tab_obj):
            with tab_obj:
                st.subheader(f"ไฟล์: {filename}")
                st.caption(f"File ID: `{_fid_map.get(filename,'—')}`")
                if df.empty: st.warning("⚠️ ไม่มีข้อมูล")
                else:
                    st.dataframe(df.head(20),use_container_width=True)
                    st.caption(f"ทั้งหมด {len(df)} แถว | {len(df.columns)} คอลัมน์")
                    col_d1,col_d2=st.columns(2)
                    with col_d1:
                        buf=io.BytesIO()
                        with pd.ExcelWriter(buf,engine="xlsxwriter") as w: df.to_excel(w,index=False)
                        st.download_button("⬇️ Excel",buf.getvalue(),filename,use_container_width=True)
                    with col_d2: st.download_button("⬇️ CSV",df.to_csv(index=False).encode("utf-8-sig"),filename.replace(".xlsx",".csv"),"text/csv",use_container_width=True)
                st.divider(); st.warning("⚠️ การอัปโหลดจะเขียนทับข้อมูลเดิมทั้งหมด")
                up=st.file_uploader(f"อัปโหลดทับ {filename}",type=["xlsx"],key=f"up_{filename}")
                if up:
                    try:
                        new_df=pd.read_excel(up); st.info(f"{len(new_df)} แถว, {len(new_df.columns)} คอลัมน์"); st.dataframe(new_df.head(3))
                        if st.button("✅ ยืนยันอัปโหลด",key=f"confirm_{filename}",type="primary"):
                            backup_excel(filename,df)
                            if write_excel_to_drive(filename,new_df,known_file_id=_fid_map.get(filename)):
                                st.toast("✅ อัปเดตสำเร็จ",icon="✅"); time.sleep(1); st.rerun()
                    except Exception as e: st.error(f"❌ อ่านไฟล์ไม่ได้: {e}")
        admin_file_panel(df_leave,FILE_LEAVE,tab1); admin_file_panel(df_travel,FILE_TRAVEL,tab2)
        admin_file_panel(read_attendance_report(),FILE_ATTEND,tab3); admin_file_panel(df_staff,FILE_STAFF,tab4)
        with tab5:
            st.subheader("🔧 ตั้งค่าและ Debug")
            st.info(f"FOLDER_ID: `{FOLDER_ID}`\nFILE_ATTEND: `{FILE_ATTEND}`")

            st.divider()
            st.subheader("🔍 Debug ไฟล์สแกนนิ้ว (attendance_report.xlsx)")
            st.caption("ใช้เพื่อตรวจสอบว่าโค้ดอ่านไฟล์ถูกต้องหรือไม่")

            if st.button("🔬 วิเคราะห์ไฟล์สแกนนิ้ว", key="btn_debug_att"):
                fid_att = get_file_id(FILE_ATTEND)
                if not fid_att:
                    st.error("❌ ไม่พบไฟล์ attendance_report.xlsx ใน Drive")
                else:
                    try:
                        req = get_drive_service().files().get_media(fileId=fid_att, supportsAllDrives=True)
                        fh2 = io.BytesIO()
                        dl2 = MediaIoBaseDownload(fh2, req)
                        done2 = False
                        while not done2: _, done2 = dl2.next_chunk()
                        fh2.seek(0)
                        df_debug = pd.read_excel(fh2, engine="openpyxl", header=0, dtype=str)
                        df_debug.columns = [str(c).strip() for c in df_debug.columns]

                        st.success(f"✅ อ่านไฟล์ได้: {len(df_debug)} แถว, {len(df_debug.columns)} คอลัมน์")

                        st.markdown("**📋 Column names ที่พบ:**")
                        cols_df = pd.DataFrame({"ลำดับ": range(1, len(df_debug.columns)+1), "ชื่อ Column": df_debug.columns.tolist()})
                        st.dataframe(cols_df, use_container_width=True, height=200)

                        st.markdown("**👀 ตัวอย่างข้อมูล 10 แถวแรก (raw):**")
                        st.dataframe(df_debug.head(10), use_container_width=True)

                        st.markdown("**🔄 ผลหลังผ่าน read_attendance_report():**")
                        read_attendance_report.clear()  # ล้างเฉพาะ attendance cache
                        df_parsed = read_attendance_report()
                        if df_parsed.empty:
                            st.error("❌ read_attendance_report() คืนค่าว่าง — column ไม่ตรงหรือข้อมูลผิดรูปแบบ")
                            st.markdown("**💡 วิธีแก้:** ตรวจสอบว่าไฟล์มี column ต่อไปนี้อย่างน้อย 1 ชื่อ:")
                            st.code("""
ชื่อพนักงาน / ชื่อ-สกุล / ชื่อ / Name / Employee Name
วันที่ / Date / Check Date / Attendance Date
เวลาเข้า / เข้า / Check In / Time In / First Check
เวลาออก / ออก / Check Out / Time Out / Last Check
                            """)
                        else:
                            st.success(f"✅ parse สำเร็จ: {len(df_parsed)} แถว")
                            st.dataframe(df_parsed.head(10), use_container_width=True)

                            # สรุปสถิติ
                            n_no_in  = len(df_parsed[df_parsed["เวลาเข้า"]==""])
                            n_no_out = len(df_parsed[df_parsed["เวลาออก"]==""])
                            n_both   = len(df_parsed[(df_parsed["เวลาเข้า"]!="")&(df_parsed["เวลาออก"]!="")])
                            st.markdown(f"""
**📊 สรุปคุณภาพข้อมูล:**
- มีทั้งเข้า+ออก: `{n_both}` แถว
- ไม่มีเวลาเข้า: `{n_no_in}` แถว
- ไม่มีเวลาออก: `{n_no_out}` แถว
- บุคลากรที่พบ: `{df_parsed["ชื่อ-สกุล"].nunique()}` คน
- ช่วงวันที่: `{df_parsed["วันที่"].min().strftime("%Y-%m-%d")}` ถึง `{df_parsed["วันที่"].max().strftime("%Y-%m-%d")}`
                            """)
                    except Exception as e:
                        st.error(f"❌ เกิดข้อผิดพลาด: {e}")
        with tab6:
            st.subheader("👆 บันทึกเวลาทำการสำหรับผู้ที่ลืมสแกนนิ้ว")
            df_manual_tab=_dc("cache_manual"); _manual_fid=get_file_id(FILE_MANUAL_SCAN)
            with st.form("form_manual_scan"):
                ms_name=st.selectbox("ชื่อ-สกุล *",get_active_staff(df_staff))
                ms_date=st.date_input("วันที่ลืมสแกน *",value=dt.date.today(),max_value=dt.date.today())
                c_t1,c_t2=st.columns(2)
                ms_time_in=c_t1.time_input("เวลาเข้างาน *",value=dt.time(8,30))
                ms_time_out=c_t2.time_input("เวลาออกงาน *",value=dt.time(16,30))
                if st.form_submit_button("💾 บันทึกข้อมูลสแกนนิ้ว",type="primary"):
                    new_row={"ชื่อ-สกุล":ms_name,"วันที่":pd.to_datetime(ms_date),"เวลาเข้า":ms_time_in.strftime("%H:%M"),"เวลาออก":ms_time_out.strftime("%H:%M"),"หมายเหตุ":f"Admin คีย์แทน — {dt.datetime.now().strftime('%d/%m/%Y %H:%M')}"}
                    df_manual_upd=pd.concat([df_manual_tab,pd.DataFrame([new_row])],ignore_index=True)
                    if write_excel_to_drive(FILE_MANUAL_SCAN,df_manual_upd,known_file_id=_manual_fid):
                        log_activity("คีย์สแกนนิ้ว",f"Admin คีย์ {ms_date} เข้า {ms_time_in.strftime('%H:%M')} ออก {ms_time_out.strftime('%H:%M')}",ms_name)
                        st.toast("✅ บันทึกสแกนนิ้วสำเร็จ",icon="✅"); time.sleep(1); st.rerun()
        with tab_hol:
            st.subheader("🎌 จัดการวันหยุดพิเศษ")
            df_hol_custom,_hol_fid=load_holidays_with_id()
            hol_yr_opts=list(range(dt.date.today().year+1,dt.date.today().year-3,-1))
            hol_view_year=st.selectbox("ดูปี (พ.ศ.)",[y+543 for y in hol_yr_opts])
            hol_view_year_ad=hol_view_year-543
            df_hol_display=load_holidays_all(hol_view_year_ad)
            if not df_hol_display.empty: st.dataframe(df_hol_display,use_container_width=True)
            st.divider(); st.subheader("➕ เพิ่มวันหยุดพิเศษ")
            with st.form("form_add_holiday"):
                h_col1,h_col2=st.columns(2)
                with h_col1: ha_date=st.date_input("วันที่",value=dt.date.today()); ha_name=st.text_input("ชื่อวันหยุด *")
                with h_col2: ha_type=st.selectbox("ประเภท",["วันหยุดราชการ","วันหยุดพิเศษ (ผนวก)","อื่นๆ"]); ha_note=st.text_input("หมายเหตุ")
                if st.form_submit_button("➕ เพิ่มวันหยุด",type="primary"):
                    if not ha_name.strip(): st.error("❌ กรุณาระบุชื่อวันหยุด")
                    else:
                        new_hol={"วันที่":pd.Timestamp(ha_date),"ชื่อวันหยุด":ha_name.strip(),"ประเภท":ha_type,"หมายเหตุ":ha_note.strip()}
                        df_hol_new=pd.concat([df_hol_custom,pd.DataFrame([new_hol])],ignore_index=True).sort_values("วันที่").reset_index(drop=True)
                        if write_excel_to_drive(FILE_HOLIDAYS,df_hol_new,known_file_id=_hol_fid):
                            st.toast(f"✅ เพิ่ม '{ha_name}' สำเร็จ",icon="🎌"); time.sleep(0.5); st.rerun()
