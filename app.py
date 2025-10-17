# ====================================================
# 📋 โปรแกรมติดตามการลาและไปราชการ (สคร.9)
# ✅ Final Version: Smart Cache, Refresh, Backup, Drive Access Fix
# ====================================================

import io
import os
import shutil
import datetime as dt
import altair as alt
import pandas as pd
import numpy as np
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# ===========================
# 🔐 Auth & App Config
# ===========================
st.set_page_config(page_title="สคร.9 - ติดตามการลา/ราชการ/สแกน", layout="wide")

# ✅ ใช้ scope ที่ถูกต้อง (แก้ไขแล้ว)
creds = service_account.Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=["https://www.googleapis.com/auth/drive"]
)
ADMIN_PASSWORD = st.secrets.get("admin_password", "admin123")

# ===========================
# 🗂️ Shared Drive Config
# ===========================
FOLDER_ID = "1YFJZvs59ahRHmlRrKcQwepWJz6A-4B7d"
FILE_ATTEND = "attendance_report.xlsx"
FILE_LEAVE  = "leave_report.xlsx"
FILE_TRAVEL = "travel_report.xlsx"

service = build("drive", "v3", credentials=creds)

# ===========================
# 🔧 Drive Helper Functions
# ===========================
def get_file_id(filename: str):
    """ค้นหา ID ของไฟล์ใน Shared Drive"""
    q = f"name='{filename}' and '{FOLDER_ID}' in parents and trashed=false"
    res = service.files().list(
        q=q,
        fields="files(id,name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()
    files = res.get("files", [])
    return files[0]["id"] if files else None


@st.cache_data(ttl=600)
def read_excel_from_drive(filename: str) -> pd.DataFrame:
    """อ่านไฟล์ Excel จาก Shared Drive"""
    try:
        file_id = get_file_id(filename)
        if not file_id:
            st.warning(f"⚠️ ไม่พบไฟล์ {filename} ใน Shared Drive")
            return pd.DataFrame()

        req = service.files().get_media(fileId=file_id, supportsAllDrives=True)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)
        st.sidebar.success(f"📄 โหลดไฟล์ {filename} สำเร็จจาก Drive")
        return pd.read_excel(fh, engine="openpyxl")
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์ {filename}: {e}")
        return pd.DataFrame()


def write_excel_to_drive(filename: str, df: pd.DataFrame):
    """เขียนไฟล์กลับขึ้น Shared Drive"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    media = MediaIoBaseUpload(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    file_id = get_file_id(filename)
    if file_id:
        service.files().update(fileId=file_id, media_body=media, supportsAllDrives=True).execute()
    else:
        service.files().create(
            body={"name": filename, "parents": [FOLDER_ID]},
            media_body=media,
            fields="id",
            supportsAllDrives=True
        ).execute()


def backup_excel(original_filename: str, df: pd.DataFrame):
    """สร้างไฟล์สำรองข้อมูลพร้อมประทับเวลา"""
    if df.empty:
        return
    now = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    backup_filename = f"backup_{now}_{original_filename}"
    st.sidebar.info(f"⚙️ กำลังสำรองข้อมูล: {backup_filename}")
    write_excel_to_drive(backup_filename, df)

# ====================================================
# ⚙️ Smart Cache + ปุ่มรีเฟรช + เวลา Sync
# ====================================================

LOCAL_CACHE_DIR = "cached_files"
os.makedirs(LOCAL_CACHE_DIR, exist_ok=True)

# 🔁 ปุ่มรีเฟรชข้อมูลจาก Drive
st.sidebar.markdown("---")
if st.sidebar.button("🔁 รีเฟรชข้อมูลจาก Drive (อัปเดตล่าสุด)"):
    try:
        shutil.rmtree(LOCAL_CACHE_DIR)
        os.makedirs(LOCAL_CACHE_DIR, exist_ok=True)
        st.sidebar.success("✅ ล้าง cache เรียบร้อย กำลังโหลดข้อมูลใหม่...")
        st.experimental_rerun()
    except Exception as e:
        st.sidebar.error(f"⚠️ ไม่สามารถล้าง cache ได้: {e}")


def update_sync_time():
    """อัปเดตเวลาซิงก์ล่าสุด"""
    sync_path = os.path.join(LOCAL_CACHE_DIR, "last_sync.txt")
    with open(sync_path, "w", encoding="utf-8") as f:
        f.write(dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))


def get_sync_time():
    """ดึงเวลาซิงก์ล่าสุด"""
    sync_path = os.path.join(LOCAL_CACHE_DIR, "last_sync.txt")
    if os.path.exists(sync_path):
        with open(sync_path, "r", encoding="utf-8") as f:
            return f.read().strip()
    return "— ยังไม่เคยซิงก์ข้อมูล —"


def load_excel_smart_cache(filename, from_drive=True):
    """
    โหลดไฟล์ Excel แบบอัจฉริยะ:
    1️⃣ ถ้ามีใน cache → โหลดทันที
    2️⃣ ถ้าไม่มี → ดึงจาก Drive
    3️⃣ ถ้าดึงสำเร็จ → เก็บ cache และบันทึกเวลาซิงก์
    """
    local_path = os.path.join(LOCAL_CACHE_DIR, filename)

    if os.path.exists(local_path):
        st.success(f"📄 โหลดไฟล์ {filename} จาก cache local ✅")
        return pd.read_excel(local_path)

    elif from_drive:
        st.info(f"🔄 ไม่พบ {filename} ในเครื่อง — กำลังดึงจาก Shared Drive...")
        df = read_excel_from_drive(filename)
        if df.empty:
            st.error(f"❌ ไม่พบไฟล์ {filename} ใน Shared Drive")
            return pd.DataFrame()

        try:
            df.to_excel(local_path, index=False)
            update_sync_time()
            st.success(f"✅ โหลดไฟล์ {filename} จาก Drive และบันทึก cache สำเร็จ")
        except Exception as e:
            st.warning(f"⚠️ โหลดสำเร็จแต่บันทึก cache ไม่ได้: {e}")
        return df

    else:
        st.error(f"❌ ไม่พบไฟล์ {filename} ทั้งใน local และ Shared Drive")
        return pd.DataFrame()


# ✅ โหลดข้อมูลทั้งสามชุด
df_att = load_excel_smart_cache(FILE_ATTEND)
df_leave = load_excel_smart_cache(FILE_LEAVE)
df_travel = load_excel_smart_cache(FILE_TRAVEL)

# 🕒 แสดงเวลาซิงก์ล่าสุดใน Sidebar
st.sidebar.caption(f"🕒 ซิงก์ข้อมูลล่าสุด: {get_sync_time()}")

# ====================================================
# 🎯 ส่วน UI หลัก (เหมือนเดิม)
# ====================================================
st.markdown("##### **สำนักงานป้องกันควบคุมโรคที่ 9 จังหวัดนครราชสีมา**")
st.title("📋 ระบบติดตามการลา ไปราชการ และการปฏิบัติงาน")

if 'submitted' not in st.session_state:
    st.session_state.submitted = False

def callback_submit():
    st.session_state.submitted = True

menu = st.sidebar.radio(
    "เลือกเมนู",
    ["หน้าหลัก", "📊 Dashboard", "📅 การมาปฏิบัติงาน", "🧭 การไปราชการ", "🕒 การลา", "🧑‍💼 ผู้ดูแลระบบ"]
)

if menu == "หน้าหลัก":
    st.info("💡 ระบบนี้ใช้สำหรับบันทึกข้อมูลการลา การไปราชการ และดูสรุปการปฏิบัติงานของบุคลากร สคร.9\n\n"
            "โปรดเลือกเมนูทางซ้ายเพื่อเริ่มต้นใช้งาน")
    st.image("https://ddc.moph.go.th/uploads/files/11120210817094038.jpg",
             caption="สำนักงานป้องกันควบคุมโรคที่ 9 นครราชสีมา")

# ส่วนอื่น (Dashboard, การมาปฏิบัติงาน, การลา, การไปราชการ, ผู้ดูแลระบบ)
# สามารถวางโค้ดเดิมของคุณต่อจากนี้ได้เลย
