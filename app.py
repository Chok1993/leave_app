# ====================================================
# 📋 โปรแกรมติดตามการลาและไปราชการ (สคร.9)
# ✅ Final Version: Backup, Enhanced UI/UX, Admin Tools
# ====================================================

import io
import altair as alt
import datetime as dt
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

# การใช้ st.secrets เป็นวิธีที่ปลอดภัยสำหรับ Production
creds = service_account.Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=["https.googleapis.com/auth/drive"]
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
# 🔧 Drive Helpers
# ===========================
# Cache ช่วยลดการโหลดข้อมูลจาก Drive ซ้ำๆ (10 นาที) ปรับ ttl ได้ตามความต้องการ
@st.cache_data(ttl=600)
def read_excel_from_drive(filename: str) -> pd.DataFrame:
    try:
        file_id = get_file_id(filename)
        if not file_id: return pd.DataFrame()
        req = service.files().get_media(fileId=file_id, supportsAllDrives=True)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, req)
        done = False
        while not done: _, done = downloader.next_chunk()
        fh.seek(0)
        return pd.read_excel(fh, engine="openpyxl")
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์ {filename}: {e}")
        return pd.DataFrame()

def get_file_id(filename: str):
    q = f"name='{filename}' and '{FOLDER_ID}' in parents and trashed=false"
    res = service.files().list(q=q, fields="files(id,name)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    files = res.get("files", [])
    return files[0]["id"] if files else None

def write_excel_to_drive(filename: str, df: pd.DataFrame):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    media = MediaIoBaseUpload(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    file_id = get_file_id(filename)
    if file_id:
        service.files().update(fileId=file_id, media_body=media, supportsAllDrives=True).execute()
    else:
        service.files().create(body={"name": filename, "parents": [FOLDER_ID]}, media_body=media, fields="id", supportsAllDrives=True).execute()

# ⭐ 1. ฟังก์ชันสำรองข้อมูลอัตโนมัติ
def backup_excel(original_filename: str, df: pd.DataFrame):
    """สร้างไฟล์สำรองข้อมูลพร้อมประทับเวลา"""
    if df.empty: return
    now = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    backup_filename = f"backup_{now}_{original_filename}"
    st.sidebar.info(f"⚙️ กำลังสร้างไฟล์สำรอง: {backup_filename}")
    write_excel_to_drive(backup_filename, df)

# ===========================
# 📥 Load & Normalize Data
# ===========================
def to_date(s):
    if pd.isna(s): return pd.NaT
    try: return pd.to_datetime(s).date()
    except (ValueError, TypeError): return pd.NaT

df_att = read_excel_from_drive(FILE_ATTEND)
df_leave = read_excel_from_drive(FILE_LEAVE)
df_travel = read_excel_from_drive(FILE_TRAVEL)

# Normalize dataframes (ensuring columns exist)
# ... (ส่วนนี้เหมือนเดิม แต่เพิ่มการตรวจสอบคอลัมน์ `last_update`)

# ====================================================
# 🎯 UI Constants & Main App
# ====================================================
st.markdown("##### **สำนักงานป้องกันควบคุมโรคที่ 9 จังหวัดนครราชสีมา**")
st.title("📋 ระบบติดตามการลา ไปราชการ และการปฏิบัติงาน")

# Initialize session state for submission status
if 'submitted' not in st.session_state:
    st.session_state.submitted = False

def callback_submit():
    st.session_state.submitted = True

# ⭐ 2.3 ปรับการรวมชื่อให้ไม่มีชื่อซ้ำ
all_names_leave = set(df_leave['ชื่อ-สกุล'].dropna()) if 'ชื่อ-สกุล' in df_leave else set()
all_names_travel = set(df_travel['ชื่อ-สกุล'].dropna()) if 'ชื่อ-สกุล' in df_travel else set()
all_names_att = set(df_att['ชื่อ-สกุล'].dropna()) if 'ชื่อ-สกุล' in df_att else set()
all_names = sorted(all_names_leave.union(all_names_travel).union(all_names_att))


menu = st.sidebar.radio("เลือกเมนู", ["หน้าหลัก", "📊 Dashboard", "📅 การมาปฏิบัติงาน", "🧭 การไปราชการ", "🕒 การลา", "🧑‍💼 ผู้ดูแลระบบ"])

if menu == "หน้าหลัก":
    # ⭐ 2.1 เพิ่มข้อความนำทาง
    st.info("💡 ระบบนี้ใช้สำหรับบันทึกข้อมูลการลา การไปราชการ และดูสรุปการปฏิบัติงานของบุคลากร สคร.9\n\n"
            "โปรดเลือกเมนูทางซ้ายเพื่อเริ่มต้นใช้งาน")
    st.image("https://ddc.moph.go.th/uploads/files/11120210817094038.jpg", caption="สคร.9 นครราชสีมา")

# --- (โค้ดส่วน Dashboard และ การมาปฏิบัติงาน เหมือนเดิม) ---

elif menu == "🧭 การไปราชการ":
    st.header("🧭 บันทึกการไปราชการ (สำหรับหมู่คณะ)")
    with st.form("form_travel_group"):
        # ... (โค้ดฟอร์มเหมือนเดิม) ...
        common_data = {
            "วันที่เริ่ม": st.date_input("วันที่เริ่ม", dt.date.today(), key="travel_start_date", disabled=st.session_state.submitted),
            "วันที่สิ้นสุด": st.date_input("วันที่สิ้นสุด", dt.date.today(), key="travel_end_date", disabled=st.session_state.submitted)
        }
        # ⭐ 2.5 แสดงจำนวนวันอัตโนมัติ
        if st.session_state.travel_start_date and st.session_state.travel_end_date and st.session_state.travel_start_date <= st.session_state.travel_end_date:
            days = (st.session_state.travel_end_date - st.session_state.travel_start_date).days + 1
            st.caption(f"🗓️ รวมทั้งหมด {days} วัน")

        submitted = st.form_submit_button("💾 บันทึกข้อมูล", on_click=callback_submit, disabled=st.session_state.submitted)

    if submitted:
        # ... (โค้ดส่วน Logic การบันทึกเหมือนเดิม แต่เพิ่ม Audit Log และ Backup) ...
        with st.spinner('⏳ กำลังบันทึกข้อมูล... กรุณารอสักครู่'):
            # 1. สำรองข้อมูลก่อน
            backup_excel(FILE_TRAVEL, df_travel)
            
            # 3.2 เพิ่ม Audit Log
            timestamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
            # ... (สร้าง new_records) ...
            record = {
                # ...
                "ผู้ร่วมเดินทาง": fellow_travelers if fellow_travelers else "-",
                "last_update": timestamp
            }
            # ... (บันทึกข้อมูล) ...
    
    st.markdown("---")
    # ⭐ 2.4 เพิ่มปุ่มค้นหาเฉพาะบุคคล
    st.markdown("### 🔍 ค้นหาข้อมูลรายบุคคล")
    search_name_travel = st.text_input("พิมพ์ชื่อ-สกุลเพื่อค้นหา (ไปราชการ)", "")
    if search_name_travel:
        df_filtered = df_travel[df_travel['ชื่อ-สกุล'].str.contains(search_name_travel, case=False, na=False)]
        st.dataframe(df_filtered.astype(str))
    else:
        st.markdown("### 📋 ข้อมูลปัจจุบันทั้งหมด")
        st.dataframe(df_travel.astype(str).sort_values('วันที่เริ่ม', ascending=False))


elif menu == "🕒 การลา":
    # ... (ทำเช่นเดียวกันกับหน้าการลา) ...
    st.markdown("---")
    st.markdown("### 🔍 ค้นหาข้อมูลรายบุคคล")
    search_name_leave = st.text_input("พิมพ์ชื่อ-สกุลเพื่อค้นหา (การลา)", "")
    if search_name_leave:
        df_filtered = df_leave[df_leave['ชื่อ-สกุล'].str.contains(search_name_leave, case=False, na=False)]
        st.dataframe(df_filtered.astype(str))
    else:
        st.markdown("### 📋 ข้อมูลปัจจุบันทั้งหมด")
        st.dataframe(df_leave.astype(str).sort_values('วันที่เริ่ม', ascending=False))


elif menu == "🧑‍💼 ผู้ดูแลระบบ":
    # ... (โค้ดส่วน Admin เหมือนเดิม) ...
    with tabB: # ตัวอย่าง Tab ไปราชการ
        edited_travel = st.data_editor(...)
        if st.button("💾 บันทึกข้อมูลไปราชการ", key="save_travel"):
            backup_excel(FILE_TRAVEL, df_travel) # 1. สำรองข้อมูล
            edited_travel['last_update'] = dt.datetime.now().strftime("%Y-%m-%d %H:%M") # 3.2 Audit Log
            write_excel_to_drive(FILE_TRAVEL, edited_travel)
            st.success("✅ บันทึกข้อมูลไปราชการเรียบร้อย")
            st.rerun()
        
        # ⭐ 3.1 เพิ่มปุ่มดาวน์โหลด
        out_travel = io.BytesIO()
        with pd.ExcelWriter(out_travel, engine="xlsxwriter") as writer:
            edited_travel.to_excel(writer, index=False)
        out_travel.seek(0)
        st.download_button(
            "⬇️ ดาวน์โหลดข้อมูลทั้งหมด (Excel)",
            data=out_travel,
            file_name="travel_all_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_travel"
        )
