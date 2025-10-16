# ====================================================
# 📋 โปรแกรมติดตามการลาและไปราชการ (สคร.9)
# ✅ Final Version: อัปเกรด read_excel_from_drive ให้ทนทานขึ้น
# ====================================================

import io
import mimetypes
import altair as alt
import datetime as dt
import pandas as pd
import numpy as np
import streamlit as st

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload, MediaFileUpload

# ===========================
# 🔐 Auth & App Config
# ===========================
st.set_page_config(page_title="สคร.9 - ติดตามการลา/ราชการ/สแกน", layout="wide")

creds = service_account.Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=["https://www.googleapis.com/auth/drive"]
)
ADMIN_PASSWORD = st.secrets.get("admin_password", "admin123")

# ===========================
# 🗂️ Shared Drive Config
# ===========================
FOLDER_ID = "1YFJZvs59ahRHmlRrKcQwepWJz6A-4B7d"
ATTACHMENT_FOLDER_NAME = "เอกสารแนบ_ไปราชการ"
FILE_ATTEND = "attendance_report.xlsx"
FILE_LEAVE  = "leave_report.xlsx"
FILE_TRAVEL = "travel_report.xlsx"

service = build("drive", "v3", credentials=creds)

# ===========================
# 🔧 Drive Helpers
# ===========================

# ‼️ --- ฟังก์ชันที่ได้รับการอัปเกรด --- ‼️
@st.cache_data(ttl=600)
def read_excel_from_drive(filename: str) -> pd.DataFrame:
    """อ่านไฟล์ Excel จาก Shared Drive; ถ้าไม่มีไฟล์ จะคืนค่า DataFrame ว่าง"""
    try:
        file_id = get_file_id(filename)
        # ✅ [แก้ไขสาเหตุที่ 1] แจ้งเตือนถ้าหาไฟล์ไม่เจอ
        if not file_id:
            st.warning(f"⚠️ ไม่พบไฟล์ '{filename}' ใน Google Drive กรุณาตรวจสอบชื่อไฟล์ให้ถูกต้อง")
            return pd.DataFrame()

        req = service.files().get_media(fileId=file_id, supportsAllDrives=True)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, req)
        done = False
        while not done: _, done = downloader.next_chunk()
        fh.seek(0)
        
        try:
            # ✅ [แก้ไขสาเหตุที่ 2] อ่านชีตแรกอัตโนมัติ ไม่ว่าชื่ออะไร
            xls = pd.ExcelFile(fh, engine="openpyxl")
            if not xls.sheet_names:
                st.error(f"ไฟล์ '{filename}' ไม่มีชีตข้อมูล")
                return pd.DataFrame()
            
            # อ่านชีตแรก
            df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])

            # ✅ [แก้ไขสาเหตุที่ 3] ตรวจสอบ Header อัตโนมัติ
            # ถ้าคอลัมน์ที่ควรจะมี (เช่น 'วันที่') ไม่อยู่ใน df ให้ลองอ่านใหม่โดยเริ่มที่แถวถัดไป
            expected_cols = ["วันที่", "ชื่อพนักงาน", "ชื่อ-สกุล"]
            if not any(col in df.columns for col in expected_cols):
                fh.seek(0) # ย้อนกลับไปอ่านไฟล์ใหม่
                df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=1)

            return df
        
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการอ่านโครงสร้างไฟล์ Excel '{filename}': {e}")
            return pd.DataFrame()

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดร้ายแรงในการเชื่อมต่อเพื่ออ่านไฟล์ {filename}: {e}")
        return pd.DataFrame()

def get_file_id(filename: str, parent_id=FOLDER_ID):
    """หา ID ของไฟล์หรือโฟลเดอร์ใน Parent ที่กำหนด"""
    q = f"name='{filename}' and '{parent_id}' in parents and trashed=false"
    res = service.files().list(q=q, fields="files(id,name)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    files = res.get("files", [])
    return files[0]["id"] if files else None

def write_excel_to_drive(filename: str, df: pd.DataFrame):
    # (โค้ดส่วนนี้เหมือนเดิม)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer: df.to_excel(writer, index=False)
    output.seek(0)
    media = MediaIoBaseUpload(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    file_id = get_file_id(filename)
    if file_id: service.files().update(fileId=file_id, media_body=media, supportsAllDrives=True).execute()
    else: service.files().create(body={"name": filename, "parents": [FOLDER_ID]}, media_body=media, fields="id", supportsAllDrives=True).execute()

def backup_excel(original_filename: str, df: pd.DataFrame):
    if df.empty: return
    now = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    backup_filename = f"backup_{now}_{original_filename}"
    write_excel_to_drive(backup_filename, df)

@st.cache_resource
def get_or_create_folder(folder_name, parent_folder_id):
    folder_id = get_file_id(folder_name, parent_id=parent_folder_id)
    if folder_id: return folder_id
    else:
        file_metadata = {'name': folder_name, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_folder_id]}
        folder = service.files().create(body=file_metadata, fields='id', supportsAllDrives=True).execute()
        return folder.get('id')

def upload_pdf_to_drive(file_object, filename, folder_id):
    if file_object is None: return "-"
    file_object.seek(0)
    media = MediaIoBaseUpload(file_object, mimetype='application/pdf', resumable=True)
    file_metadata = {'name': filename, 'parents': [folder_id]}
    file = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink', supportsAllDrives=True).execute()
    permission = {'type': 'anyone', 'role': 'reader'}
    service.permissions().create(fileId=file.get('id'), body=permission, supportsAllDrives=True).execute()
    return file.get('webViewLink')

def count_weekdays(start_date, end_date):
    if start_date and end_date and start_date <= end_date:
        return np.busday_count(start_date, end_date + dt.timedelta(days=1))
    return 0

# ===========================
# 📥 Load & Normalize Data
# ===========================
df_att = read_excel_from_drive(FILE_ATTEND)
df_leave = read_excel_from_drive(FILE_LEAVE)
df_travel = read_excel_from_drive(FILE_TRAVEL)

# ====================================================
# 🎯 UI Constants & Main App
# ====================================================
st.markdown("##### **สำนักงานป้องกันควบคุมโรคที่ 9 จังหวัดนครราชสีมา**")
st.title("📋 ระบบติดตามการลา ไปราชการ และการปฏิบัติงาน")

if 'submitted' not in st.session_state: st.session_state.submitted = False
def callback_submit(): st.session_state.submitted = True

all_names_leave = set(df_leave['ชื่อ-สกุล'].dropna()) if 'ชื่อ-สกุล' in df_leave.columns else set()
all_names_travel = set(df_travel['ชื่อ-สกุล'].dropna()) if 'ชื่อ-สกุล' in df_travel.columns else set()
name_col_att = next((col for col in ["ชื่อ-สกุล", "ชื่อพนักงาน", "ชื่อ"] if col in df_att.columns), None)
all_names_att = set(df_att[name_col_att].dropna()) if name_col_att else set()
all_names = sorted(all_names_leave.union(all_names_travel).union(all_names_att))

staff_groups = sorted(["กลุ่มโรคติดต่อ", "กลุ่มระบาดวิทยาฯ", "กลุ่มพัฒนาองค์กร", "กลุ่มบริหารทั่วไป", "กลุ่มโรคไม่ติดต่อ", "กลุ่มห้องปฏิบัติการฯ", "กลุ่มพัฒนานวัตกรรมฯ", "กลุ่มโรคติดต่อเรื้อรัง", "ศตม.9.1 ชัยภูมิ", "ศตม.9.2 บุรีรัมย์", "ศตม.9.3 สุรินทร์", "ศตม.9.4 ปากช่อง", "ด่านฯ ช่องจอม", "ศูนย์เวชศาสตร์ป้องกัน", "กลุ่มสื่อสารความเสี่ยง", "กลุ่มอาชีวสิ่งแวดล้อม"])
leave_types = ["ลาป่วย", "ลากิจ", "ลาพักผ่อน", "อื่นๆ"]

menu = st.sidebar.radio("เลือกเมนู", ["หน้าหลัก", "📊 Dashboard", "📅 การมาปฏิบัติงาน", "🧭 การไปราชการ", "🕒 การลา", "🧑‍💼 ผู้ดูแลระบบ"])

if menu == "หน้าหลัก":
    st.info("💡 ระบบนี้ใช้สำหรับบันทึกและสรุปข้อมูลบุคลากรใน สคร.9\n"
            "ได้แก่ การลา การไปราชการ และการมาปฏิบัติงาน พร้อมแนบไฟล์เอกสาร PDF ได้โดยตรง")
    st.image("https://ddc.moph.go.th/uploads/files/11120210817094038.jpg", caption="สำนักงานป้องกันควบคุมโรคที่ 9 นครราชสีมา", use_container_width=True)

elif menu == "📊 Dashboard":
    st.header("📊 Dashboard ภาพรวมและข้อมูลเชิงลึก")
    st.markdown("#### **ภาพรวมสะสม**")
    col1, col2, col3 = st.columns(3)
    col1.metric("เดินทางราชการ (ครั้ง)", len(df_travel))
    col2.metric("การลา (ครั้ง)", len(df_leave))
    col3.metric("ข้อมูลสแกน (แถว)", len(df_att))
    st.markdown("---")
    col_chart1, col_chart2 = st.columns(2)
    with col_chart1:
        st.markdown("##### **การลาแยกตามกลุ่มงาน**")
        if not df_leave.empty and 'กลุ่มงาน' in df_leave.columns and 'จำนวนวันลา' in df_leave.columns:
            leave_by_group = df_leave.groupby('กลุ่มงาน')['จำนวนวันลา'].sum().sort_values(ascending=False).reset_index()
            st.altair_chart(alt.Chart(leave_by_group).mark_bar().encode(x=alt.X('จำนวนวันลา:Q', title='รวมวันลา'), y=alt.Y('กลุ่มงาน:N', sort='-x', title='กลุ่มงาน'), tooltip=['กลุ่มงาน', 'จำนวนวันลา']).properties(height=300), use_container_width=True)
    with col_chart2:
        st.markdown("##### **ผู้เดินทางราชการบ่อยที่สุด (Top 5)**")
        if not df_travel.empty and 'ชื่อ-สกุล' in df_travel.columns:
            top_travelers = df_travel['ชื่อ-สกุล'].value_counts().nlargest(5).reset_index()
            top_travelers.columns = ['ชื่อ-สกุล', 'จำนวนครั้ง']
            st.altair_chart(alt.Chart(top_travelers).mark_bar(color='#ff8c00').encode(x=alt.X('จำนวนครั้ง:Q', title='จำนวนครั้ง'), y=alt.Y('ชื่อ-สกุล:N', sort='-x', title='ชื่อ-สกุล'), tooltip=['ชื่อ-สกุล', 'จำนวนครั้ง']).properties(height=300), use_container_width=True)

elif menu == "📅 การมาปฏิบัติงาน":
    st.header("📅 สรุปการมาปฏิบัติงานรายวัน (ตรวจจากสแกน + ลา + ราชการ)")
    # (โค้ดส่วนนี้เหมือนเดิมตามที่คุณให้มา)

elif menu == "🧭 การไปราชการ":
    # (โค้ดส่วนนี้เหมือนเดิม)

elif menu == "🕒 การลา":
    # (โค้ดส่วนนี้เหมือนเดิม)

elif menu == "🧑‍💼 ผู้ดูแลระบบ":
    # (โค้ดส่วนนี้เหมือนเดิม)
