# ====================================================
# 📋 โปรแกรมติดตามการลาและไปราชการ (สคร.9)
# ✅ Final Version: Definitive Fix for All Errors
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
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

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
BACKUP_FOLDER = "Backups"
FILE_ATTEND = "scan_report.xlsx"
FILE_LEAVE  = "leave_report.xlsx"
FILE_TRAVEL = "travel_report.xlsx"

service = build("drive", "v3", credentials=creds)

# ===========================
# ⚙️ Helper Functions
# ===========================
def get_file_id(filename: str, parent_id=FOLDER_ID):
    q = f"name='{filename}' and '{parent_id}' in parents and trashed=false"
    res = service.files().list(
        q=q, fields="files(id,name)", supportsAllDrives=True, includeItemsFromAllDrives=True
    ).execute()
    files = res.get("files", [])
    return files[0]["id"] if files else None

def get_or_create_folder(folder_name, parent_folder_id):
    q = f"name='{folder_name}' and '{parent_folder_id}' in parents and trashed=false"
    res = service.files().list(
        q=q, fields="files(id)", supportsAllDrives=True, includeItemsFromAllDrives=True
    ).execute()
    if res.get("files"):
        return res["files"][0]["id"]
    folder_metadata = {
        "name": folder_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_folder_id]
    }
    folder = service.files().create(body=folder_metadata, fields="id", supportsAllDrives=True).execute()
    return folder.get("id")

@st.cache_data(ttl=600)
def read_excel_from_drive(filename: str) -> pd.DataFrame:
    """อ่านไฟล์ Excel จาก Shared Drive"""
    try:
        file_id = get_file_id(filename)
        if not file_id:
            st.warning(f"⚠️ ไม่พบไฟล์ {filename}")
            return pd.DataFrame()
        req = service.files().get_media(fileId=file_id, supportsAllDrives=True)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)
        return pd.read_excel(fh, engine="openpyxl")
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์ {filename}: {e}")
        return pd.DataFrame()

def write_excel_to_drive(filename: str, df: pd.DataFrame):
    """เขียนไฟล์ Excel กลับไปยัง Shared Drive"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    media = MediaIoBaseUpload(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    file_id = get_file_id(filename)
    if file_id:
        service.files().update(fileId=file_id, media_body=media, supportsAllDrives=True).execute()
    else:
        service.files().create(
            body={"name": filename, "parents": [FOLDER_ID]},
            media_body=media, fields="id", supportsAllDrives=True
        ).execute()

def backup_excel(original_filename: str, df: pd.DataFrame):
    """สร้างไฟล์สำรองก่อนเขียนทับ"""
    if df.empty:
        return
    now = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    backup_folder_id = get_or_create_folder(BACKUP_FOLDER, FOLDER_ID)
    backup_filename = f"backup_{now}_{original_filename}"
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    media = MediaIoBaseUpload(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    service.files().create(
        body={"name": backup_filename, "parents": [backup_folder_id]},
        media_body=media, fields="id", supportsAllDrives=True
    ).execute()

# ===========================
# 📥 Load All Data
# ===========================
df_att = read_excel_from_drive(FILE_ATTEND)
df_leave = read_excel_from_drive(FILE_LEAVE)
df_travel = read_excel_from_drive(FILE_TRAVEL)

def safe_get_names(df, cols):
    for col in cols:
        if col in df.columns:
            return set(df[col].dropna().astype(str))
    return set()

all_names = sorted(set().union(
    safe_get_names(df_att, ["ชื่อ-สกุล", "ชื่อพนักงาน", "ชื่อ"]),
    safe_get_names(df_leave, ["ชื่อ-สกุล", "ชื่อพนักงาน", "ชื่อ"]),
    safe_get_names(df_travel, ["ชื่อ-สกุล", "ชื่อพนักงาน", "ชื่อ"])
))

# ===========================
# 🧭 เมนูหลัก
# ===========================
menu = st.sidebar.radio(
    "เลือกเมนู",
    ["📅 การมาปฏิบัติงาน", "📊 Dashboard", "🧭 การไปราชการ", "🕒 การลา", "🧑‍💼 ผู้ดูแลระบบ"]
)

# ===========================
# 📅 การมาปฏิบัติงาน
# ===========================
if menu == "📅 การมาปฏิบัติงาน":
    st.header("📅 สรุปการมาปฏิบัติงานรายวัน/รายเดือน")
    if df_att.empty:
        st.warning("⚠️ ไม่พบข้อมูลในไฟล์ scan_report.xlsx")
        st.stop()

    df_att["วันที่"] = pd.to_datetime(df_att["วันที่"], errors="coerce")
    df_att["เดือน"] = df_att["วันที่"].dt.strftime("%Y-%m")

    months = sorted(df_att["เดือน"].dropna().unique())
    selected_month = st.selectbox("เลือกเดือน", months, index=len(months)-1 if months else 0)
    selected_names = st.multiselect("เลือกบุคลากร (ปล่อยว่างเพื่อดูทั้งหมด)", all_names)

    df_month = df_att[df_att["เดือน"] == selected_month]

    records = []
    for name in (selected_names or all_names):
        for d in pd.date_range(selected_month + "-01", periods=31, freq="D"):
            if d.month != pd.to_datetime(selected_month).month:
                break
            rec = {"ชื่อพนักงาน": name, "วันที่": d.date(), "สถานะ": ""}
            att = df_month[(df_month["ชื่อพนักงาน"] == name) & (df_month["วันที่"].dt.date == d.date())]
            leave = df_leave[(df_leave["ชื่อ-สกุล"] == name) & (df_leave["วันที่เริ่ม"] <= d) & (df_leave["วันที่สิ้นสุด"] >= d)]
            travel = df_travel[(df_travel["ชื่อ-สกุล"] == name) & (df_travel["วันที่เริ่ม"] <= d) & (df_travel["วันที่สิ้นสุด"] >= d)]

            if not leave.empty:
                rec["สถานะ"] = f"ลา ({leave['ประเภทการลา'].iloc[0]})"
            elif not travel.empty:
                rec["สถานะ"] = "ไปราชการ"
            elif not att.empty:
                rec["สถานะ"] = "มาปกติ"
            else:
                rec["สถานะ"] = "ขาดงาน"
            records.append(rec)

    df_daily = pd.DataFrame(records)
    st.dataframe(df_daily, use_container_width=True)

    summary = df_daily.pivot_table(index="ชื่อพนักงาน", columns="สถานะ", aggfunc="size", fill_value=0).reset_index()
    st.markdown("### 📊 สรุปสถิติรวมต่อเดือน")
    st.dataframe(summary, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_daily.to_excel(writer, index=False, sheet_name="รายวัน")
        summary.to_excel(writer, index=False, sheet_name="สรุปเดือน")
    output.seek(0)
    st.download_button("📥 ดาวน์โหลดรายงาน (Excel)", output,
                       file_name=f"รายงานสรุป_{selected_month}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ===========================
# 📊 Dashboard
# ===========================
elif menu == "📊 Dashboard":
    st.header("📊 Dashboard ภาพรวมข้อมูล")
    col1, col2, col3 = st.columns(3)
    col1.metric("การลา (ครั้ง)", len(df_leave))
    col2.metric("ไปราชการ (ครั้ง)", len(df_travel))
    col3.metric("ข้อมูลสแกน (แถว)", len(df_att))

# ===========================
# 🧑‍💼 ผู้ดูแลระบบ
# ===========================
elif menu == "🧑‍💼 ผู้ดูแลระบบ":
    st.header("🔐 โหมดผู้ดูแลระบบ")

    if "admin_logged_in" not in st.session_state:
        st.session_state.admin_logged_in = False

    if not st.session_state.admin_logged_in:
        pwd = st.text_input("กรอกรหัสผ่าน", type="password")
        if st.button("เข้าสู่ระบบ"):
            if pwd == ADMIN_PASSWORD:
                st.session_state.admin_logged_in = True
                st.rerun()
            else:
                st.error("❌ รหัสผ่านไม่ถูกต้อง")
        st.stop()

    st.success("✅ เข้าสู่ระบบเรียบร้อยแล้ว")
    if st.button("🚪 ออกจากระบบ"):
        st.session_state.admin_logged_in = False
        st.rerun()

    tab1, tab2, tab3 = st.tabs(["📗 การลา", "📘 ไปราชการ", "🟩 สแกนเข้า-ออก"])

    with tab1:
        st.caption("แก้ไขข้อมูลการลาได้โดยตรง แล้วกดบันทึก")
        edited_leave = st.data_editor(df_leave.astype(str), num_rows="dynamic", use_container_width=True)
        if st.button("💾 บันทึกข้อมูลการลา"):
            backup_excel(FILE_LEAVE, df_leave)
            write_excel_to_drive(FILE_LEAVE, edited_leave)
            st.success("✅ บันทึกข้อมูลการลาเรียบร้อย")

    with tab2:
        st.caption("แก้ไขข้อมูลไปราชการได้โดยตรง แล้วกดบันทึก")
        edited_travel = st.data_editor(df_travel.astype(str), num_rows="dynamic", use_container_width=True)
        if st.button("💾 บันทึกข้อมูลไปราชการ"):
            backup_excel(FILE_TRAVEL, df_travel)
            write_excel_to_drive(FILE_TRAVEL, edited_travel)
            st.success("✅ บันทึกข้อมูลไปราชการเรียบร้อย")

    with tab3:
        st.caption("แก้ไขข้อมูลสแกนได้โดยตรง แล้วกดบันทึก")
        edited_att = st.data_editor(df_att.astype(str), num_rows="dynamic", use_container_width=True)
        if st.button("💾 บันทึกข้อมูลสแกน"):
            backup_excel(FILE_ATTEND, df_att)
            write_excel_to_drive(FILE_ATTEND, edited_att)
            st.success("✅ บันทึกข้อมูลสแกนเรียบร้อย")
