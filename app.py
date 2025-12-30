# ====================================================
# 📋 ระบบติดตามการลา/ไปราชการ สคร.9 (ฉบับสมบูรณ์)
# 🛠️ Feature: Smart Date Unified System (ป้องกันวันที่ผิดพลาด)
# ====================================================

import io
import datetime as dt
import pandas as pd
import numpy as np
import altair as alt
import streamlit as st

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# ===========================
# 🔐 1. Configuration & Auth
# ===========================
st.set_page_config(page_title="สคร.9 - ระบบติดตามงาน", layout="wide")

if "gcp_service_account" not in st.secrets:
    st.error("❌ ไม่พบการตั้งค่า Google Cloud Credentials ใน Secrets")
    st.stop()

# เชื่อมต่อ Google Drive API
creds = service_account.Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=["https://www.googleapis.com/auth/drive"]
)
service = build("drive", "v3", credentials=creds)

FOLDER_ID = "1YFJZvs59ahRHmlRrKcQwepWJz6A-4B7d"
FILE_ATTEND = "attendance_report.xlsx"
FILE_LEAVE  = "leave_report.xlsx"
FILE_TRAVEL = "travel_report.xlsx"
ATTACHMENT_FOLDER_NAME = "Attachments_Leave_App"
ADMIN_PASSWORD = st.secrets.get("admin_password", "admin123")

# ===========================
# 🔧 2. Smart Date Engine (หัวใจหลักการแก้ไข)
# ===========================

def smart_parse_date(val):
    """
    แปลงค่าวันที่จากทุกรูปแบบ (พ.ศ., ค.ศ., string, datetime) 
    ให้เป็นวัตถุวันที่มาตรฐานของ Python
    """
    if pd.isna(val) or str(val).strip() == "":
        return pd.NaT
    
    s = str(val).strip()
    
    # 1. ลองแปลงจาก Format ต่างๆ
    for fmt in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"]:
        try:
            d = dt.datetime.strptime(s, fmt)
            # ถ้าปี > 2500 แสดงว่าเป็น พ.ศ. ให้แปลงเป็น ค.ศ.
            if d.year > 2500:
                d = d.replace(year=d.year - 543)
            return d
        except ValueError:
            continue
            
    # 2. ถ้ายังไม่ได้ ให้ใช้ pandas ช่วยเดา
    try:
        d = pd.to_datetime(s, errors='coerce')
        if d is not pd.NaT and d.year > 2500:
            d = d.replace(year=d.year - 543)
        return d
    except:
        return pd.NaT

def format_date_for_ui(d):
    """แปลงวันที่เป็น string รูปแบบ วว/ดด/พ.ศ. สำหรับแสดงผลบนหน้าจอ"""
    if pd.isna(d): return "-"
    return f"{d.day:02d}/{d.month:02d}/{d.year + 543}"

# ===========================
# 📂 3. Drive & Data Functions
# ===========================

def get_file_id(filename):
    q = f"name='{filename}' and '{FOLDER_ID}' in parents and trashed=false"
    res = service.files().list(q=q, fields="files(id)", supportsAllDrives=True).execute()
    files = res.get("files", [])
    return files[0]["id"] if files else None

def read_data(filename):
    fid = get_file_id(filename)
    if not fid: return pd.DataFrame()
    
    req = service.files().get_media(fileId=fid, supportsAllDrives=True)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, req)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    
    df = pd.read_excel(fh, engine="openpyxl")
    
    # ทำความสะอาดวันที่ทันทีที่โหลด
    date_cols = ["วันที่เริ่ม", "วันที่สิ้นสุด", "วันที่", "Timestamp", "last_update"]
    for col in date_cols:
        if col in df.columns:
            df[col] = df[col].apply(smart_parse_date).dt.normalize()
    
    return df

def save_data(filename, df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    
    fid = get_file_id(filename)
    media = MediaIoBaseUpload(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    if fid:
        service.files().update(fileId=fid, media_body=media, supportsAllDrives=True).execute()
    else:
        meta = {"name": filename, "parents": [FOLDER_ID]}
        service.files().create(body=meta, media_body=media, supportsAllDrives=True).execute()
    st.cache_data.clear()

# ===========================
# 🚀 4. Main Interface Logic
# ===========================

df_leave = read_data(FILE_LEAVE)
df_travel = read_data(FILE_TRAVEL)
df_att = read_data(FILE_ATTEND)

# รวมรายชื่อบุคลากร (Clean name)
def get_unique_names():
    names = set()
    if not df_leave.empty: names.update(df_leave["ชื่อ-สกุล"].dropna().unique())
    if not df_travel.empty: names.update(df_travel["ชื่อ-สกุล"].dropna().unique())
    if not df_att.empty:
        col = next((c for c in ["ชื่อพนักงาน", "ชื่อ-สกุล"] if c in df_att.columns), None)
        if col: names.update(df_att[col].dropna().unique())
    return sorted([str(n).strip() for n in names if str(n).strip() != 'nan'])

ALL_NAMES = get_unique_names()

st.sidebar.title("MENU")
menu = st.sidebar.radio("ไปที่หน้า:", ["📊 แดชบอร์ด", "🕒 บันทึกการลา", "🧭 บันทึกราชการ", "📅 ตารางปฏิบัติงาน", "🔒 ผู้ดูแลระบบ"])

# ---------------------------
# 📊 แดชบอร์ด
# ---------------------------
if menu == "📊 แดชบอร์ด":
    st.header("ภาพรวมข้อมูลบุคลากร")
    c1, c2, c3 = st.columns(3)
    c1.metric("รายการลาทั้งหมด", len(df_leave))
    c2.metric("รายการไปราชการ", len(df_travel))
    c3.metric("ข้อมูลสแกนนิ้ว", len(df_att))
    
    st.divider()
    if not df_leave.empty:
        st.subheader("สถิติการลาแยกตามประเภท")
        chart = alt.Chart(df_leave).mark_bar().encode(
            x=alt.X('count():Q', title='จำนวนครั้ง'),
            y=alt.Y('ประเภทการลา:N', sort='-x'),
            color='ประเภทการลา:N'
        ).properties(height=300)
        st.altair_chart(chart, use_container_width=True)

# ---------------------------
# 🕒 บันทึกการลา (แก้ไขการเลือกชื่อและวันที่)
# ---------------------------
elif menu == "🕒 บันทึกการลา":
    st.header("📝 บันทึกใบลา")
    with st.form("leave_form"):
        col1, col2 = st.columns(2)
        with col1:
            name = st.selectbox("ชื่อ-สกุล", ALL_NAMES)
            l_type = st.selectbox("ประเภทการลา", ["ลาป่วย", "ลากิจส่วนตัว", "ลาพักผ่อน", "ลาคลอด", "อื่นๆ"])
        with col2:
            start_d = st.date_input("เริ่มวันที่")
            end_d = st.date_input("ถึงวันที่")
        
        reason = st.text_input("เหตุผล (ถ้ามี)")
        submit = st.form_submit_button("บันทึกข้อมูล")
        
        if submit:
            new_data = {
                "ลำดับ": len(df_leave) + 1,
                "ชื่อ-สกุล": name,
                "ประเภทการลา": l_type,
                "วันที่เริ่ม": pd.to_datetime(start_d),
                "วันที่สิ้นสุด": pd.to_datetime(end_d),
                "จำนวนวันลา": (end_d - start_d).days + 1,
                "เหตุผล": reason,
                "last_update": dt.datetime.now()
            }
            df_updated = pd.concat([df_leave, pd.DataFrame([new_data])], ignore_index=True)
            save_data(FILE_LEAVE, df_updated)
            st.success(f"บันทึกข้อมูลของ {name} สำเร็จ")
            st.rerun()

    st.subheader("ประวัติการลาล่าสุด")
    # แสดงผลวันที่แบบ พ.ศ. ในตาราง
    if not df_leave.empty:
        view_df = df_leave.copy()
        view_df["วันที่เริ่ม"] = view_df["วันที่เริ่ม"].apply(format_date_for_ui)
        view_df["วันที่สิ้นสุด"] = view_df["วันที่สิ้นสุด"].apply(format_date_for_ui)
        st.dataframe(view_df.tail(10), use_container_width=True)

# ---------------------------
# 📅 ตารางปฏิบัติงาน (หัวใจของการเช็คชื่อ)
# ---------------------------
elif menu == "📅 ตารางปฏิบัติงาน":
    st.header("📅 ตรวจสอบสถานะปฏิบัติงานรายเดือน")
    
    col_a, col_b = st.columns(2)
    with col_a:
        sel_person = st.selectbox("เลือกชื่อบุคลากร", ALL_NAMES)
    with col_b:
        # ดึงรายเดือนที่มีข้อมูลสแกน
        if not df_att.empty:
            df_att['month_year'] = df_att['วันที่'].dt.strftime('%m/%Y')
            months = df_att['month_year'].unique()
            sel_month = st.selectbox("เลือกเดือน/ปี", months)
        else:
            st.warning("ไม่มีข้อมูลสแกนในระบบ")
            st.stop()

    if sel_person and sel_month:
        m, y = map(int, sel_month.split('/'))
        num_days = 31 # หรือใช้ calendar.monthrange เพื่อความแม่นยำ
        date_range = pd.date_range(start=f"{y}-{m}-01", end=pd.Timestamp(y, m, 1) + pd.offsets.MonthEnd(0))
        
        report = []
        for d in date_range:
            status = "มาทำงาน"
            # เช็ควันหยุด
            if d.weekday() >= 5: status = "วันหยุดเสาร์-อาทิตย์"
            
            # เช็คลา (ใช้ smart date ที่แปลงแล้ว)
            if not df_leave.empty:
                is_l = df_leave[(df_leave["ชื่อ-สกุล"] == sel_person) & 
                                (df_leave["วันที่เริ่ม"] <= d) & 
                                (df_leave["วันที่สิ้นสุด"] >= d)]
                if not is_l.empty:
                    status = f"ลา ({is_l.iloc[0]['ประเภทการลา']})"
            
            # เช็คราชการ
            if not df_travel.empty:
                is_t = df_travel[(df_travel["ชื่อ-สกุล"] == sel_person) & 
                                 (df_travel["วันที่เริ่ม"] <= d) & 
                                 (df_travel["วันที่สิ้นสุด"] >= d)]
                if not is_t.empty:
                    status = "ไปราชการ"

            report.append({"วันที่": format_date_for_ui(d), "สถานะ": status})
        
        st.table(pd.DataFrame(report))

# ---------------------------
# 🔒 ผู้ดูแลระบบ (จัดการไฟล์ดิบ)
# ---------------------------
elif menu == "🔒 ผู้ดูแลระบบ":
    st.header("Admin Management")
    pw = st.text_input("รหัสผ่าน", type="password")
    if pw == ADMIN_PASSWORD:
        st.success("เข้าสู่ระบบจัดการไฟล์")
        
        file_to_manage = st.selectbox("เลือกไฟล์ที่ต้องการจัดการ", [FILE_LEAVE, FILE_TRAVEL, FILE_ATTEND])
        current_df = read_data(file_to_manage)
        
        st.write(f"ข้อมูลปัจจุบันใน {file_to_manage}:")
        st.dataframe(current_df.head(10))
        
        st.divider()
        st.warning("⚠️ การอัปโหลดไฟล์ใหม่จะลบข้อมูลเดิมในไฟล์นั้นทั้งหมด")
        uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel ใหม่ (.xlsx)", type="xlsx")
        if uploaded_file:
            if st.button("ยืนยันการเขียนทับข้อมูล"):
                new_df = pd.read_excel(uploaded_file)
                save_data(file_to_manage, new_df)
                st.success("อัปเดตข้อมูลเรียบร้อยแล้ว")
                st.rerun()
    elif pw:
        st.error("รหัสผ่านไม่ถูกต้อง")
