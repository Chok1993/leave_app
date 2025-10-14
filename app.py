# ====================================================
# โปรแกรมบันทึกข้อมูลการลา / การไปราชการ (สคร.9)
# ✅ เชื่อมต่อกับ Google Drive โดยใช้ Service Account (สำหรับ Streamlit Cloud)
# ====================================================

import streamlit as st
import pandas as pd
import datetime as dt
import altair as alt
from fpdf import FPDF
import matplotlib.pyplot as plt
from io import BytesIO
import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from googleapiclient.errors import HttpError
import io

# ====================================================
# 🔐 โหลดข้อมูลจาก Secrets (service_account.json จาก Streamlit Cloud)
# ====================================================

creds = service_account.Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=["https://www.googleapis.com/auth/drive"]
)

# ====================================================
# 🧭 ตั้งค่า ID ของโฟลเดอร์ Google Drive
# ====================================================

FOLDER_ID = "1YjoU7QqbMgCIf547HlTq5rzvUK5gXhUI"  # แก้ให้ตรงกับโฟลเดอร์จริงของคุณ
service = build('drive', 'v3', credentials=creds)

# ====================================================
# 🗂️ ฟังก์ชันสำหรับอ่าน/เขียนไฟล์ Excel ใน Google Drive
# ====================================================

def get_file_id(filename):
    """ค้นหาไฟล์ใน Google Drive ตามชื่อ"""
    try:
        results = service.files().list(
            q=f"name='{filename}' and '{FOLDER_ID}' in parents and trashed=false",
            fields="files(id, name)"
        ).execute()
        files = results.get("files", [])
        return files[0]["id"] if files else None
    except HttpError as e:
        st.error(f"❌ ไม่สามารถเข้าถึง Google Drive ได้: {e}")
        return None


def read_excel_from_drive(filename):
    """อ่านไฟล์ Excel จาก Google Drive"""
    file_id = get_file_id(filename)
    if not file_id:
        return pd.DataFrame()
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    fh.seek(0)
    return pd.read_excel(fh)


def write_excel_to_drive(filename, df):
    """บันทึก DataFrame กลับขึ้น Google Drive"""
    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        output.seek(0)

        file_id = get_file_id(filename)
        media = MediaIoBaseUpload(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        if file_id:
            service.files().update(fileId=file_id, media_body=media).execute()
        else:
            file_metadata = {"name": filename, "parents": [FOLDER_ID]}
            service.files().create(body=file_metadata, media_body=media, fields="id").execute()

    except HttpError as e:
        st.error(f"❌ เกิดข้อผิดพลาดในการเขียนไฟล์: {e}")

# ====================================================
# ⚙️ โหลดข้อมูลเริ่มต้น
# ====================================================

FILE_SCAN = "scan_report.xlsx"
FILE_REPORT = "leave_report.xlsx"

df_scan = read_excel_from_drive(FILE_SCAN)
df_report = read_excel_from_drive(FILE_REPORT)

# ====================================================
# 🎯 รายชื่อกลุ่มงาน
# ====================================================

staff_groups = [
    "กลุ่มโรคติดต่อ", "กลุ่มระบาดวิทยาและตอบโต้ภาวะฉุกเฉินทางสาธารณสุข",
    "กลุ่มพัฒนาองค์กร", "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.1 จ.ชัยภูมิ",
    "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.2 จ.บุรีรัมย์", "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.3 จ.สุรินทร์",
    "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.4 ปากช่อง", "ด่านควบคุมโรคช่องจอม จ.สุรินทร์",
    "กลุ่มโรคไม่ติดต่อ", "งานควบคุมโรคเขตเมือง", "กลุ่มโรคติดต่อเรื้อรัง",
    "กลุ่มห้องปฏิบัติการทางการแพทย์", "กลุ่มสื่อสารความเสี่ยง", "กลุ่มโรคจากการประกอบอาชีพและสิ่งแวดล้อม",
    "ศูนย์บริการเวชศาสตร์ป้องกัน", "กลุ่มบริหารทั่วไป", "กลุ่มพัฒนานวัตกรรมและวิจัย"
]

# ====================================================
# 🧭 เมนูหลัก
# ====================================================

st.sidebar.title("เลือกเมนู")
menu = st.sidebar.radio("", ["📊 Dashboard", "🧭 การไปราชการ", "🕒 การลา", "🧑‍💼 ผู้ดูแลระบบ"])

st.title("📋 ระบบติดตามการลาและไปราชการ (สคร.9)")

# ====================================================
# 📈 Dashboard
# ====================================================

if menu == "📊 Dashboard":
    st.header("ภาพรวมสถิติ")

    total_travel = len(df_scan)
    total_leave = len(df_report)

    col1, col2 = st.columns(2)
    col1.metric("จำนวนผู้ไปราชการ", total_travel)
    col2.metric("จำนวนผู้ลา", total_leave)

    if not df_report.empty:
        leave_by_type = df_report.groupby("ประเภทการลา")["จำนวนวันลา"].sum().reset_index()
        leave_by_type["ลำดับ"] = range(1, len(leave_by_type) + 1)
        st.dataframe(leave_by_type)
        st.altair_chart(
            alt.Chart(leave_by_type).mark_bar().encode(
                x="ประเภทการลา", y="จำนวนวันลา", color="ประเภทการลา"
            ),
            use_container_width=True
        )

# ====================================================
# 🧭 ฟอร์มการไปราชการ
# ====================================================

elif menu == "🧭 การไปราชการ":
    st.header("บันทึกข้อมูลการไปราชการ")
    with st.form("form_scan"):
        data = {
            "ลำดับ": len(df_scan) + 1,
            "ชื่อ-สกุล": st.text_input("ชื่อ-สกุล"),
            "กลุ่มงาน": st.selectbox("กลุ่มงาน", staff_groups),
            "กิจกรรม": st.text_input("กิจกรรม"),
            "สถานที่": st.text_input("สถานที่"),
            "วันที่เริ่ม": st.date_input("วันที่เริ่ม", dt.date.today()),
            "วันที่สิ้นสุด": st.date_input("วันที่สิ้นสุด", dt.date.today())
        }
        data["จำนวนวัน"] = (data["วันที่สิ้นสุด"] - data["วันที่เริ่ม"]).days + 1
        submitted = st.form_submit_button("บันทึกข้อมูล")

    if submitted:
        df_scan = pd.concat([df_scan, pd.DataFrame([data])], ignore_index=True)
        write_excel_to_drive(FILE_SCAN, df_scan)
        st.success("✅ บันทึกข้อมูลเรียบร้อยแล้ว!")

# ====================================================
# 🕒 ฟอร์มการลา
# ====================================================

elif menu == "🕒 การลา":
    st.header("บันทึกข้อมูลการลา")
    with st.form("form_leave"):
        data = {
            "ลำดับ": len(df_report) + 1,
            "ชื่อ-สกุล": st.text_input("ชื่อ-สกุล"),
            "กลุ่มงาน": st.selectbox("กลุ่มงาน", staff_groups),
            "ประเภทการลา": st.selectbox("ประเภทการลา", ["ลาป่วย", "ลากิจ", "ลาพักผ่อน", "อื่นๆ"]),
            "วันที่เริ่ม": st.date_input("วันที่เริ่ม", dt.date.today()),
            "วันที่สิ้นสุด": st.date_input("วันที่สิ้นสุด", dt.date.today())
        }
        data["จำนวนวันลา"] = (data["วันที่สิ้นสุด"] - data["วันที่เริ่ม"]).days + 1
        submitted = st.form_submit_button("บันทึกข้อมูล")

    if submitted:
        df_report = pd.concat([df_report, pd.DataFrame([data])], ignore_index=True)
        write_excel_to_drive(FILE_REPORT, df_report)
        st.success("✅ บันทึกข้อมูลเรียบร้อยแล้ว!")

# ====================================================
# 👩‍💼 ผู้ดูแลระบบ
# ====================================================

elif menu == "🧑‍💼 ผู้ดูแลระบบ":
    st.header("📦 จัดการข้อมูล")
    st.write("คุณสามารถดาวน์โหลดข้อมูลปัจจุบันทั้งหมดได้ที่นี่ 👇")

    col1, col2 = st.columns(2)
    with col1:
        if not df_scan.empty:
            st.download_button("📥 ดาวน์โหลดข้อมูลไปราชการ", df_scan.to_csv(index=False).encode('utf-8-sig'), "scan_report.csv")
    with col2:
        if not df_report.empty:
            st.download_button("📥 ดาวน์โหลดข้อมูลการลา", df_report.to_csv(index=False).encode('utf-8-sig'), "leave_report.csv")
