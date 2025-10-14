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

FOLDER_ID = "1YjoU7QqbMgCIf547HlTq5rzvUK5gXhUI"  # ⚙️ แก้ให้ตรงกับโฟลเดอร์จริงของคุณ

# สร้างบริการ Drive API
service = build('drive', 'v3', credentials=creds)

# ====================================================
# 🗂️ ฟังก์ชันอ่าน/เขียนไฟล์ Excel ใน Google Drive
# ====================================================

def get_file_id(filename):
    """ค้นหาไฟล์ใน Google Drive ตามชื่อ"""
    results = service.files().list(
        q=f"name='{filename}' and '{FOLDER_ID}' in parents and trashed=false",
        fields="files(id, name)"
    ).execute()
    files = results.get("files", [])
    return files[0]["id"] if files else None


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

st.title("📋 ระบบติดตามการลาและไปราชการ (สคร.9)")
menu = st.sidebar.radio("เลือกเมนู", ["📊 Dashboard", "🧭 การไปราชการ", "🕒 การลา", "🛠️ ผู้ดูแลระบบ"])

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
# 🛠️ ระบบผู้ดูแล (Admin) - แก้ไข/ลบเฉพาะรายการ
# ====================================================

elif menu == "🛠️ ผู้ดูแลระบบ":
    st.header("🔐 เข้าสู่ระบบผู้ดูแล")

    ADMIN_PASSWORD = "12345"  # 🔑 เปลี่ยนรหัสผ่านได้ที่นี่
    password = st.text_input("กรอกรหัสผ่าน", type="password")

    if password == ADMIN_PASSWORD:
        st.success("เข้าสู่ระบบเรียบร้อย ✅")

        tab1, tab2 = st.tabs(["📄 ข้อมูลการไปราชการ", "📋 ข้อมูลการลา"])

        # ============ ข้อมูลการไปราชการ ============
        with tab1:
            st.subheader("📄 รายการข้อมูลการไปราชการ")

            if not df_scan.empty:
                df_scan_display = df_scan.reset_index().rename(columns={"index": "ลำดับ"})
                df_scan_display["ลำดับ"] = df_scan_display["ลำดับ"] + 1  # ✅ เริ่มนับจาก 1
                st.dataframe(df_scan_display, use_container_width=True)

                selected_index = st.number_input(
                    "เลือกหมายเลขลำดับที่ต้องการแก้ไข/ลบ",
                    min_value=1, max_value=len(df_scan_display), step=1
                ) - 1

                edit_mode = st.radio("เลือกการดำเนินการ", ["✏️ แก้ไขข้อมูล", "🗑️ ลบรายการ"])

                if edit_mode == "✏️ แก้ไขข้อมูล":
                    st.write("กรอกข้อมูลใหม่สำหรับรายการนี้:")
                    updated = {}
                    updated["ชื่อ-สกุล"] = st.text_input("ชื่อ-สกุล", df_scan.loc[selected_index, "ชื่อ-สกุล"])
                    updated["กลุ่มงาน"] = st.selectbox("กลุ่มงาน", staff_groups,
                                                       index=staff_groups.index(df_scan.loc[selected_index, "กลุ่มงาน"]))
                    updated["กิจกรรม"] = st.text_input("กิจกรรม", df_scan.loc[selected_index, "กิจกรรม"])
                    updated["สถานที่"] = st.text_input("สถานที่", df_scan.loc[selected_index, "สถานที่"])
                    updated["วันที่เริ่ม"] = st.date_input("วันที่เริ่ม", df_scan.loc[selected_index, "วันที่เริ่ม"])
                    updated["วันที่สิ้นสุด"] = st.date_input("วันที่สิ้นสุด", df_scan.loc[selected_index, "วันที่สิ้นสุด"])
                    updated["จำนวนวัน"] = (updated["วันที่สิ้นสุด"] - updated["วันที่เริ่ม"]).days + 1

                    if st.button("💾 บันทึกการแก้ไข"):
                        df_scan.loc[selected_index] = updated
                        write_excel_to_drive(FILE_SCAN, df_scan)
                        st.success("✅ แก้ไขข้อมูลเรียบร้อยแล้ว!")

                elif edit_mode == "🗑️ ลบรายการ":
                    if st.button("🗑️ ยืนยันการลบ"):
                        df_scan = df_scan.drop(selected_index).reset_index(drop=True)
                        write_excel_to_drive(FILE_SCAN, df_scan)
                        st.warning("⚠️ ลบข้อมูลเรียบร้อยแล้ว")

            else:
                st.info("ไม่มีข้อมูลไปราชการในระบบ")

        # ============ ข้อมูลการลา ============
        with tab2:
            st.subheader("📋 รายการข้อมูลการลา")

            if not df_report.empty:
                df_report_display = df_report.reset_index().rename(columns={"index": "ลำดับ"})
                df_report_display["ลำดับ"] = df_report_display["ลำดับ"] + 1  # ✅ เริ่มนับจาก 1
                st.dataframe(df_report_display, use_container_width=True)

                selected_index = st.number_input(
                    "เลือกหมายเลขลำดับที่ต้องการแก้ไข/ลบ (การลา)",
                    min_value=1, max_value=len(df_report_display), step=1
                ) - 1

                edit_mode = st.radio("เลือกการดำเนินการ (การลา)", ["✏️ แก้ไขข้อมูล", "🗑️ ลบรายการ"])

                if edit_mode == "✏️ แก้ไขข้อมูล":
                    st.write("กรอกข้อมูลใหม่สำหรับรายการนี้:")
                    updated = {}
                    updated["ชื่อ-สกุล"] = st.text_input("ชื่อ-สกุล", df_report.loc[selected_index, "ชื่อ-สกุล"])
                    updated["กลุ่มงาน"] = st.selectbox("กลุ่มงาน", staff_groups,
                                                       index=staff_groups.index(df_report.loc[selected_index, "กลุ่มงาน"]))
                    updated["ประเภทการลา"] = st.selectbox(
                        "ประเภทการลา", ["ลาป่วย", "ลากิจ", "ลาพักผ่อน", "อื่นๆ"],
                        index=["ลาป่วย", "ลากิจ", "ลาพักผ่อน", "อื่นๆ"].index(df_report.loc[selected_index, "ประเภทการลา"])
                    )
                    updated["วันที่เริ่ม"] = st.date_input("วันที่เริ่ม", df_report.loc[selected_index, "วันที่เริ่ม"])
                    updated["วันที่สิ้นสุด"] = st.date_input("วันที่สิ้นสุด", df_report.loc[selected_index, "วันที่สิ้นสุด"])
                    updated["จำนวนวันลา"] = (updated["วันที่สิ้นสุด"] - updated["วันที่เริ่ม"]).days + 1

                    if st.button("💾 บันทึกการแก้ไข (การลา)"):
                        df_report.loc[selected_index] = updated
                        write_excel_to_drive(FILE_REPORT, df_report)
                        st.success("✅ แก้ไขข้อมูลเรียบร้อยแล้ว!")

                elif edit_mode == "🗑️ ลบรายการ":
                    if st.button("🗑️ ยืนยันการลบ (การลา)"):
                        df_report = df_report.drop(selected_index).reset_index(drop=True)
                        write_excel_to_drive(FILE_REPORT, df_report)
                        st.warning("⚠️ ลบข้อมูลเรียบร้อยแล้ว")

            else:
                st.info("ไม่มีข้อมูลการลาในระบบ")

    elif password != "":
        st.error("❌ รหัสผ่านไม่ถูกต้อง")
