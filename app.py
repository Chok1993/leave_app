# ====================================================
# 📋 โปรแกรมติดตามการลาและไปราชการ (สคร.9)
# ✅ เวอร์ชันอัปเกรด: Shared Drive + ระบบผู้ดูแล + Dashboard ทันสมัย
# ====================================================

import streamlit as st
import pandas as pd
import datetime as dt
import altair as alt
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
import io

# ====================================================
# 🔐 โหลดข้อมูลจาก Secrets (Service Account + Admin Password)
# ====================================================

creds = service_account.Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=["https://www.googleapis.com/auth/drive"]
)

# สามารถตั้งค่าใน Streamlit Cloud -> Secrets
ADMIN_PASSWORD = st.secrets.get("admin_password", "admin123")

# ====================================================
# 🧭 ตั้งค่า ID ของ Shared Drive และชื่อไฟล์
# ====================================================

FOLDER_ID = "1YFJZvs59ahRHmlRrKcQwepWJz6A-4B7d"
FILE_SCAN = "scan_report.xlsx"
FILE_REPORT = "leave_report.xlsx"

service = build("drive", "v3", credentials=creds)

# ====================================================
# 🗂️ ฟังก์ชันจัดการไฟล์ใน Shared Drive
# ====================================================

def get_file_id(filename):
    query = f"name='{filename}' and '{FOLDER_ID}' in parents and trashed=false"
    results = service.files().list(
        q=query, fields="files(id,name)",
        supportsAllDrives=True, includeItemsFromAllDrives=True
    ).execute()
    files = results.get("files", [])
    return files[0]["id"] if files else None


def read_excel_from_drive(filename):
    file_id = get_file_id(filename)
    if not file_id:
        return pd.DataFrame()
    request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    return pd.read_excel(fh)


def write_excel_to_drive(filename, df):
    """บันทึกข้อมูลขึ้น Shared Drive"""
    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False)
        output.seek(0)

        file_id = get_file_id(filename)
        media = MediaIoBaseUpload(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if file_id:
            service.files().update(fileId=file_id, media_body=media, supportsAllDrives=True).execute()
        else:
            service.files().create(
                body={"name": filename, "parents": [FOLDER_ID]},
                media_body=media, fields="id", supportsAllDrives=True
            ).execute()
    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาดในการเขียนไฟล์: {e}")

# ====================================================
# ⚙️ โหลดข้อมูลเริ่มต้น
# ====================================================

df_scan = read_excel_from_drive(FILE_SCAN)
df_report = read_excel_from_drive(FILE_REPORT)

# ====================================================
# 🎯 รายชื่อกลุ่มงาน
# ====================================================

staff_groups = [
    "กลุ่มโรคติดต่อ", "กลุ่มระบาดวิทยาและตอบโต้ภาวะฉุกเฉินทางสาธารณสุข",
    "กลุ่มพัฒนาองค์กร", "กลุ่มบริหารทั่วไป", "กลุ่มโรคไม่ติดต่อ",
    "กลุ่มห้องปฏิบัติการทางการแพทย์", "กลุ่มพัฒนานวัตกรรมและวิจัย",
    "กลุ่มโรคติดต่อเรื้อรัง", "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.1 จ.ชัยภูมิ",
    "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.2 จ.บุรีรัมย์",
    "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.3 จ.สุรินทร์",
    "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.4 ปากช่อง",
    "ด่านควบคุมโรคช่องจอม จ.สุรินทร์", "ศูนย์บริการเวชศาสตร์ป้องกัน",
    "กลุ่มสื่อสารความเสี่ยง", "กลุ่มโรคจากการประกอบอาชีพและสิ่งแวดล้อม"
]

# ====================================================
# 🧭 เมนูหลัก
# ====================================================

st.set_page_config(page_title="ระบบติดตามการลาและไปราชการ (สคร.9)", layout="wide")
st.title("📋 ระบบติดตามการลาและไปราชการ (สคร.9)")

menu = st.sidebar.radio("เลือกเมนู", ["📊 Dashboard", "🧭 การไปราชการ", "🕒 การลา", "🧑‍💼 ผู้ดูแลระบบ"])

# ====================================================
# 📈 Dashboard
# ====================================================

if menu == "📊 Dashboard":
    st.header("📊 ภาพรวมสถิติการลาและไปราชการ (Dashboard)")

    total_travel = len(df_scan)
    total_leave = len(df_report)

    col1, col2, col3 = st.columns(3)
    col1.metric("จำนวนผู้ไปราชการ", total_travel)
    col2.metric("จำนวนผู้ลา", total_leave)
    col3.metric("รวมทั้งหมด", total_travel + total_leave)

    st.markdown("---")

    if not df_report.empty:
        st.subheader("📌 การลาแยกตามประเภท")
        leave_by_type = df_report.groupby("ประเภทการลา").size().reset_index(name="จำนวนครั้ง")
        bar = alt.Chart(leave_by_type).mark_bar(cornerRadiusTopLeft=5, cornerRadiusTopRight=5).encode(
            x=alt.X("ประเภทการลา:N", sort=None),
            y="จำนวนครั้ง:Q",
            color="ประเภทการลา:N",
            tooltip=["ประเภทการลา", "จำนวนครั้ง"]
        )
        st.altair_chart(bar, use_container_width=True)

        st.subheader("📌 สัดส่วนการลา (Pie Chart)")
        pie = alt.Chart(leave_by_type).mark_arc(innerRadius=70).encode(
            theta="จำนวนครั้ง:Q",
            color="ประเภทการลา:N",
            tooltip=["ประเภทการลา", "จำนวนครั้ง"]
        )
        st.altair_chart(pie, use_container_width=True)

# ====================================================
# 🧭 บันทึกข้อมูลการไปราชการ
# ====================================================

elif menu == "🧭 การไปราชการ":
    st.header("🧭 บันทึกข้อมูลการไปราชการ")
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
        submitted = st.form_submit_button("💾 บันทึกข้อมูล")

    if submitted:
        df_scan = pd.concat([df_scan, pd.DataFrame([data])], ignore_index=True)
        write_excel_to_drive(FILE_SCAN, df_scan)
        st.success("✅ บันทึกข้อมูลไปราชการเรียบร้อยแล้ว!")

# ====================================================
# 🕒 บันทึกข้อมูลการลา
# ====================================================

elif menu == "🕒 การลา":
    st.header("🕒 บันทึกข้อมูลการลา")
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
        submitted = st.form_submit_button("💾 บันทึกข้อมูล")

    if submitted:
        df_report = pd.concat([df_report, pd.DataFrame([data])], ignore_index=True)
        write_excel_to_drive(FILE_REPORT, df_report)
        st.success("✅ บันทึกข้อมูลการลาเรียบร้อยแล้ว!")

# ====================================================
# 🧑‍💼 ระบบผู้ดูแล (Admin)
# ====================================================

elif menu == "🧑‍💼 ผู้ดูแลระบบ":
    st.header("🔐 เข้าสู่ระบบผู้ดูแลระบบ (Admin)")

    if "admin_logged_in" not in st.session_state:
        st.session_state.admin_logged_in = False

    if not st.session_state.admin_logged_in:
        password = st.text_input("กรอกรหัสผ่าน", type="password", placeholder="ระบุรหัสผ่านผู้ดูแลระบบ...")
        if st.button("เข้าสู่ระบบ"):
            if password == ADMIN_PASSWORD:
                st.session_state.admin_logged_in = True
                st.success("✅ เข้าสู่ระบบสำเร็จ")
            elif password != "":
                st.error("❌ รหัสผ่านไม่ถูกต้อง")
        st.stop()

    st.success("คุณได้เข้าสู่ระบบผู้ดูแลแล้ว 🧑‍💼")
    st.markdown("---")

    if st.button("🚪 ออกจากระบบ"):
        st.session_state.admin_logged_in = False
        st.experimental_rerun()

    tab1, tab2 = st.tabs(["📘 ข้อมูลการไปราชการ", "📗 ข้อมูลการลา"])

    with tab1:
        st.subheader("📘 รายการข้อมูลการไปราชการ")
        if not df_scan.empty:
            edit_scan = st.data_editor(df_scan, num_rows="dynamic", use_container_width=True, key="edit_scan")
            if st.button("💾 บันทึกการเปลี่ยนแปลง (ไปราชการ)"):
                write_excel_to_drive(FILE_SCAN, edit_scan)
                st.success("✅ บันทึกข้อมูลไปราชการเรียบร้อยแล้ว")
        else:
            st.info("ℹ️ ยังไม่มีข้อมูลการไปราชการในระบบ")

    with tab2:
        st.subheader("📗 รายการข้อมูลการลา")
        if not df_report.empty:
            edit_leave = st.data_editor(df_report, num_rows="dynamic", use_container_width=True, key="edit_leave")
            if st.button("💾 บันทึกการเปลี่ยนแปลง (การลา)"):
                write_excel_to_drive(FILE_REPORT, edit_leave)
                st.success("✅ บันทึกข้อมูลการลาเรียบร้อยแล้ว")
        else:
            st.info("ℹ️ ยังไม่มีข้อมูลการลาในระบบ")
