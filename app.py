# ====================================================
# 📋 โปรแกรมติดตามการลาและไปราชการ (สคร.9)
# ✅ Final Version: นับเฉพาะวันทำการ (จันทร์-ศุกร์)
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

def backup_excel(original_filename: str, df: pd.DataFrame):
    if df.empty: return
    now = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    backup_filename = f"backup_{now}_{original_filename}"
    write_excel_to_drive(backup_filename, df)

# ⭐ ฟังก์ชันสำหรับนับวันทำการ (จันทร์-ศุกร์)
def count_weekdays(start_date, end_date):
    """นับจำนวนวันทำการ (จันทร์-ศุกร์) ระหว่างวันที่สองวัน แบบรวมวันสิ้นสุด"""
    if start_date and end_date and start_date <= end_date:
        # np.busday_count จะไม่นับวันสุดท้าย จึงต้องบวกไป 1 วันเพื่อให้ครอบคลุม
        return np.busday_count(start_date, end_date + dt.timedelta(days=1))
    return 0

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

# ====================================================
# 🎯 UI Constants & Main App
# ====================================================
st.markdown("##### **สำนักงานป้องกันควบคุมโรคที่ 9 จังหวัดนครราชสีมา**")
st.title("📋 ระบบติดตามการลา ไปราชการ และการปฏิบัติงาน")

if 'submitted' not in st.session_state:
    st.session_state.submitted = False

def callback_submit():
    st.session_state.submitted = True

all_names_leave = set(df_leave['ชื่อ-สกุล'].dropna()) if 'ชื่อ-สกุล' in df_leave.columns else set()
all_names_travel = set(df_travel['ชื่อ-สกุล'].dropna()) if 'ชื่อ-สกุล' in df_travel.columns else set()
all_names_att = set(df_att['ชื่อ-สกุล'].dropna()) if 'ชื่อ-สกุล' in df_att.columns else set()
all_names = sorted(all_names_leave.union(all_names_travel).union(all_names_att))

staff_groups = sorted([
    "กลุ่มโรคติดต่อ", "กลุ่มระบาดวิทยาและตอบโต้ภาวะฉุกเฉินทางสาธารณสุข", "กลุ่มพัฒนาองค์กร", "กลุ่มบริหารทั่วไป", "กลุ่มโรคไม่ติดต่อ",
    "กลุ่มห้องปฏิบัติการทางการแพทย์", "กลุ่มพัฒนานวัตกรรมและวิจัย", "กลุ่มโรคติดต่อเรื้อรัง", "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.1 จ.ชัยภูมิ",
    "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.2 จ.บุรีรัมย์", "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.3 จ.สุรินทร์",
    "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.4 ปากช่อง", "ด่านควบคุมโรคช่องจอม จ.สุรินทร์", "ศูนย์บริการเวชศาสตร์ป้องกัน",
    "กลุ่มสื่อสารความเสี่ยง", "กลุ่มโรคจากการประกอบอาชีพและสิ่งแวดล้อม"
])
leave_types = ["ลาป่วย", "ลากิจ", "ลาพักผ่อน", "อื่นๆ"]

menu = st.sidebar.radio("เลือกเมนู", ["หน้าหลัก", "📊 Dashboard", "📅 การมาปฏิบัติงาน", "🧭 การไปราชการ", "🕒 การลา", "🧑‍💼 ผู้ดูแลระบบ"])

if menu == "หน้าหลัก":
    st.info("💡 ระบบนี้ใช้สำหรับบันทึกข้อมูลการลา การไปราชการ และดูสรุปการปฏิบัติงานของบุคลากร สคร.9\n\n"
            "โปรดเลือกเมนูทางซ้ายเพื่อเริ่มต้นใช้งาน")
    st.image("https://ddc.moph.go.th/uploads/files/11120210817094038.jpg", caption="สคร.9 นครราชสีมา")

elif menu == "📊 Dashboard":
    st.header("📊 Dashboard ภาพรวมและข้อมูลเชิงลึก")
    # (โค้ดส่วน Dashboard เหมือนเดิม)

elif menu == "📅 การมาปฏิบัติงาน":
    st.header("📅 สรุปการมาปฏิบัติงานรายเดือน")
    # (โค้ดส่วนนี้เหมือนเดิม)

elif menu == "🧭 การไปราชการ":
    st.header("🧭 บันทึกการไปราชการ (สำหรับหมู่คณะ)")
    with st.form("form_travel_group"):
        common_data_ui = {
            "กลุ่มงาน": st.selectbox("กลุ่มงาน", staff_groups, disabled=st.session_state.submitted),
            "กิจกรรม": st.text_input("กิจกรรม/โครงการ", disabled=st.session_state.submitted),
            "สถานที่": st.text_input("สถานที่", disabled=st.session_state.submitted),
            "วันที่เริ่ม": st.date_input("วันที่เริ่ม", dt.date.today(), key="travel_start_date", disabled=st.session_state.submitted),
            "วันที่สิ้นสุด": st.date_input("วันที่สิ้นสุด", dt.date.today(), key="travel_end_date", disabled=st.session_state.submitted)
        }
        
        # ⭐ อัปเดตการแสดงผลให้ใช้ฟังก์ชันใหม่
        days = count_weekdays(st.session_state.travel_start_date, st.session_state.travel_end_date)
        if days > 0:
            st.caption(f"🗓️ รวมเฉพาะวันทำการ {days} วัน")
        
        selected_names = st.multiselect("เลือกชื่อเจ้าหน้าที่ (เลือกได้หลายคน)", options=all_names, disabled=st.session_state.submitted)
        new_names_str = st.text_area("เพิ่มชื่อเจ้าหน้าที่ใหม่ (กรณีไม่มีในตัวเลือก)", placeholder="ใส่ 1 ชื่อต่อ 1 บรรทัด", disabled=st.session_state.submitted)
        
        submitted_travel = st.form_submit_button("💾 บันทึกข้อมูล", on_click=callback_submit, disabled=st.session_state.submitted)

    if submitted_travel:
        new_names = [name.strip() for name in new_names_str.split('\n') if name.strip()]
        final_names = list(set(selected_names + new_names))
        if not final_names:
            st.warning("กรุณาเลือกหรือกรอก 'ชื่อ-สกุล' อย่างน้อย 1 คน")
            st.session_state.submitted = False
        elif common_data_ui["วันที่เริ่ม"] > common_data_ui["วันที่สิ้นสุด"]:
            st.error("'วันที่เริ่ม' ต้องมาก่อน 'วันที่สิ้นสุด'")
            st.session_state.submitted = False
        else:
            with st.spinner('⏳ กำลังบันทึกข้อมูล... กรุณารอสักครู่'):
                backup_excel(FILE_TRAVEL, df_travel)
                new_records = []
                timestamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
                
                # ⭐ อัปเดตการคำนวณวันให้ใช้ฟังก์ชันใหม่
                num_days = count_weekdays(common_data_ui["วันที่เริ่ม"], common_data_ui["วันที่สิ้นสุด"])
                
                for name in final_names:
                    fellow_travelers = ", ".join([other for other in final_names if other != name])
                    record = {**common_data_ui, "ชื่อ-สกุล": name, "จำนวนวัน": num_days, "ผู้ร่วมเดินทาง": fellow_travelers if fellow_travelers else "-", "last_update": timestamp}
                    new_records.append(record)
                if new_records:
                    df_travel_new = pd.concat([df_travel, pd.DataFrame(new_records)], ignore_index=True)
                    write_excel_to_drive(FILE_TRAVEL, df_travel_new)
                    st.success(f"✅ บันทึกข้อมูลไปราชการของเจ้าหน้าที่ {len(final_names)} คนเรียบร้อยแล้ว!")
                    st.session_state.submitted = False
                    st.rerun()

    st.markdown("---")
    st.markdown("### 🔍 ค้นหาข้อมูลรายบุคคล")
    search_name_travel = st.text_input("พิมพ์ชื่อ-สกุลเพื่อค้นหา (ไปราชการ)", "")
    if search_name_travel:
        df_filtered_travel = df_travel[df_travel['ชื่อ-สกุล'].str.contains(search_name_travel, case=False, na=False)]
        st.dataframe(df_filtered_travel.astype(str))
    else:
        st.markdown("### 📋 ข้อมูลปัจจุบันทั้งหมด")
        st.dataframe(df_travel.astype(str).sort_values('วันที่เริ่ม', ascending=False))

elif menu == "🕒 การลา":
    st.header("🕒 บันทึกข้อมูลการลา")

    with st.form("form_leave"):
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("ชื่อ-สกุล", disabled=st.session_state.submitted, help="กรอกชื่อและนามสกุลเต็ม")
            start_date = st.date_input("วันที่เริ่มลา", dt.date.today(), key="leave_start_date", disabled=st.session_state.submitted)
        with col2:
            group = st.selectbox("กลุ่มงาน", staff_groups, disabled=st.session_state.submitted)
            end_date = st.date_input("วันที่สิ้นสุดการลา", dt.date.today(), key="leave_end_date", disabled=st.session_state.submitted)
        
        leave_type = st.selectbox("ประเภทการลา", leave_types, disabled=st.session_state.submitted)

        # ⭐ อัปเดตการแสดงผลให้ใช้ฟังก์ชันใหม่
        days = count_weekdays(st.session_state.leave_start_date, st.session_state.leave_end_date)
        if days > 0:
            st.caption(f"🗓️ รวมเฉพาะวันทำการ {days} วัน")

        submitted_leave = st.form_submit_button("💾 บันทึกข้อมูล", on_click=callback_submit, disabled=st.session_state.submitted)

    if submitted_leave:
        data = {"ชื่อ-สกุล": name, "กลุ่มงาน": group, "ประเภทการลา": leave_type, "วันที่เริ่ม": start_date, "วันที่สิ้นสุด": end_date}
        if not data["ชื่อ-สกุล"]:
            st.warning("กรุณากรอก 'ชื่อ-สกุล'")
            st.session_state.submitted = False
        elif data["วันที่เริ่ม"] > data["วันที่สิ้นสุด"]:
            st.error("'วันที่เริ่ม' ต้องมาก่อน 'วันที่สิ้นสุด'")
            st.session_state.submitted = False
        else:
            with st.spinner('⏳ กำลังบันทึกข้อมูล...'):
                backup_excel(FILE_LEAVE, df_leave)
                
                # ⭐ อัปเดตการคำนวณวันให้ใช้ฟังก์ชันใหม่
                data["จำนวนวันลา"] = count_weekdays(data["วันที่เริ่ม"], data["วันที่สิ้นสุด"])
                data["last_update"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
                
                df_leave_new = pd.concat([df_leave, pd.DataFrame([data])], ignore_index=True)
                write_excel_to_drive(FILE_LEAVE, df_leave_new)
                
                st.success("✅ บันทึกข้อมูลการลาเรียบร้อยแล้ว!")
                st.session_state.submitted = False
                st.rerun()

    st.markdown("---")
    st.markdown("### 🔍 ค้นหาข้อมูลรายบุคคล")
    search_name_leave = st.text_input("พิมพ์ชื่อ-สกุลเพื่อค้นหา (การลา)", "")
    if search_name_leave:
        df_filtered_leave = df_leave[df_leave['ชื่อ-สกุล'].str.contains(search_name_leave, case=False, na=False)]
        st.dataframe(df_filtered_leave.astype(str))
    else:
        st.markdown("### 📋 ข้อมูลปัจจุบันทั้งหมด")
        st.dataframe(df_leave.astype(str).sort_values('วันที่เริ่ม', ascending=False))

elif menu == "🧑‍💼 ผู้ดูแลระบบ":
    st.header("🔐 ผู้ดูแลระบบ")
    # (โค้ดส่วน Admin เหมือนเดิม)
