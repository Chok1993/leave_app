# ====================================================
# 📋 โปรแกรมติดตามการลาและไปราชการ (สคร.9)
# ✅ Production Ready: ป้องกันการกดซ้ำ + เพิ่มข้อมูลผู้ร่วมเดินทาง
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

# เชื่อมต่อ Google API โดยใช้ข้อมูลจาก Streamlit Secrets
creds = service_account.Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=["https://www.googleapis.com/auth/drive"]
)
ADMIN_PASSWORD = st.secrets.get("admin_password", "admin123") # รหัสผ่าน Admin

# ===========================
# 🗂️ Shared Drive Config
# ===========================
FOLDER_ID = "1YFJZvs59ahRHmlRrKcQwepWJz6A-4B7d"

# ชื่อไฟล์มาตรฐาน
FILE_ATTEND = "attendance_report.xlsx"
FILE_LEAVE  = "leave_report.xlsx"
FILE_TRAVEL = "travel_report.xlsx"

service = build("drive", "v3", credentials=creds)

# ===========================
# 🔧 Drive Helpers
# ===========================
@st.cache_data(ttl=600)
def read_excel_from_drive(filename: str) -> pd.DataFrame:
    """อ่านไฟล์ Excel จาก Shared Drive; ถ้าไม่มีไฟล์ จะคืนค่า DataFrame ว่าง"""
    try:
        file_id = get_file_id(filename)
        if not file_id:
            return pd.DataFrame()
        req = service.files().get_media(fileId=file_id, supportsAllDrives=True)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)
        try:
            return pd.read_excel(fh)
        except Exception:
            fh.seek(0)
            return pd.read_excel(fh, engine="openpyxl")
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์ {filename}: {e}")
        return pd.DataFrame()


def get_file_id(filename: str):
    """หา file ID ในโฟลเดอร์เป้าหมายบน Google Drive"""
    q = f"name='{filename}' and '{FOLDER_ID}' in parents and trashed=false"
    res = service.files().list(
        q=q,
        fields="files(id,name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()
    files = res.get("files", [])
    return files[0]["id"] if files else None


def write_excel_to_drive(filename: str, df: pd.DataFrame):
    """บันทึก DataFrame กลับไปยังไฟล์ Excel บน Shared Drive"""
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
        service.files().update(
            fileId=file_id, media_body=media, supportsAllDrives=True
        ).execute()
    else:
        service.files().create(
            body={"name": filename, "parents": [FOLDER_ID]},
            media_body=media,
            fields="id",
            supportsAllDrives=True
        ).execute()

# ===========================
# 📥 Load & Normalize Data
# ===========================
def to_date(s):
    if pd.isna(s): return pd.NaT
    try:
        return pd.to_datetime(s).date()
    except (ValueError, TypeError):
        return pd.NaT

def to_time(s):
    if pd.isna(s): return None
    try:
        return pd.to_datetime(str(s)).time()
    except (ValueError, TypeError):
        return None

df_att = read_excel_from_drive(FILE_ATTEND)
if not df_att.empty:
    name_col = 'ชื่อพนักงาน' if 'ชื่อพนักงาน' in df_att.columns else 'ชื่อ-สกุล'
    if name_col in df_att.columns:
        df_att = df_att.rename(columns={name_col: 'ชื่อ-สกุล'})
        df_att['วันที่'] = df_att['วันที่'].apply(to_date)
        for c in ['เวลาเข้า', 'เวลาออก']:
            if c in df_att.columns:
                df_att[c] = df_att[c].apply(to_time).astype(str).replace('NaT', '')
else:
    df_att = pd.DataFrame(columns=['ชื่อ-สกุล','วันที่','เวลาเข้า','เวลาออก','สาย','ออกก่อน','หมายเหตุ'])

df_leave = read_excel_from_drive(FILE_LEAVE)
if not df_leave.empty:
    for c in ['วันที่เริ่ม', 'วันที่สิ้นสุด']:
        if c in df_leave.columns:
            df_leave[c] = df_leave[c].apply(to_date)
else:
    df_leave = pd.DataFrame(columns=['ชื่อ-สกุล','กลุ่มงาน','ประเภทการลา','วันที่เริ่ม','วันที่สิ้นสุด','จำนวนวันลา','หมายเหตุ'])

df_travel = read_excel_from_drive(FILE_TRAVEL)
if not df_travel.empty:
    for c in ['วันที่เริ่ม', 'วันที่สิ้นสุด']:
        if c in df_travel.columns:
            df_travel[c] = df_travel[c].apply(to_date)
else:
    df_travel = pd.DataFrame(columns=['ชื่อ-สกุล','กลุ่มงาน','กิจกรรม','สถานที่','วันที่เริ่ม','วันที่สิ้นสุด','จำนวนวัน','หมายเหตุ', 'ผู้ร่วมเดินทาง'])

# =================================================================
# 🧪 Helpers & Data Processing
# =================================================================
@st.cache_data
def get_daily_status(_df_leave, _df_travel):
    def expand_date_range(df, start_col='วันที่เริ่ม', end_col='วันที่สิ้นสุด'):
        out = []
        for _, r in df.iterrows():
            s, e = r.get(start_col), r.get(end_col)
            if pd.isna(s) or pd.isna(e) or s > e:
                continue
            for d in pd.date_range(s, e, freq='D'):
                row = {'ชื่อ-สกุล': r.get('ชื่อ-สกุล'), 'วันที่': d.date()}
                if 'ประเภทการลา' in r:
                    row['สถานะ'] = f"ลา({r.get('ประเภทการลา', '')})"
                elif 'กิจกรรม' in r:
                    row['สถานะ'] = "ไปราชการ"
                out.append(row)
        return pd.DataFrame(out)

    daily_leave = expand_date_range(_df_leave)
    daily_travel = expand_date_range(_df_travel)
    daily_status = pd.concat([daily_leave, daily_travel]).drop_duplicates(subset=['ชื่อ-สกุล', 'วันที่'], keep='first')
    return daily_status

daily_status = get_daily_status(df_leave, df_travel)

def determine_status(row, status_map):
    status = status_map.get((row['ชื่อ-สกุล'], row['วันที่']))
    if status: return status
    if 'เสาร์' in str(row.get('หมายเหตุ', '')) or 'อาทิตย์' in str(row.get('หมายเหตุ', '')): return 'วันหยุด'
    is_late = str(row.get('สาย', '')).strip() not in ['', '0', '0:00', '00:00', 'None']
    if is_late: return 'สาย'
    if pd.notna(row.get('เวลาเข้า')) or pd.notna(row.get('เวลาออก')): return 'มาปกติ'
    return 'ไม่พบข้อมูล'

def build_attendance_view(month: int, year: int):
    start_date = dt.date(year, month, 1)
    end_date = (start_date + dt.timedelta(days=32)).replace(day=1) - dt.timedelta(days=1)
    att_m = df_att[(df_att['วันที่'] >= start_date) & (df_att['วันที่'] <= end_date)].copy() if not df_att.empty else df_att.copy()
    status_m = daily_status[(daily_status['วันที่'] >= start_date) & (daily_status['วันที่'] <= end_date)]
    status_map = { (r['ชื่อ-สกุล'], r['วันที่']): r['สถานะ'] for _, r in status_m.iterrows() }
    att_m['สถานะ'] = att_m.apply(determine_status, args=(status_map,), axis=1)
    att_m = att_m.sort_values(['ชื่อ-สกุล', 'วันที่'])
    summary = (att_m.groupby(['ชื่อ-สกุล', 'สถานะ'], dropna=False).size().reset_index(name='จำนวนวัน'))
    pivot = summary.pivot_table(index='ชื่อ-สกุล', columns='สถานะ', values='จำนวนวัน', aggfunc='sum', fill_value=0).reset_index()
    if 'สาย' in pivot.columns: pivot = pivot.rename(columns={'สาย': 'จำนวนครั้งมาสาย'})
    else: pivot['จำนวนครั้งมาสาย'] = 0
    return att_m, pivot

# ====================================================
# 🎯 UI Constants & Main App
# ====================================================
# Initialize session state for submission status
if 'submitted' not in st.session_state:
    st.session_state.submitted = False

def callback_submit():
    st.session_state.submitted = True

staff_groups = sorted([
    "กลุ่มโรคติดต่อ", "กลุ่มระบาดวิทยาและตอบโต้ภาวะฉุกเฉินทางสาธารณสุข", "กลุ่มพัฒนาองค์กร", "กลุ่มบริหารทั่วไป", "กลุ่มโรคไม่ติดต่อ",
    "กลุ่มห้องปฏิบัติการทางการแพทย์", "กลุ่มพัฒนานวัตกรรมและวิจัย", "กลุ่มโรคติดต่อเรื้อรัง", "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.1 จ.ชัยภูมิ",
    "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.2 จ.บุรีรัมย์", "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.3 จ.สุรินทร์",
    "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.4 ปากช่อง", "ด่านควบคุมโรคช่องจอม จ.สุรินทร์", "ศูนย์บริการเวชศาสตร์ป้องกัน",
    "กลุ่มสื่อสารความเสี่ยง", "กลุ่มโรคจากการประกอบอาชีพและสิ่งแวดล้อม"
])
leave_types = ["ลาป่วย", "ลากิจ", "ลาพักผ่อน", "อื่นๆ"]

st.title("📋 ระบบติดตามการลา ไปราชการ และการมาปฏิบัติงาน (สคร.9)")
menu = st.sidebar.radio("เลือกเมนู", ["📊 Dashboard", "📅 การมาปฏิบัติงาน", "🧭 การไปราชการ", "🕒 การลา", "🧑‍💼 ผู้ดูแลระบบ"])

# --- 📊 Dashboard ---
if menu == "📊 Dashboard":
    st.header("📊 Dashboard ภาพรวมและข้อมูลเชิงลึก")
    st.markdown("#### **ภาพรวมสะสม**")
    col1, col2, col3 = st.columns(3)
    col1.metric("เดินทางราชการ (ครั้ง)", len(df_travel))
    col2.metric("การลา (ครั้ง)", len(df_leave))
    col3.metric("ข้อมูลสแกน (แถว)", len(df_att))
    st.markdown("---")
    # ... (โค้ด Dashboard ส่วนที่เหลือเหมือนเดิม) ...

# --- 📅 Attendance View ---
elif menu == "📅 การมาปฏิบัติงาน":
    st.header("📅 สรุปการมาปฏิบัติงานรายเดือน")
    # ... (โค้ดส่วนนี้เหมือนเดิม) ...

# --- 🧭 Travel Form (เวอร์ชัน Production) ---
elif menu == "🧭 การไปราชการ":
    st.header("🧭 บันทึกการไปราชการ (สำหรับหมู่คณะ)")
    all_names = sorted(pd.concat([df_att['ชื่อ-สกุล'], df_leave['ชื่อ-สกุล'], df_travel['ชื่อ-สกุล']]).dropna().unique())
    
    with st.form("form_travel_group"):
        st.info("สำหรับหมู่คณะ: กรุณาเลือกรายชื่อทั้งหมดที่ไปราชการในครั้งนี้")
        common_data = {
            "กลุ่มงาน": st.selectbox("กลุ่มงาน", staff_groups, disabled=st.session_state.submitted),
            "กิจกรรม": st.text_input("กิจกรรม/โครงการ", disabled=st.session_state.submitted),
            "สถานที่": st.text_input("สถานที่", disabled=st.session_state.submitted),
            "วันที่เริ่ม": st.date_input("วันที่เริ่ม", dt.date.today(), disabled=st.session_state.submitted),
            "วันที่สิ้นสุด": st.date_input("วันที่สิ้นสุด", dt.date.today(), disabled=st.session_state.submitted)
        }
        selected_names = st.multiselect("1. เลือกชื่อเจ้าหน้าที่ (เลือกได้หลายคน)", options=all_names, disabled=st.session_state.submitted)
        new_names_str = st.text_area("2. เพิ่มชื่อเจ้าหน้าที่ใหม่", placeholder="ใส่ 1 ชื่อต่อ 1 บรรทัด", disabled=st.session_state.submitted)
        submitted = st.form_submit_button("💾 บันทึกข้อมูล", on_click=callback_submit, disabled=st.session_state.submitted)

    if submitted:
        new_names = [name.strip() for name in new_names_str.split('\n') if name.strip()]
        final_names = list(set(selected_names + new_names))
        if not final_names:
            st.warning("กรุณาเลือกหรือกรอก 'ชื่อ-สกุล' อย่างน้อย 1 คน")
            st.session_state.submitted = False # Reset state on validation fail
        elif common_data["วันที่เริ่ม"] > common_data["วันที่สิ้นสุด"]:
            st.error("'วันที่เริ่ม' ต้องมาก่อน 'วันที่สิ้นสุด'")
            st.session_state.submitted = False # Reset state on validation fail
        else:
            with st.spinner('⏳ กำลังบันทึกข้อมูล... กรุณารอสักครู่'):
                new_records = []
                num_days = (common_data["วันที่สิ้นสุด"] - common_data["วันที่เริ่ม"]).days + 1
                for name in final_names:
                    # สร้างรายชื่อผู้ร่วมเดินทาง
                    fellow_travelers = ", ".join([other for other in final_names if other != name])
                    record = {
                        "ชื่อ-สกุล": name,
                        "กลุ่มงาน": common_data["กลุ่มงาน"],
                        "กิจกรรม": common_data["กิจกรรม"],
                        "สถานที่": common_data["สถานที่"],
                        "วันที่เริ่ม": common_data["วันที่เริ่ม"],
                        "วันที่สิ้นสุด": common_data["วันที่สิ้นสุด"],
                        "จำนวนวัน": num_days,
                        "ผู้ร่วมเดินทาง": fellow_travelers if fellow_travelers else "-"
                    }
                    new_records.append(record)
                
                if new_records:
                    df_to_add = pd.DataFrame(new_records)
                    df_travel_new = pd.concat([df_travel, df_to_add], ignore_index=True)
                    write_excel_to_drive(FILE_TRAVEL, df_travel_new)
                    st.success(f"✅ บันทึกข้อมูลไปราชการของเจ้าหน้าที่ {len(final_names)} คนเรียบร้อยแล้ว!")
                    st.session_state.submitted = False
                    st.rerun()
    st.markdown("--- \n ### 📋 ข้อมูลปัจจุบัน")
    st.dataframe(df_travel.astype(str).sort_values('วันที่เริ่ม', ascending=False), use_container_width=True, height=420)

# --- 🕒 Leave Form ---
elif menu == "🕒 การลา":
    st.header("🕒 บันทึกการลา")
    with st.form("form_leave"):
        data = {
            "ชื่อ-สกุล": st.text_input("ชื่อ-สกุล", disabled=st.session_state.submitted),
            "กลุ่มงาน": st.selectbox("กลุ่มงาน", staff_groups, disabled=st.session_state.submitted),
            "ประเภทการลา": st.selectbox("ประเภทการลา", leave_types, disabled=st.session_state.submitted),
            "วันที่เริ่ม": st.date_input("วันที่เริ่ม", dt.date.today(), disabled=st.session_state.submitted),
            "วันที่สิ้นสุด": st.date_input("วันที่สิ้นสุด", dt.date.today(), disabled=st.session_state.submitted)
        }
        submitted = st.form_submit_button("💾 บันทึกข้อมูล", on_click=callback_submit, disabled=st.session_state.submitted)

    if submitted:
        if not data["ชื่อ-สกุล"]:
            st.warning("กรุณากรอก 'ชื่อ-สกุล'")
            st.session_state.submitted = False
        elif data["วันที่เริ่ม"] > data["วันที่สิ้นสุด"]:
            st.error("'วันที่เริ่ม' ต้องมาก่อน 'วันที่สิ้นสุด'")
            st.session_state.submitted = False
        else:
            with st.spinner('⏳ กำลังบันทึกข้อมูล...'):
                data["จำนวนวันลา"] = (data["วันที่สิ้นสุด"] - data["วันที่เริ่ม"]).days + 1
                df_leave_new = pd.concat([df_leave, pd.DataFrame([data])], ignore_index=True)
                write_excel_to_drive(FILE_LEAVE, df_leave_new)
                st.success("✅ บันทึกข้อมูลการลาเรียบร้อยแล้ว!")
                st.session_state.submitted = False
                st.rerun()

    st.markdown("--- \n ### 📋 ข้อมูลปัจจุบัน")
    st.dataframe(df_leave.astype(str).sort_values('วันที่เริ่ม', ascending=False), use_container_width=True, height=420)

# --- 🧑‍💼 Admin Panel ---
elif menu == "🧑‍💼 ผู้ดูแลระบบ":
    # ... (โค้ดส่วน Admin เหมือนเดิม) ...
