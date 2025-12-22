# ====================================================
# 📋 โปรแกรมติดตามการลาและไปราชการ (สคร.9)
# ✅ Final Stable Build - พร้อมใช้งานจริง (Leave_App_Data)
# ====================================================

import io
import mimetypes
import altair as alt
import datetime as dt
import pandas as pd
import numpy as np
import streamlit as st
import re # ## ---> ADDED THIS

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
service = build("drive", "v3", credentials=creds)
ADMIN_PASSWORD = st.secrets.get("admin_password", "admin123")

# ===========================
# 🗂️ Google Drive Config
# ===========================
FOLDER_ID = "1YFJZvs59ahRHmlRrKcQwepWJz6A-4B7d"  # Leave_App_Data
FILE_ATTEND = "attendance_report.xlsx"
FILE_LEAVE  = "leave_report.xlsx"
FILE_TRAVEL = "travel_report.xlsx"
## ---> START: ADDED THIS SECTION
ATTACHMENT_FOLDER_NAME = "Attachments_Leave_App" # ชื่อโฟลเดอร์สำหรับเก็บไฟล์แนบ
## ---> END: ADDED THIS SECTION

# ===========================
# 🔧 Helper Functions
# ===========================
def list_files_in_folder(folder_id):
    """แสดงรายชื่อไฟล์ในโฟลเดอร์ เพื่อเช็กว่าระบบมองเห็นไฟล์ไหม"""
    try:
        results = service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            fields="files(id, name)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        files = results.get("files", [])
        if not files:
            st.warning("⚠️ ไม่พบไฟล์ในโฟลเดอร์นี้เลย")
        else:
            st.info("📂 รายชื่อไฟล์ในโฟลเดอร์:")
            for f in files:
                st.write(f"✅ {f['name']}")
    except Exception as e:
        st.error(f"❌ ไม่สามารถอ่านรายชื่อไฟล์จาก Drive ได้: {e}")

def get_file_id(filename: str, parent_id=FOLDER_ID):
    """ค้นหาไฟล์ใน Google Drive ตามชื่อ"""
    q = f"name='{filename}' and '{parent_id}' in parents and trashed=false"
    res = service.files().list(
        q=q,
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()
    files = res.get("files", [])
    return files[0]["id"] if files else None

@st.cache_data(ttl=600)
def read_excel_from_drive(filename: str) -> pd.DataFrame:
    """อ่านไฟล์ Excel จาก Google Drive"""
    try:
        file_id = get_file_id(filename)
        if not file_id:
            st.warning(f"⚠️ ไม่พบไฟล์ '{filename}' ใน Drive กรุณาตรวจสอบชื่อไฟล์")
            list_files_in_folder(FOLDER_ID)
            return pd.DataFrame()

        req = service.files().get_media(fileId=file_id, supportsAllDrives=True)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)

        xls = pd.ExcelFile(fh, engine="openpyxl")
        df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
        if df.empty:
            st.warning(f"⚠️ ไฟล์ '{filename}' ไม่มีข้อมูล")
        return df

    except Exception as e:
        st.error(f"❌ อ่านไฟล์ {filename} ไม่สำเร็จ: {e}")
        return pd.DataFrame()

## ---> START: ADDED THIS SECTION
def count_weekdays(start_date, end_date):
    """นับจำนวนวันทำการ (จันทร์-ศุกร์) ระหว่าง 2 วันที่"""
    if start_date is None or end_date is None:
        return 0
    if isinstance(start_date, dt.datetime):
        start_date = start_date.date()
    if isinstance(end_date, dt.datetime):
        end_date = end_date.date()
    return np.busday_count(start_date, end_date + dt.timedelta(days=1))

def write_excel_to_drive(filename: str, df: pd.DataFrame):
    """เขียน DataFrame ลงในไฟล์ Excel บน Google Drive (เขียนทับไฟล์เดิม)"""
    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Data")
        output.seek(0)

        file_id = get_file_id(filename)
        media = MediaIoBaseUpload(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if file_id:
            service.files().update(
                fileId=file_id,
                media_body=media,
                supportsAllDrives=True
            ).execute()
        else: # ถ้าไฟล์ไม่มีอยู่ ให้สร้างใหม่
            file_metadata = {
                "name": filename,
                "parents": [FOLDER_ID]
            }
            service.files().create(
                body=file_metadata,
                media_body=media,
                supportsAllDrives=True,
                fields="id"
            ).execute()
        st.cache_data.clear() # ล้าง cache ทุกครั้งหลังเขียนไฟล์
    except Exception as e:
        st.error(f"❌ บันทึกไฟล์ {filename} ไม่สำเร็จ: {e}")

def backup_excel(filename: str, current_df: pd.DataFrame):
    """สร้างไฟล์สำรอง (backup) ก่อนที่จะเขียนทับ"""
    if current_df.empty: return
    try:
        file_id = get_file_id(filename)
        if file_id:
            timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_name = f"BAK_{timestamp}_{filename}"
            service.files().copy(
                fileId=file_id,
                body={"name": backup_name, "parents": [FOLDER_ID]},
                supportsAllDrives=True
            ).execute()
    except Exception as e:
        st.warning(f"⚠️ ไม่สามารถสร้างไฟล์สำรองได้: {e}")

def get_or_create_folder(folder_name: str, parent_id: str):
    """ค้นหาโฟลเดอร์ ถ้าไม่เจอก็สร้างใหม่"""
    q = f"name='{folder_name}' and '{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    res = service.files().list(q=q, fields="files(id)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    folder = res.get("files", [])
    if folder:
        return folder[0]["id"]
    else:
        file_metadata = {'name': folder_name, 'parents': [parent_id], 'mimeType': 'application/vnd.google-apps.folder'}
        new_folder = service.files().create(body=file_metadata, supportsAllDrives=True, fields='id').execute()
        return new_folder.get('id')

def upload_pdf_to_drive(uploaded_file, new_filename, folder_id):
    """อัปโหลดไฟล์ PDF ไปยัง Drive และคืนค่าเป็น Link"""
    try:
        file_metadata = {'name': new_filename, 'parents': [folder_id]}
        media = MediaIoBaseUpload(io.BytesIO(uploaded_file.getvalue()), mimetype='application/pdf', resumable=True)
        created_file = service.files().create(body=file_metadata, media_body=media, supportsAllDrives=True, fields='id, webViewLink').execute()

        file_id = created_file.get('id')
        permission = {'type': 'anyone', 'role': 'reader'}
        service.permissions().create(fileId=file_id, body=permission, supportsAllDrives=True).execute()

        return created_file.get('webViewLink')
    except Exception as e:
        st.error(f"❌ อัปโหลดไฟล์ไม่สำเร็จ: {e}")
        return "-"
## ---> END: ADDED THIS SECTION

# ====================================================
# 📥 Load Data
# ====================================================
df_att = read_excel_from_drive(FILE_ATTEND)
df_leave = read_excel_from_drive(FILE_LEAVE)
df_travel = read_excel_from_drive(FILE_TRAVEL)

# ป้องกัน NoneType
df_att = df_att if isinstance(df_att, pd.DataFrame) else pd.DataFrame()
df_leave = df_leave if isinstance(df_leave, pd.DataFrame) else pd.DataFrame()
df_travel = df_travel if isinstance(df_travel, pd.DataFrame) else pd.DataFrame()

# ====================================================
# 👥 รวมรายชื่อบุคลากร (แก้ไข TypeError)
# ====================================================
name_col_att = next((col for col in ["ชื่อ-สกุล", "ชื่อพนักงาน", "ชื่อ"] if col in df_att.columns), None)
all_names_leave = set(df_leave["ชื่อ-สกุล"].dropna()) if "ชื่อ-สกุล" in df_leave.columns else set()
all_names_travel = set(df_travel["ชื่อ-สกุล"].dropna()) if "ชื่อ-สกุล" in df_travel.columns else set()
all_names_att = set(df_att[name_col_att].dropna()) if name_col_att else set()

# ✅ แปลงทุกค่าก่อนเรียงลำดับ เพื่อป้องกัน TypeError ('<' not supported between str and int)
all_names = sorted(map(str, set().union(all_names_leave, all_names_travel, all_names_att)))

## ---> START: ADDED THIS SECTION
# ====================================================
# ⚙️ ค่าตั้งต้น & Lists
# ====================================================
staff_groups = [
    "กลุ่มบริหารทั่วไป",
    "กลุ่มบริหารทั่วไป (งานธุรการ)",
    "กลุ่มบริหารทั่วไป (งานการเงินและบัญชี)",
    "กลุ่มบริหารทั่วไป (งานการเจ้าหน้าที่)",
    "กลุ่มบริหารทั่วไป (งานพัสดุและยานพาหนะ (งานพัสดุ))",
    "กลุ่มบริหารทั่วไป (งานพัสดุและยานพาหนะ (งานยานพาหนะ))",
    "กลุ่มบริหารทั่วไป (งานพัสดุและยานพาหนะ (งานอาคารสถานที่))",
    "กลุ่มยุทธศาสตร์และแผนงาน",
    "กลุ่มระบาดวิทยาและตอบโต้ภาวะฉุกเฉินทางสาธารณสุข",
    "กลุ่มโรคติดต่อ",
    "กลุ่มโรคไม่ติดต่อ",
    "กลุ่มโรคติดต่อเรื้อรัง",
    "กลุ่มโรคติดต่อนำโดยแมลง",
    "กลุ่มโรคติดต่อนำโดยแมลง (ศตม. 9.1 จ.ชัยภูมิ)",
    "กลุ่มโรคติดต่อนำโดยแมลง (ศตม. 9.2 จ.บุรีรัมย์)",
    "กลุ่มโรคติดต่อนำโดยแมลง (ศตม. 9.3 จ.สุรินทร์)",
    "กลุ่มโรคติดต่อนำโดยแมลง (ศตม. 9.4 อ.ปากช่อง)",
    "กลุ่มโรคจากการประกอบอาชีพและสิ่งแวดล้อม",
    "กลุ่มห้องปฏิบัติการทางการแพทย์ด้านควบคุมโรค",
    "กลุ่มสื่อสารความเสี่ยงโรคและภัยสุขภาพ",
    "กลุ่มพัฒนานวัตกรรมและวิจัย",
    "กลุ่มพัฒนาองค์กร",
    "ศูนย์ฝึกอบรมนักระบาดวิทยาภาคสนาม",
    "ศูนย์บริการเวชศาสตร์ป้องกัน",
    "งานกฎหมาย",
    "งานเภสัชกรรม",
    "ด่านควบคุมโรคติดต่อระหว่างประเทศ",
    "ด่านควบคุมโรคติดต่อระหว่างประเทศท่าอากาศยานบุรีรัมย์ จังหวัดบุรีรัมย์",
    "ด่านควบคุมโรคติดต่อระหว่างประเทศช่องจอม อำเภอกาบเชิง จังหวัดสุรินทร์",
    "อื่นๆ"
]

leave_types = ["ลาป่วย", "ลากิจส่วนตัว", "ลาพักผ่อน", "ลาคลอดบุตร", "ลาอุปสมบท"]
## ---> END: ADDED THIS SECTION

# ====================================================
# 🧭 Interface
# ====================================================
st.markdown("##### **สำนักงานป้องกันควบคุมโรคที่ 9 จังหวัดนครราชสีมา**")
st.title("📋 ระบบติดตามการลา ไปราชการ และการปฏิบัติงาน")

# ==============================================
# 🧩 ฟังก์ชันโหลดข้อมูลใหม่ + ตรวจ header อัตโนมัติ
# ==============================================
def load_all_data(force_refresh=False):
    """โหลดข้อมูลจากไฟล์ใน Drive ทั้งหมด"""
    if force_refresh:
        st.cache_data.clear()

    def read_smart_excel(filename):
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

            # ลองอ่าน 2 แบบ
            try:
                df = pd.read_excel(fh, engine="openpyxl", header=0)
            except Exception:
                fh.seek(0)
                df = pd.read_excel(fh, engine="openpyxl", header=1)

            df = df.dropna(how="all")  # ตัดแถวว่างทั้งหมด
            return df
        except Exception as e:
            st.error(f"อ่านไฟล์ {filename} ไม่ได้: {e}")
            return pd.DataFrame()

    df_scan = read_smart_excel(FILE_ATTEND)
    df_leave = read_smart_excel(FILE_LEAVE)
    df_travel = read_smart_excel(FILE_TRAVEL)

    return df_scan, df_leave, df_travel

menu = st.sidebar.radio("เลือกเมนู", ["หน้าหลัก", "📊 Dashboard", "📅 การมาปฏิบัติงาน", "🧭 การไปราชการ", "🕒 การลา", "🧑‍💼 ผู้ดูแลระบบ"])

# ===========================
# 🏠 หน้าหลัก
# ===========================
if menu == "หน้าหลัก":
    st.info("💡 ระบบนี้ใช้สำหรับบันทึกและสรุปข้อมูลบุคลากรใน สคร.9\n"
            "ได้แก่ การลา การไปราชการ และการมาปฏิบัติงาน พร้อมแนบไฟล์เอกสาร PDF ได้โดยตรง")
    st.image("https://ddc.moph.go.th/uploads/files/11120210817094038.jpg", caption="สำนักงานป้องกันควบคุมโรคที่ 9 นครราชสีมา", use_container_width=True)

# ===========================
# 📊 Dashboard (Modern UI)
# ===========================
elif menu == "📊 Dashboard":
    st.header("📊 HR Dashboard – Overview & Insights")
    st.caption("ภาพรวมการเดินทางราชการ • การลา • การสแกนเข้าออกงาน (Travel • Leave • Attendance)")

    # ---------- Overview Cards ----------
    st.markdown("### 🔎 Overview Summary")
    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric(
            label="เดินทางราชการ (ครั้ง) | Business Trips",
            value=len(df_travel)
        )
    with col2:
        st.metric(
            label="การลา (ครั้ง) | Leave Records",
            value=len(df_leave)
        )
    with col3:
        st.metric(
            label="ข้อมูลสแกน (แถว) | Attendance Rows",
            value=len(df_att)
        )

    st.markdown("---")

    # ---------- Analytical Views ----------
    st.markdown("### 📈 Key Analytics")
    col_chart1, col_chart2 = st.columns(2)

    # การลาแยกตามกลุ่มงาน
    with col_chart1:
        st.markdown("#### Leave by Division (การลาแยกตามกลุ่มงาน)")
        if not df_leave.empty and 'กลุ่มงาน' in df_leave.columns and 'จำนวนวันลา' in df_leave.columns:
            df_leave['จำนวนวันลา'] = pd.to_numeric(df_leave['จำนวนวันลา'], errors='coerce')
            leave_by_group = (
                df_leave
                .dropna(subset=['กลุ่มงาน', 'จำนวนวันลา'])
                .groupby('กลุ่มงาน', as_index=False)['จำนวนวันลา']
                .sum()
                .sort_values('จำนวนวันลา', ascending=False)
            )
            chart_leave_group = (
                alt.Chart(leave_by_group)
                .mark_bar(color="#4F46E5")
                .encode(
                    x=alt.X('จำนวนวันลา:Q', title='Total Leave Days'),
                    y=alt.Y('กลุ่มงาน:N', sort='-x', title='กลุ่มงาน (Division)'),
                    tooltip=['กลุ่มงาน', 'จำนวนวันลา']
                )
                .properties(height=300)
            )
            st.altair_chart(chart_leave_group, use_container_width=True)
        else:
            st.info("⚠️ ไม่มีข้อมูลการลาเพียงพอสำหรับแสดงตามกลุ่มงาน")

    # ผู้เดินทางราชการบ่อยที่สุด
    with col_chart2:
        st.markdown("#### Top 5 Frequent Travelers (ผู้เดินทางราชการบ่อยที่สุด)")
        if not df_travel.empty and 'ชื่อ-สกุล' in df_travel.columns:
            top_travelers = (
                df_travel['ชื่อ-สกุล']
                .value_counts()
                .nlargest(5)
                .reset_index()
            )
            top_travelers.columns = ['ชื่อ-สกุล', 'จำนวนครั้ง']

            chart_top_travel = (
                alt.Chart(top_travelers)
                .mark_bar(color="#0EA5E9")
                .encode(
                    x=alt.X('จำนวนครั้ง:Q', title='Number of Trips'),
                    y=alt.Y('ชื่อ-สกุล:N', sort='-x', title='ชื่อ-สกุล (Name)'),
                    tooltip=['ชื่อ-สกุล', 'จำนวนครั้ง']
                )
                .properties(height=300)
            )
            st.altair_chart(chart_top_travel, use_container_width=True)
        else:
            st.info("⚠️ ไม่มีข้อมูลเดินทางราชการเพียงพอสำหรับจัดอันดับ")

    st.markdown("---")
    st.markdown("### 🧩 HR Insights")

    hr_col1, hr_col2 = st.columns(2)

    # สัดส่วนประเภทการลา
    with hr_col1:
        st.markdown("#### Leave Type Distribution (สัดส่วนประเภทการลา)")
        if not df_leave.empty and 'ประเภทการลา' in df_leave.columns and 'จำนวนวันลา' in df_leave.columns:
            df_leave['จำนวนวันลา'] = pd.to_numeric(df_leave['จำนวนวันลา'], errors='coerce')
            df_leave_cleaned = df_leave.dropna(subset=['จำนวนวันลา'])

            leave_type_dist = (
                df_leave_cleaned
                .groupby('ประเภทการลา', as_index=False)['จำนวนวันลา']
                .sum()
            )
            chart_leave_type = (
                alt.Chart(leave_type_dist)
                .mark_arc(innerRadius=60)
                .encode(
                    theta=alt.Theta(field="จำนวนวันลา", type="quantitative", title="Total Leave Days"),
                    color=alt.Color(field="ประเภทการลา", type="nominal", title="Leave Type"),
                    tooltip=['ประเภทการลา', 'จำนวนวันลา']
                )
                .properties(height=300)
            )
            st.altair_chart(chart_leave_type, use_container_width=True)
        else:
            st.info("⚠️ ไม่มีข้อมูลประเภทการลาเพียงพอสำหรับแสดงสัดส่วน")

    # แนวโน้มการลารายเดือน
    with hr_col2:
        st.markdown("#### Monthly Leave Trend (แนวโน้มการลารายเดือน)")
        if not df_leave.empty and 'วันที่เริ่ม' in df_leave.columns and 'จำนวนวันลา' in df_leave.columns:
            df_leave_copy = df_leave.copy()
            df_leave_copy['วันที่เริ่ม'] = pd.to_datetime(df_leave_copy['วันที่เริ่ม'], errors='coerce')
            df_leave_copy['จำนวนวันลา'] = pd.to_numeric(df_leave_copy['จำนวนวันลา'], errors='coerce')
            df_leave_copy.dropna(subset=['วันที่เริ่ม', 'จำนวนวันลา'], inplace=True)

            df_leave_copy['เดือน'] = df_leave_copy['วันที่เริ่ม'].dt.to_period('M').dt.to_timestamp()
            monthly_leave = (
                df_leave_copy
                .groupby('เดือน', as_index=False)['จำนวนวันลา']
                .sum()
            )

            chart_monthly_trend = (
                alt.Chart(monthly_leave)
                .mark_line(point=True, strokeWidth=3, color="#22C55E")
                .encode(
                    x=alt.X('เดือน:T', title='Month'),
                    y=alt.Y('จำนวนวันลา:Q', title='Total Leave Days'),
                    tooltip=['เดือน', 'จำนวนวันลา']
                )
                .properties(height=300)
            )
            st.altair_chart(chart_monthly_trend, use_container_width=True)
        else:
            st.info("⚠️ ไม่มีข้อมูลวันที่เริ่มการลาเพียงพอสำหรับแสดงแนวโน้มรายเดือน")

    # ---------- Attendance Risk ----------
    st.markdown("### 🚨 Attendance Risk – Top 5 (สาย/ขาด/ออกก่อนเวลา)")
    st.caption("Top 5 พนักงานที่มาสายหรือขาดงานในเดือนล่าสุด (ยกเว้นวันที่ลา/ไปราชการ)")

    if not df_att.empty:
        df_att_copy = df_att.copy()
        df_att_copy["วันที่"] = pd.to_datetime(df_att_copy["วันที่"], errors="coerce")

        if not df_leave.empty:
            df_leave['วันที่เริ่ม'] = pd.to_datetime(df_leave['วันที่เริ่ม'], errors='coerce')
            df_leave['วันที่สิ้นสุด'] = pd.to_datetime(df_leave['วันที่สิ้นสุด'], errors='coerce')
        if not df_travel.empty:
            df_travel['วันที่เริ่ม'] = pd.to_datetime(df_travel['วันที่เริ่ม'], errors='coerce')
            df_travel['วันที่สิ้นสุด'] = pd.to_datetime(df_travel['วันที่สิ้นสุด'], errors='coerce')

        latest_month_str = df_att_copy["วันที่"].dt.strftime("%Y-%m").max()
        df_month = df_att_copy[df_att_copy["วันที่"].dt.strftime("%Y-%m") == latest_month_str]

        name_col = next((c for c in ["ชื่อ-สกุล", "ชื่อพนักงาน", "ชื่อ"] if c in df_month.columns), None)

        if name_col and latest_month_str:
            records = []
            WORK_START = dt.time(8, 30)

            all_names_month = df_month[name_col].dropna().unique().tolist()
            month_start = pd.to_datetime(latest_month_str)
            all_days_in_month = pd.date_range(
                start=month_start,
                end=month_start + pd.offsets.MonthEnd(0),
                freq='D'
            )

            for name in all_names_month:
                for day in all_days_in_month:
                    if day.weekday() >= 5:
                        continue

                    is_on_leave = (
                        not df_leave.empty and
                        not df_leave[
                            (df_leave["ชื่อ-สกุล"] == name) &
                            (df_leave["วันที่เริ่ม"] <= day) &
                            (df_leave["วันที่สิ้นสุด"] >= day)
                        ].empty
                    )
                    is_on_travel = (
                        not df_travel.empty and
                        not df_travel[
                            (df_travel["ชื่อ-สกุล"] == name) &
                            (df_travel["วันที่เริ่ม"] <= day) &
                            (df_travel["วันที่สิ้นสุด"] >= day)
                        ].empty
                    )

                    if is_on_leave or is_on_travel:
                        continue

                    att_record = df_month[
                        (df_month[name_col] == name) &
                        (df_month['วันที่'].dt.date == day.date())
                    ]

                    status = ""
                    if att_record.empty:
                        status = "ขาดงาน"
                    else:
                        time_val = att_record.iloc[0].get("เวลาเข้า")
                        t_in = None

                        if time_val:
                            if isinstance(time_val, dt.time):
                                t_in = time_val
                            else:
                                parsed_dt = pd.to_datetime(str(time_val), errors='coerce')
                                if pd.notna(parsed_dt):
                                    t_in = parsed_dt.time()

                        if not t_in:
                            status = "ขาดงาน"
                        elif t_in > WORK_START:
                            status = "มาสาย"

                    if status:
                        records.append({"ชื่อพนักงาน": name, "สถานะ": status})

            if records:
                df_issues = pd.DataFrame(records)
                top_issues = (
                    df_issues['ชื่อพนักงาน']
                    .value_counts()
                    .nlargest(5)
                    .reset_index()
                )
                top_issues.columns = ['ชื่อพนักงาน', 'จำนวนครั้ง']

                chart_top_issues = (
                    alt.Chart(top_issues)
                    .mark_bar(color="#EF4444")
                    .encode(
                        x=alt.X('จำนวนครั้ง:Q', title='Counts (Late/Absent)'),
                        y=alt.Y('ชื่อพนักงาน:N', sort='-x', title='Employee'),
                        tooltip=['ชื่อพนักงาน', 'จำนวนครั้ง']
                    )
                    .properties(height=300)
                )
                st.altair_chart(chart_top_issues, use_container_width=True)
            else:
                st.info("✅ ไม่พบข้อมูลการมาสายหรือขาดงานในเดือนล่าสุด")
        else:
            st.info("⚠️ ไม่พบคอลัมน์ชื่อพนักงาน หรือเดือนล่าสุดไม่ถูกต้อง")
    else:
        st.info("⚠️ ไม่มีข้อมูลการสแกน (Attendance Data) ให้วิเคราะห์")

# ===========================
# 📅 การมาปฏิบัติงาน
# ===========================
elif menu == "📅 การมาปฏิบัติงาน":
    st.header("📅 สรุปการมาปฏิบัติงานรายวัน (ตรวจจากสแกน + ลา + ราชการ)")

    if df_att.empty:
        st.warning("ยังไม่มีข้อมูลสแกนเข้า-ออกในระบบ")
        st.stop()

    # ✅ ตรวจสอบชื่อคอลัมน์บุคลากร
    name_col = next((c for c in ["ชื่อ-สกุล", "ชื่อพนักงาน", "ชื่อ"] if c in df_att.columns), None)
    if not name_col:
        st.error("⚠️ ไม่พบคอลัมน์ชื่อบุคลากร (เช่น 'ชื่อพนักงาน' หรือ 'ชื่อ-สกุล')")
        st.stop()

    # ---------------------------------------------------------
    # 🛠️ ส่วนที่แก้ไข: แปลงวันที่ทั้งหมดให้เป็น 00:00:00 (Normalize)
    # ---------------------------------------------------------
    # 1. จัดการ df_att
    df_att["วันที่"] = pd.to_datetime(df_att["วันที่"], errors="coerce").dt.normalize()
    
    # 2. จัดการ df_leave และ df_travel
    for df in [df_leave, df_travel]:
        # ตัดช่องว่างชื่อ (ถ้ามี)
        if "ชื่อ-สกุล" in df.columns:
             df["ชื่อ-สกุล"] = df["ชื่อ-สกุล"].astype(str).str.strip()
             
        for c in ["วันที่เริ่ม", "วันที่สิ้นสุด"]:
            if c in df.columns:
                # 🔥 ไฮไลท์: ใช้ .dt.normalize() เพื่อตัดเวลาทิ้ง เหลือแค่ YYYY-MM-DD 00:00:00
                df[c] = pd.to_datetime(df[c], errors="coerce").dt.normalize()
    # ---------------------------------------------------------

    # ✅ สร้างคอลัมน์เดือน
    df_att["เดือน"] = df_att["วันที่"].dt.strftime("%Y-%m")
    months = sorted(df_att["เดือน"].dropna().unique())
    selected_month = st.selectbox("เลือกเดือนที่ต้องการดู", months, index=len(months)-1)

    # รวมรายชื่อทั้งหมด (ทั้งจากสแกน, ลา, ราชการ) เพื่อให้ครบถ้วน
    all_names_union = sorted(list(set(df_att[name_col].astype(str).str.strip().unique()) | set(df_leave.get("ชื่อ-สกุล", [])) | set(df_travel.get("ชื่อ-สกุล", []))))
    selected_names = st.multiselect("เลือกชื่อบุคลากร (ว่าง=ทุกคน)", all_names_union)
    
    df_month = df_att[df_att["เดือน"] == selected_month].copy()
    # ตัดช่องว่างชื่อใน df_month ด้วยเพื่อความชัวร์
    df_month[name_col] = df_month[name_col].astype(str).str.strip()

    WORK_START = dt.time(8, 30)
    WORK_END = dt.time(16, 30)
    
    # สร้างช่วงวันที่ในเดือนที่เลือก (d ใน loop นี้จะเป็น timestamp 00:00:00)
    month_start = pd.to_datetime(selected_month + "-01")
    month_end = (month_start + pd.offsets.MonthEnd(0))
    date_range = pd.date_range(month_start, month_end, freq="D")

    records = []
    names_to_process = selected_names if selected_names else all_names_union

    # Progress bar (เผื่อข้อมูลเยอะ)
    prog = st.progress(0)

    for i, name in enumerate(names_to_process):
        prog.progress((i + 1) / len(names_to_process))
        
        for d in date_range:
            rec = {"ชื่อพนักงาน": name, "วันที่": d.date(), "เวลาเข้า": "", "เวลาออก": "", "หมายเหตุ": "", "สถานะ": ""}

            # 1. ดึงข้อมูลสแกน (เทียบวันที่แบบ 00:00:00 == 00:00:00)
            att = df_month[(df_month[name_col] == name) & (df_month["วันที่"] == d)]
            
            # 2. ตรวจสอบการลา (เทียบช่วงวันที่แบบไม่มีเวลา)
            in_leave = False
            leave_type = ""
            # กรองข้อมูลลาของคนนี้ และวันที่ d อยู่ในช่วงวันลา
            user_leave = df_leave[df_leave["ชื่อ-สกุล"] == name]
            if not user_leave.empty:
                # เนื่องจาก normalize มาแล้ว เปรียบเทียบได้เลย
                match_leave = user_leave[(user_leave["วันที่เริ่ม"] <= d) & (user_leave["วันที่สิ้นสุด"] >= d)]
                if not match_leave.empty:
                    in_leave = True
                    leave_type = match_leave.iloc[0]["ประเภทการลา"]

            # 3. ตรวจสอบไปราชการ
            in_travel = False
            user_travel = df_travel[df_travel["ชื่อ-สกุล"] == name]
            if not user_travel.empty:
                match_travel = user_travel[(user_travel["วันที่เริ่ม"] <= d) & (user_travel["วันที่สิ้นสุด"] >= d)]
                if not match_travel.empty:
                    in_travel = True

            # --- กำหนดสถานะ ---
            if in_leave:
                rec["สถานะ"] = f"ลา ({leave_type})"
            elif in_travel:
                rec["สถานะ"] = "ไปราชการ"
            elif not att.empty:
                row = att.iloc[0]
                rec["เวลาเข้า"] = row.get("เวลาเข้า", "")
                rec["เวลาออก"] = row.get("เวลาออก", "")
                rec["หมายเหตุ"] = row.get("หมายเหตุ", "")
                
                if d.weekday() >= 5: # เสาร์-อาทิตย์
                    rec["สถานะ"] = "วันหยุด"
                else:
                    try:
                        t_in = pd.to_datetime(str(rec["เวลาเข้า"])).time() if rec["เวลาเข้า"] else None
                        t_out = pd.to_datetime(str(rec["เวลาออก"])).time() if rec["เวลาออก"] else None
                    except Exception:
                        t_in, t_out = None, None
                    
                    if not t_in and not t_out:
                        rec["สถานะ"] = "ขาดงาน"
                    elif t_in and t_in > WORK_START:
                        if not t_out or t_out < WORK_END:
                            rec["สถานะ"] = "มาสายและออกก่อน"
                        else:
                            rec["สถานะ"] = "มาสาย"
                    elif not t_out or t_out < WORK_END:
                        rec["สถานะ"] = "ออกก่อน"
                    else:
                        rec["สถานะ"] = "มาปกติ"
            else:
                rec["สถานะ"] = "วันหยุด" if d.weekday() >= 5 else "ขาดงาน"
            
            records.append(rec)
            
    prog.empty()

    df_daily = pd.DataFrame(records)
    if not df_daily.empty:
        df_daily = df_daily.sort_values(["ชื่อพนักงาน", "วันที่"])

    def color_status(val):
        colors = {
            "มาปกติ": "background-color:#d4edda",
            "มาสาย": "background-color:#ffeeba",
            "ออกก่อน": "background-color:#f8d7da",
            "มาสายและออกก่อน": "background-color:#fcd5b5",
            "ลา": "background-color:#d1ecf1",
            "ไปราชการ": "background-color:#fff3cd",
            "วันหยุด": "background-color:#e2e3e5",
            "ขาดงาน": "background-color:#f5c6cb"
        }
        for key in colors:
            if key in str(val):
                return colors[key]
        return ""

    st.markdown("### 📋 ตารางสรุปสถานะรายวัน")
    st.dataframe(df_daily.style.applymap(color_status, subset=["สถานะ"]), use_container_width=True, height=600)

    st.markdown("---")
    st.subheader("📊 สรุปสถิติรวมต่อเดือนต่อคน")

    def simplify_status(s):
        return "ลา" if isinstance(s, str) and s.startswith("ลา") else s
    df_daily["สถานะย่อ"] = df_daily["สถานะ"].apply(simplify_status)

    summary = df_daily.pivot_table(index="ชื่อพนักงาน", columns="สถานะย่อ", aggfunc="size", fill_value=0).reset_index()
    st.dataframe(summary, use_container_width=True)

    excel_output = io.BytesIO()
    with pd.ExcelWriter(excel_output, engine="xlsxwriter") as writer:
        df_daily.to_excel(writer, index=False, sheet_name="รายวัน")
        summary.to_excel(writer, index=False, sheet_name="สรุปสถิติรวม")
    excel_output.seek(0)
    st.download_button("📥 ดาวน์โหลดรายงานสรุป (รายวัน + รวมต่อเดือน)", data=excel_output, file_name=f"รายงานสรุป_{selected_month}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ----------------------------
# 🧭 การไปราชการ
# ----------------------------
elif menu == "🧭 การไปราชการ":
    st.header("🧭 บันทึกข้อมูลไปราชการ (หมู่คณะ)")
    with st.form("form_travel_group"):
        data_ui = {
            "กลุ่มงาน": st.selectbox("กลุ่มงาน", staff_groups),
            "กิจกรรม": st.text_input("กิจกรรม/โครงการ"),
            "สถานที่": st.text_input("สถานที่"),
            "วันที่เริ่ม": st.date_input("วันที่เริ่ม", dt.date.today()),
            "วันที่สิ้นสุด": st.date_input("วันที่สิ้นสุด", dt.date.today())
        }
        days = count_weekdays(data_ui["วันที่เริ่ม"], data_ui["วันที่สิ้นสุด"])
        if days > 0:
            st.caption(f"🗓️ รวมวันทำการ {days} วัน")

        selected = st.multiselect("รายชื่อผู้เดินทาง", options=all_names)
        new_names = st.text_area("เพิ่มชื่อใหม่ (ถ้ามี)", placeholder="ใส่ชื่อทีละบรรทัด")
        upload = st.file_uploader("แนบไฟล์คำสั่ง (PDF)", type="pdf")

        submitted = st.form_submit_button("💾 บันทึกข้อมูล")

    if submitted:
        names = list(set(selected + [x.strip() for x in new_names.splitlines() if x.strip()]))
        if not names:
            st.warning("กรุณาเลือกหรือกรอกชื่ออย่างน้อย 1 คน")
        elif data_ui["วันที่เริ่ม"] > data_ui["วันที่สิ้นสุด"]:
            st.error("วันที่เริ่มต้องไม่เกินวันที่สิ้นสุด")
        else:
            folder_id = get_or_create_folder(ATTACHMENT_FOLDER_NAME, FOLDER_ID)
            file_link = "-"
            if upload:
                safe_name = re.sub(r"[^\wก-๙]", "_", data_ui["กิจกรรม"])
                filename = f"{data_ui['วันที่เริ่ม']}_{safe_name}_{names[0]}.pdf"
                file_link = upload_pdf_to_drive(upload, filename, folder_id)

            backup_excel(FILE_TRAVEL, df_travel)
            now = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
            records = []
            for n in names:
                records.append({
                    **data_ui, "ชื่อ-สกุล": n,
                    "ผู้ร่วมเดินทาง": ", ".join([x for x in names if x != n]) or "-",
                    "จำนวนวัน": days,
                    "ลิงก์เอกสาร": file_link,
                    "last_update": now
                })
            df_new = pd.concat([df_travel, pd.DataFrame(records)], ignore_index=True)
            write_excel_to_drive(FILE_TRAVEL, df_new)
            st.success("✅ บันทึกข้อมูลเรียบร้อยแล้ว!")
            st.rerun()

    st.markdown("### 📋 ข้อมูลไปราชการทั้งหมด")
    if not df_travel.empty:
        st.dataframe(df_travel.astype(str).sort_values("วันที่เริ่ม", ascending=False),
                     column_config={"ลิงก์เอกสาร": st.column_config.LinkColumn("เอกสารแนบ", display_text="🔗 เปิดไฟล์")})
    else:
        st.info("ยังไม่มีข้อมูลไปราชการ")

# ----------------------------
# 🕒 การลา
# ----------------------------
elif menu == "🕒 การลา":
    st.header("🕒 บันทึกข้อมูลการลา")
    with st.form("form_leave"):
        name = st.text_input("ชื่อ-สกุล")
        group = st.selectbox("กลุ่มงาน", staff_groups)
        leave_type = st.selectbox("ประเภทการลา", leave_types)
        start = st.date_input("วันที่เริ่ม", dt.date.today())
        end = st.date_input("วันที่สิ้นสุด", dt.date.today())
        days = count_weekdays(start, end)
        if days > 0:
            st.caption(f"🗓️ รวมวันทำการ {days} วัน")
        submit_leave = st.form_submit_button("💾 บันทึกข้อมูล")

    if submit_leave:
        if not name:
            st.warning("กรุณากรอกชื่อ")
        elif start > end:
            st.error("วันที่เริ่มต้องไม่เกินวันที่สิ้นสุด")
        else:
            backup_excel(FILE_LEAVE, df_leave)
            rec = {
                "ชื่อ-สกุล": name, "กลุ่มงาน": group, "ประเภทการลา": leave_type,
                "วันที่เริ่ม": start, "วันที่สิ้นสุด": end,
                "จำนวนวันลา": days,
                "last_update": dt.datetime.now().strftime("%Y-%m-%d %H:%M")
            }
            df_new = pd.concat([df_leave, pd.DataFrame([rec])], ignore_index=True)
            write_excel_to_drive(FILE_LEAVE, df_new)
            st.success("✅ บันทึกข้อมูลการลาเรียบร้อย")
            st.rerun()

    st.markdown("### 📋 ข้อมูลการลาทั้งหมด")
    if not df_leave.empty:
        st.dataframe(df_leave.astype(str).sort_values("วันที่เริ่ม", ascending=False))
    else:
        st.info("ยังไม่มีข้อมูลการลา")

elif menu == "🧑‍💼 ผู้ดูแลระบบ":
    st.header("🔐 ผู้ดูแลระบบ")
    
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

    st.success("คุณได้เข้าสู่ระบบผู้ดูแลแล้ว 🧑‍💼")
    if st.button("🚪 ออกจากระบบ"):
        st.session_state.admin_logged_in = False
        st.rerun()

    tabA, tabB, tabC = st.tabs(["📗 แก้ไขข้อมูลการลา", "📘 แก้ไขข้อมูลไปราชการ", "🟩 แก้ไขข้อมูลสแกน"])

    with tabA:
        st.caption("แก้ไขตารางด้านล่างได้โดยตรง (เพิ่ม/ลบ/แก้ไข) แล้วกดปุ่มบันทึก")
        edited_leave = st.data_editor(df_leave.astype(str), num_rows="dynamic", use_container_width=True, key="ed_leave")
        if st.button("💾 บันทึกข้อมูลการลา", key="save_leave"):
            with st.spinner("กำลังบันทึก..."):
                backup_excel(FILE_LEAVE, df_leave)
                edited_leave['last_update'] = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
                write_excel_to_drive(FILE_LEAVE, pd.DataFrame(edited_leave))
                st.success("✅ บันทึกข้อมูลการลาเรียบร้อย")
                st.rerun()
        
        out_leave = io.BytesIO()
        with pd.ExcelWriter(out_leave, engine="xlsxwriter") as writer: pd.DataFrame(edited_leave).to_excel(writer, index=False)
        out_leave.seek(0)
        st.download_button("⬇️ ดาวน์โหลดข้อมูลทั้งหมด (Excel)", data=out_leave, file_name="leave_all_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_leave")

    with tabB:
        st.caption("แก้ไขตารางด้านล่างได้โดยตรง (เพิ่ม/ลบ/แก้ไข) แล้วกดปุ่มบันทึก")
        edited_travel = st.data_editor(df_travel.astype(str), num_rows="dynamic", use_container_width=True, key="ed_travel")
        if st.button("💾 บันทึกข้อมูลไปราชการ", key="save_travel"):
            with st.spinner("กำลังบันทึก..."):
                backup_excel(FILE_TRAVEL, df_travel)
                edited_travel['last_update'] = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
                write_excel_to_drive(FILE_TRAVEL, pd.DataFrame(edited_travel))
                st.success("✅ บันทึกข้อมูลไปราชการเรียบร้อย")
                st.rerun()
        out_travel = io.BytesIO()
        with pd.ExcelWriter(out_travel, engine="xlsxwriter") as writer: pd.DataFrame(edited_travel).to_excel(writer, index=False)
        out_travel.seek(0)
        st.download_button("⬇️ ดาวน์โหลดข้อมูลทั้งหมด (Excel)", data=out_travel, file_name="travel_all_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_travel")

    with tabC:
        st.caption("ข้อมูลสแกนมีขนาดใหญ่ แนะนำให้แก้ไขเท่าที่จำเป็น (เช่น เติมหมายเหตุ)")
        edited_att = st.data_editor(df_att.astype(str), num_rows="dynamic", use_container_width=True, key="ed_att")
        if st.button("💾 บันทึกข้อมูลสแกน", key="save_att"):
            with st.spinner("กำลังบันทึก..."):
                backup_excel(FILE_ATTEND, df_att)
                write_excel_to_drive(FILE_ATTEND, pd.DataFrame(edited_att))
                st.success("✅ บันทึกข้อมูลสแกนเรียบร้อย")
                st.rerun()
        out_att = io.BytesIO()
        with pd.ExcelWriter(out_att, engine="xlsxwriter") as writer: pd.DataFrame(edited_att).to_excel(writer, index=False)
        out_att.seek(0)
        st.download_button("⬇️ ดาวน์โหลดข้อมูลทั้งหมด (Excel)", data=out_att, file_name="attendance_all_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_att")









