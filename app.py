# ====================================================
# 📋 โปรแกรมติดตามการลาและไปราชการ (สคร.9)
# ✨ OPTIMIZED VERSION - Fixed & Complete
# ====================================================

import io
import time
import hashlib
import logging
import datetime as dt
from typing import Dict, List, Optional, Tuple

import pandas as pd
import numpy as np
import altair as alt
import streamlit as st

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# ===========================
# 🔧 Logging Configuration
# ===========================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# ===========================
# 🔐 1. System Configuration & Auth
# ===========================
st.set_page_config(
    page_title="สคร.9 - ระบบติดตามการปฏิบัติงาน",
    page_icon="📋",
    layout="wide"
)

# Security: Check secrets
if "gcp_service_account" not in st.secrets:
    st.error("❌ Critical Error: ไม่พบข้อมูล 'gcp_service_account' ใน secrets.toml")
    logger.error("Missing gcp_service_account in secrets")
    st.stop()

# Initialize Google Drive connection with retry
@st.cache_resource
def init_drive_service():
    """Initialize Google Drive service with error handling"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            creds = service_account.Credentials.from_service_account_info(
                st.secrets["gcp_service_account"],
                scopes=["https://www.googleapis.com/auth/drive"]
            )
            service = build("drive", "v3", credentials=creds)
            logger.info("Successfully connected to Google Drive")
            return service
        except Exception as e:
            logger.warning(f"Connection attempt {attempt + 1} failed: {e}")
            if attempt == max_retries - 1:
                st.error(f"❌ ไม่สามารถเชื่อมต่อ Google Drive ได้หลัง {max_retries} ครั้ง")
                st.stop()
            time.sleep(2 ** attempt)
    return None

service = init_drive_service()

# ===========================
# 🗂️ 2. Constants & Configuration
# ===========================
FOLDER_ID = "1YFJZvs59ahRHmlRrKcQwepWJz6A-4B7d"
FILE_ATTEND = "attendance_report.xlsx"
FILE_LEAVE = "leave_report.xlsx"
FILE_TRAVEL = "travel_report.xlsx"
ATTACHMENT_FOLDER_NAME = "Attachments_Leave_App"

# Column name standardization
COLUMN_MAPPING = {
    "ชื่อพนักงาน": "ชื่อ-สกุล",
    "ชื่อ": "ชื่อ-สกุล",
    "fullname": "ชื่อ-สกุล"
}

# รายชื่อกลุ่มงาน
STAFF_GROUPS = [
    "กลุ่มบริหารทั่วไป", "กลุ่มบริหารทั่วไป (งานธุรการ)", 
    "กลุ่มบริหารทั่วไป (งานการเงินและบัญชี)",
    "กลุ่มบริหารทั่วไป (งานการเจ้าหน้าที่)", 
    "กลุ่มบริหารทั่วไป (งานพัสดุและยานพาหนะ (งานพัสดุ))",
    "กลุ่มบริหารทั่วไป (งานพัสดุและยานพาหนะ (งานยานพาหนะ))", 
    "กลุ่มบริหารทั่วไป (งานพัสดุและยานพาหนะ (งานอาคารสถานที่))",
    "กลุ่มยุทธศาสตร์และแผนงาน", 
    "กลุ่มระบาดวิทยาและตอบโต้ภาวะฉุกเฉินทางสาธารณสุข",
    "กลุ่มโรคติดต่อ", "กลุ่มโรคไม่ติดต่อ", "กลุ่มโรคติดต่อเรื้อรัง", 
    "กลุ่มโรคติดต่อนำโดยแมลง",
    "กลุ่มโรคติดต่อนำโดยแมลง (ศตม. 9.1 จ.ชัยภูมิ)", 
    "กลุ่มโรคติดต่อนำโดยแมลง (ศตม. 9.2 จ.บุรีรัมย์)",
    "กลุ่มโรคติดต่อนำโดยแมลง (ศตม. 9.3 จ.สุรินทร์)", 
    "กลุ่มโรคติดต่อนำโดยแมลง (ศตม. 9.4 อ.ปากช่อง)",
    "กลุ่มโรคจากการประกอบอาชีพและสิ่งแวดล้อม", 
    "กลุ่มห้องปฏิบัติการทางการแพทย์ด้านควบคุมโรค",
    "กลุ่มสื่อสารความเสี่ยงโรคและภัยสุขภาพ", "กลุ่มพัฒนานวัตกรรมและวิจัย", 
    "กลุ่มพัฒนาองค์กร",
    "ศูนย์ฝึกอบรมนักระบาดวิทยาภาคสนาม", "ศูนย์บริการเวชศาสตร์ป้องกัน", 
    "งานกฎหมาย", "งานเภสัชกรรม", "ด่านควบคุมโรคติดต่อระหว่างประเทศ", "อื่นๆ"
]

LEAVE_TYPES = [
    "ลาป่วย", "ลากิจส่วนตัว", "ลาพักผ่อน", 
    "ลาคลอดบุตร", "ลาอุปสมบท", "ลาช่วยเหลือภริยาที่คลอดบุตร"
]

# ===========================
# 🔧 3. Core Helper Functions
# ===========================

def get_file_id(filename: str, parent_id: str = FOLDER_ID) -> Optional[str]:
    """ค้นหา File ID จาก Google Drive with error handling"""
    try:
        q = f"name='{filename}' and '{parent_id}' in parents and trashed=false"
        res = service.files().list(
            q=q, fields="files(id, name)", 
            supportsAllDrives=True, 
            includeItemsFromAllDrives=True
        ).execute()
        files = res.get("files", [])
        if files:
            # logger.info(f"Found file: {filename}")
            return files[0]["id"]
        logger.warning(f"File not found: {filename}")
        return None
    except Exception as e:
        logger.error(f"Error finding file {filename}: {e}")
        return None

def get_or_create_folder(folder_name: str, parent_id: str) -> Optional[str]:
    """ค้นหาหรือสร้างโฟลเดอร์"""
    try:
        q = f"name='{folder_name}' and '{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        res = service.files().list(
            q=q, fields="files(id)", 
            supportsAllDrives=True, 
            includeItemsFromAllDrives=True
        ).execute()
        folder = res.get("files", [])
        
        if folder:
            return folder[0]["id"]
        
        # Create new folder
        file_metadata = {
            'name': folder_name, 
            'parents': [parent_id], 
            'mimeType': 'application/vnd.google-apps.folder'
        }
        new_folder = service.files().create(
            body=file_metadata, 
            supportsAllDrives=True, 
            fields='id'
        ).execute()
        logger.info(f"Created folder: {folder_name}")
        return new_folder.get('id')
    except Exception as e:
        logger.error(f"Error with folder {folder_name}: {e}")
        st.error(f"ไม่สามารถสร้างโฟลเดอร์ได้: {e}")
        return None

@st.cache_data(ttl=300)
def read_excel_from_drive(filename: str, max_retries: int = 3) -> pd.DataFrame:
    """อ่านไฟล์ Excel จาก Drive พร้อม retry mechanism"""
    for attempt in range(max_retries):
        try:
            file_id = get_file_id(filename)
            if not file_id:
                logger.warning(f"File not found: {filename}")
                return pd.DataFrame()
            
            req = service.files().get_media(fileId=file_id, supportsAllDrives=True)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, req)
            
            done = False
            while not done:
                _, done = downloader.next_chunk()
            
            fh.seek(0)
            df = pd.read_excel(fh, engine="openpyxl")
            # logger.info(f"Successfully read {filename}: {len(df)} rows")
            return df
            
        except Exception as e:
            logger.warning(f"Read attempt {attempt + 1} failed for {filename}: {e}")
            if attempt == max_retries - 1:
                st.error(f"ไม่สามารถอ่านไฟล์ {filename} ได้หลัง {max_retries} ครั้ง")
                return pd.DataFrame()
            time.sleep(2 ** attempt)
    
    return pd.DataFrame()

def write_excel_to_drive(filename: str, df: pd.DataFrame) -> bool:
    """บันทึกไฟล์ลง Drive with error handling"""
    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False)
        output.seek(0)

        file_id = get_file_id(filename)
        media = MediaIoBaseUpload(
            output, 
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        if file_id:
            service.files().update(
                fileId=file_id, 
                media_body=media, 
                supportsAllDrives=True
            ).execute()
        else:
            file_metadata = {"name": filename, "parents": [FOLDER_ID]}
            service.files().create(
                body=file_metadata, 
                media_body=media, 
                supportsAllDrives=True
            ).execute()
        
        st.cache_data.clear()
        return True
        
    except Exception as e:
        logger.error(f"Error saving {filename}: {e}")
        st.error(f"เกิดข้อผิดพลาดในการบันทึก: {e}")
        return False

def backup_excel(filename: str, current_df: pd.DataFrame):
    """สำรองไฟล์ก่อนแก้ไข"""
    if current_df.empty:
        return
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
        logger.warning(f"Backup failed for {filename}: {e}")

def upload_pdf_to_drive(uploaded_file, new_filename: str, folder_id: str) -> str:
    """อัปโหลดไฟล์ PDF พร้อม error handling"""
    try:
        file_metadata = {'name': new_filename, 'parents': [folder_id]}
        media = MediaIoBaseUpload(
            io.BytesIO(uploaded_file.getvalue()), 
            mimetype='application/pdf', 
            resumable=True
        )
        created_file = service.files().create(
            body=file_metadata, 
            media_body=media, 
            supportsAllDrives=True, 
            fields='id, webViewLink'
        ).execute()
        
        # Share with anyone
        file_id = created_file.get('id')
        service.permissions().create(
            fileId=file_id, 
            body={'type': 'anyone', 'role': 'reader'}, 
            supportsAllDrives=True
        ).execute()
        
        link = created_file.get('webViewLink', '-')
        return link
        
    except Exception as e:
        logger.error(f"PDF upload failed: {e}")
        st.error(f"การอัปโหลดไฟล์ล้มเหลว: {e}")
        return "-"

def count_weekdays(start_date, end_date) -> int:
    """นับวันทำการ (จ-ศ)"""
    if not start_date or not end_date:
        return 0
    if isinstance(start_date, dt.datetime):
        start_date = start_date.date()
    if isinstance(end_date, dt.datetime):
        end_date = end_date.date()
    return int(np.busday_count(start_date, end_date + dt.timedelta(days=1)))

# ===========================
# 🛡️ 4. Data Validation Functions
# ===========================

def validate_leave_data(
    name: str, 
    start_date, 
    end_date, 
    reason: str, 
    df_leave: pd.DataFrame
) -> List[str]:
    """Validate leave request data"""
    errors = []
    
    if not name or name.strip() == "":
        errors.append("❌ กรุณาเลือกชื่อ-สกุล")
    
    if start_date > end_date:
        errors.append("❌ วันที่เริ่มต้องน้อยกว่าหรือเท่ากับวันที่สิ้นสุด")
    
    if not reason or len(reason.strip()) < 5:
        errors.append("❌ กรุณาระบุเหตุผลอย่างน้อย 5 ตัวอักษร")
    
    # Check overlapping leaves
    if not df_leave.empty and name:
        start_dt = pd.to_datetime(start_date)
        end_dt = pd.to_datetime(end_date)
        
        existing_leaves = df_leave[
            (df_leave["ชื่อ-สกุล"] == name) &
            (df_leave["วันที่เริ่ม"] <= end_dt) &
            (df_leave["วันที่สิ้นสุด"] >= start_dt)
        ]
        if not existing_leaves.empty:
            errors.append("❌ มีการลาซ้ำในช่วงเวลานี้แล้ว")
    
    return errors

def validate_travel_data(
    staff_list: List[str], 
    project: str, 
    location: str,
    start_date,
    end_date
) -> List[str]:
    """Validate travel request data (Budget Removed)"""
    errors = []
    
    if not staff_list or len(staff_list) == 0:
        errors.append("❌ กรุณาเลือกผู้เดินทางอย่างน้อย 1 คน")
    
    if not project or len(project.strip()) < 3:
        errors.append("❌ กรุณาระบุชื่อโครงการ/กิจกรรม")
    
    if not location or len(location.strip()) < 3:
        errors.append("❌ กรุณาระบุสถานที่")
    
    if start_date > end_date:
        errors.append("❌ วันที่เริ่มต้องน้อยกว่าหรือเท่ากับวันที่สิ้นสุด")
        
    return errors

# ===========================
# 🔐 5. Security Functions
# ===========================

def check_admin_password(password: str) -> bool:
    """ตรวจสอบรหัสผ่าน Admin แบบ secure"""
    if not password:
        return False
    
    # Use hashed password if available
    if "admin_password_hash" in st.secrets:
        password_hash = hashlib.sha256(password.encode()).hexdigest()
        return password_hash == st.secrets["admin_password_hash"]
    
    # Fallback to plain password (not recommended)
    admin_pass = st.secrets.get("admin_password", "")
    if not admin_pass:
        st.error("⚠️ ระบบยังไม่ได้กำหนดรหัสผ่าน Admin")
        return False
    
    return password == admin_pass

# ===========================
# 📊 6. Data Processing Functions
# ===========================

def standardize_dataframe(df: pd.DataFrame, column_mapping: Dict) -> pd.DataFrame:
    """Standardize column names"""
    if df.empty:
        return df
    for old_name, new_name in column_mapping.items():
        if old_name in df.columns:
            df.rename(columns={old_name: new_name}, inplace=True)
    return df

def normalize_date_col(df: pd.DataFrame, col_name: str) -> pd.DataFrame:
    """Normalize date column"""
    if not df.empty and col_name in df.columns:
        df[col_name] = pd.to_datetime(df[col_name], errors='coerce').dt.normalize()
    return df

def clean_names(df: pd.DataFrame, col_name: str) -> pd.DataFrame:
    """Clean name strings"""
    if not df.empty and col_name in df.columns:
        df[col_name] = df[col_name].astype(str).str.strip()
    return df

def preprocess_dataframes(df_leave, df_travel, df_att):
    """Preprocess all dataframes"""
    # Standardize columns
    df_att = standardize_dataframe(df_att, COLUMN_MAPPING)
    
    # Normalize dates
    df_leave = normalize_date_col(df_leave, "วันที่เริ่ม")
    df_leave = normalize_date_col(df_leave, "วันที่สิ้นสุด")
    df_travel = normalize_date_col(df_travel, "วันที่เริ่ม")
    df_travel = normalize_date_col(df_travel, "วันที่สิ้นสุด")
    df_att = normalize_date_col(df_att, "วันที่")
    
    # Clean names
    df_leave = clean_names(df_leave, "ชื่อ-สกุล")
    df_travel = clean_names(df_travel, "ชื่อ-สกุล")
    df_att = clean_names(df_att, "ชื่อ-สกุล")
    
    return df_leave, df_travel, df_att

def get_all_names(df_leave, df_travel, df_att) -> List[str]:
    """รวมรายชื่อบุคลากรทั้งหมด"""
    all_names = set()
    if not df_leave.empty:
        all_names.update(df_leave["ชื่อ-สกุล"].unique())
    if not df_travel.empty:
        all_names.update(df_travel["ชื่อ-สกุล"].unique())
    if not df_att.empty:
        all_names.update(df_att["ชื่อ-สกุล"].unique())
    return sorted([n for n in all_names if n and str(n).lower() != 'nan'])

def create_attendance_lookup(user_att: pd.DataFrame) -> Dict:
    """สร้าง lookup dictionary สำหรับ attendance data"""
    att_lookup = {}
    if not user_att.empty:
        for _, row in user_att.iterrows():
            date_key = row["วันที่"].date() if pd.notna(row["วันที่"]) else None
            if date_key:
                att_lookup[date_key] = row.to_dict()
    return att_lookup

def parse_time(val):
    """Convert various time formats to time object"""
    if pd.isna(val):
        return None
    if isinstance(val, dt.time):
        return val
    try:
        return pd.to_datetime(str(val)).time()
    except:
        return None

# ===========================
# 🖥️ 7. UI Components
# ===========================

def show_progress(text: str, progress: int):
    """Show progress indicator"""
    if 'progress_bar' not in st.session_state:
        st.session_state.progress_bar = st.progress(0)
        st.session_state.status_text = st.empty()
    
    st.session_state.status_text.text(text)
    st.session_state.progress_bar.progress(progress)

def clear_progress():
    """Clear progress indicators"""
    if 'progress_bar' in st.session_state:
        st.session_state.progress_bar.empty()
        st.session_state.status_text.empty()
        del st.session_state.progress_bar
        del st.session_state.status_text

# ===========================
# 🚀 8. Main Application
# ===========================

# Sidebar menu
st.markdown("### 🏥 ระบบติดตามการลา ไปราชการ และการปฏิบัติงาน (สคร.9)")
menu = st.sidebar.radio(
    "📌 เมนูใช้งาน", 
    ["หน้าหลัก", "📊 Dashboard", "📅 ตรวจสอบการมาปฏิบัติงาน", 
     "🧭 บันทึกไปราชการ", "🕒 บันทึกการลา", "⚙️ ผู้ดูแลระบบ"]
)

# Lazy loading based on menu selection
if menu == "หน้าหลัก":
    # ===========================
    # 🏠 หน้าหลัก
    # ===========================
    st.info("👋 ยินดีต้อนรับเข้าสู่ระบบ HR Tracking System")
    st.markdown("""
    **ระบบนี้รองรับการทำงานดังนี้:**
    * ✅ **บันทึกการลา:** ลาป่วย, ลากิจ, ลาพักผ่อน พร้อมแนบไฟล์ PDF
    * ✅ **บันทึกไปราชการ:** บันทึกแบบรายบุคคลหรือหมู่คณะ พร้อมคำนวณวันทำการ
    * ✅ **ติดตามการมาปฏิบัติงาน:** ตรวจสอบข้อมูลสแกนนิ้ว เปรียบเทียบกับการลาและไปราชการ
    * ✅ **Dashboard:** ดูภาพรวมสถิติของหน่วยงาน
    
    ---
    **🆕 ปรับปรุงในเวอร์ชันนี้:**
    * 🚀 เพิ่มความเร็วในการโหลดข้อมูล (Lazy Loading)
    * 🛡️ ตรวจสอบความถูกต้องของข้อมูลก่อนบันทึก
    * 🔐 เพิ่มความปลอดภัยของระบบ
    * ⚡ ปรับปรุง Performance ในการประมวลผลข้อมูล
    * 📝 เพิ่ม Logging เพื่อติดตามการทำงาน
    """)
    
    st.image("https://ddc.moph.go.th/uploads/files/11120210817094038.jpg", use_container_width=True)

elif menu == "📊 Dashboard":
    # ===========================
    # 📊 Dashboard
    # ===========================
    with st.spinner("กำลังโหลดข้อมูล Dashboard..."):
        df_leave = read_excel_from_drive(FILE_LEAVE)
        df_travel = read_excel_from_drive(FILE_TRAVEL)
        df_att = read_excel_from_drive(FILE_ATTEND)
        
        df_leave, df_travel, df_att = preprocess_dataframes(df_leave, df_travel, df_att)
    
    st.header("📊 สรุปภาพรวมบุคลากร")
    
    # KPIs
    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("📋 จำนวนครั้งการลา", len(df_leave))
    with c2:
        st.metric("🚗 จำนวนครั้งไปราชการ", len(df_travel))
    with c3:
        st.metric("👆 ข้อมูลสแกน (รายการ)", len(df_att))
    
    st.divider()
    
    col_chart1, col_chart2 = st.columns(2)
    
    # Chart 1: Leave by Group
    with col_chart1:
        st.subheader("สถิติวันลาแยกตามกลุ่มงาน")
        if not df_leave.empty and "กลุ่มงาน" in df_leave.columns and "จำนวนวันลา" in df_leave.columns:
            df_chart = df_leave.groupby("กลุ่มงาน", as_index=False)["จำนวนวันลา"].sum()
            df_chart = df_chart.sort_values("จำนวนวันลา", ascending=False).head(10)
            
            chart = alt.Chart(df_chart).mark_bar().encode(
                x=alt.X("จำนวนวันลา:Q", title="รวมจำนวนวันลา"),
                y=alt.Y("กลุ่มงาน:N", sort="-x", title=""),
                color=alt.value("#6366f1"),
                tooltip=["กลุ่มงาน", "จำนวนวันลา"]
            ).properties(height=350)
            st.altair_chart(chart, use_container_width=True)
        else:
            st.info("ไม่มีข้อมูลการลา")

    # Chart 2: Top Travelers
    with col_chart2:
        st.subheader("ผู้ที่ไปราชการบ่อยที่สุด (5 อันดับ)")
        if not df_travel.empty and "ชื่อ-สกุล" in df_travel.columns:
            df_chart2 = df_travel["ชื่อ-สกุล"].value_counts().nlargest(5).reset_index()
            df_chart2.columns = ["ชื่อ-สกุล", "จำนวนครั้ง"]
            
            chart2 = alt.Chart(df_chart2).mark_bar().encode(
                x=alt.X("จำนวนครั้ง:Q", title="จำนวนครั้ง"),
                y=alt.Y("ชื่อ-สกุล:N", sort="-x", title=""),
                color=alt.value("#0ea5e9"),
                tooltip=["ชื่อ-สกุล", "จำนวนครั้ง"]
            ).properties(height=350)
            st.altair_chart(chart2, use_container_width=True)
        else:
            st.info("ไม่มีข้อมูลไปราชการ")

elif menu == "📅 ตรวจสอบการมาปฏิบัติงาน":
    # ===========================
    # 📅 ตรวจสอบการมาปฏิบัติงาน (FIXED)
    # ===========================
    with st.spinner("กำลังโหลดข้อมูล..."):
        df_att = read_excel_from_drive(FILE_ATTEND)
        df_leave = read_excel_from_drive(FILE_LEAVE)
        df_travel = read_excel_from_drive(FILE_TRAVEL)
        
        df_leave, df_travel, df_att = preprocess_dataframes(df_leave, df_travel, df_att)
        
        # สร้าง name_col ที่ถูกต้อง
        name_col = next((c for c in ["ชื่อ-สกุล", "ชื่อพนักงาน", "ชื่อ"] if c in df_att.columns), "ชื่อ-สกุล")
        if name_col not in df_att.columns and not df_att.empty:
             st.error("⚠️ ไม่พบคอลัมน์ชื่อบุคลากร")
             st.stop()
        
        # รวมรายชื่อทั้งหมด
        all_names_union = get_all_names(df_leave, df_travel, df_att)

    st.header("📅 สรุปการมาปฏิบัติงานรายวัน")

    # Filter Settings
    df_att["เดือน"] = df_att["วันที่"].dt.strftime("%Y-%m")
    months = sorted(df_att["เดือน"].dropna().unique())
    
    if not months:
        st.warning("ยังไม่มีข้อมูลเดือนในระบบ")
        months = [dt.datetime.now().strftime("%Y-%m")] # Fallback

    selected_month = st.selectbox("เลือกเดือนที่ต้องการดู", months, index=len(months)-1)
    selected_names = st.multiselect("เลือกชื่อบุคลากร (ว่าง=ทุกคน)", all_names_union)
    
    # Process Data
    df_month = df_att[df_att["เดือน"] == selected_month].copy()
    if not df_month.empty:
        df_month[name_col] = df_month[name_col].astype(str).str.strip()

    WORK_START = dt.time(8, 30)
    WORK_END = dt.time(16, 30)
    
    month_start = pd.to_datetime(selected_month + "-01")
    month_end = (month_start + pd.offsets.MonthEnd(0))
    date_range = pd.date_range(month_start, month_end, freq="D")

    records = []
    names_to_process = selected_names if selected_names else all_names_union

    prog = st.progress(0)
    for i, name in enumerate(names_to_process):
        prog.progress((i + 1) / len(names_to_process))
        
        for d in date_range:
            rec = {"ชื่อพนักงาน": name, "วันที่": d.date(), "เวลาเข้า": "", "เวลาออก": "", "หมายเหตุ": "", "สถานะ": ""}

            # 1. Scan Data
            att = df_month[(df_month[name_col] == name) & (df_month["วันที่"] == d)]
            
            # 2. Leave Data
            in_leave = False
            leave_type = ""
            user_leave = df_leave[df_leave["ชื่อ-สกุล"] == name]
            if not user_leave.empty:
                match_leave = user_leave[(user_leave["วันที่เริ่ม"] <= d) & (user_leave["วันที่สิ้นสุด"] >= d)]
                if not match_leave.empty:
                    in_leave = True
                    leave_type = match_leave.iloc[0]["ประเภทการลา"]

            # 3. Travel Data
            in_travel = False
            user_travel = df_travel[df_travel["ชื่อ-สกุล"] == name]
            if not user_travel.empty:
                match_travel = user_travel[(user_travel["วันที่เริ่ม"] <= d) & (user_travel["วันที่สิ้นสุด"] >= d)]
                if not match_travel.empty:
                    in_travel = True

            # --- Status Logic ---
            if in_leave:
                rec["สถานะ"] = f"ลา ({leave_type})"
            elif in_travel:
                rec["สถานะ"] = "ไปราชการ"
            elif not att.empty:
                row = att.iloc[0]
                rec["เวลาเข้า"] = row.get("เวลาเข้า", "")
                rec["เวลาออก"] = row.get("เวลาออก", "")
                rec["หมายเหตุ"] = row.get("หมายเหตุ", "")
                
                if d.weekday() >= 5: # Weekend
                    rec["สถานะ"] = "วันหยุด"
                else:
                    try:
                        t_in = pd.to_datetime(str(rec["เวลาเข้า"])).time() if rec["เวลาเข้า"] else None
                        t_out = pd.to_datetime(str(rec["เวลาออก"])).time() if rec["เวลาออก"] else None
                    except:
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
    st.dataframe(df_daily.style.applymap(color_status, subset=["สถานะ"]), use_container_width=True, height=500)

    st.markdown("---")
    st.subheader("📊 สรุปสถิติรวมต่อเดือนต่อคน")

    # --- FIX START: แก้ปัญหา KeyError ---
    def simplify_status(s):
        if isinstance(s, str) and s.startswith("ลา"):
            return "ลา"
        return s
    
    df_daily["สถานะย่อ"] = df_daily["สถานะ"].apply(simplify_status)
    summary = df_daily.pivot_table(index="ชื่อพนักงาน", columns="สถานะย่อ", aggfunc="size", fill_value=0)
    
    # บังคับแสดงคอลัมน์ให้ครบ
    required_cols = ["มาปกติ", "มาสาย", "ออกก่อน", "มาสายและออกก่อน", "ลา", "ไปราชการ", "วันหยุด", "ขาดงาน"]
    for col in required_cols:
        if col not in summary.columns:
            summary[col] = 0
            
    # Reorder columns
    existing_cols = [c for c in required_cols if c in summary.columns]
    other_cols = [c for c in summary.columns if c not in required_cols]
    summary = summary[existing_cols + other_cols]
    summary = summary.reset_index()
    
    st.dataframe(summary, use_container_width=True)
    # --- FIX END ---

    excel_output = io.BytesIO()
    with pd.ExcelWriter(excel_output, engine="xlsxwriter") as writer:
        df_daily.to_excel(writer, index=False, sheet_name="รายวัน")
        summary.to_excel(writer, index=False, sheet_name="สรุปสถิติรวม")
    excel_output.seek(0)
    st.download_button("📥 ดาวน์โหลดรายงานสรุป", data=excel_output, file_name=f"Summary_{selected_month}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif menu == "🧭 บันทึกไปราชการ":
    # ===========================
    # 🧭 บันทึกไปราชการ (UPDATE: เพิ่มช่องพิมพ์ชื่อเอง)
    # ===========================
    with st.spinner("กำลังโหลดข้อมูล..."):
        df_travel = read_excel_from_drive(FILE_TRAVEL)
        df_leave = read_excel_from_drive(FILE_LEAVE)
        df_att = read_excel_from_drive(FILE_ATTEND)
        
        df_leave, df_travel, df_att = preprocess_dataframes(df_leave, df_travel, df_att)
        ALL_NAMES_SORTED = get_all_names(df_leave, df_travel, df_att)
    
    st.header("📝 แบบฟอร์มขออนุมัติเดินทางไปราชการ")
    
    with st.form("form_travel"):
        c1, c2 = st.columns(2)
        with c1:
            group_job = st.selectbox("กลุ่มงาน", STAFF_GROUPS)
            project = st.text_input("ชื่อโครงการ/กิจกรรม", placeholder="ระบุชื่อโครงการ")
            location = st.text_input("สถานที่", placeholder="เช่น กรุงเทพฯ")
        with c2:
            d_start = st.date_input("วันที่เริ่มเดินทาง", value=dt.date.today())
            d_end = st.date_input("วันที่สิ้นสุดเดินทาง", value=dt.date.today())
        
        st.markdown("---")
        st.markdown("**👥 ข้อมูลผู้เดินทาง**")
        
        # 1. เลือกจากรายชื่อที่มีในระบบ
        selected_staff = st.multiselect("เลือกผู้เดินทาง (ที่มีในระบบ)", ALL_NAMES_SORTED)
        
        # 2. พิมพ์รายชื่อเพิ่มเติม (ถ้ามี)
        extra_staff_text = st.text_area(
            "ระบุรายชื่อเพิ่มเติม (กรณีไม่มีให้เลือกในช่องบน)", 
            placeholder="พิมพ์ชื่อ-สกุล แล้วคั่นด้วยเครื่องหมายจุลภาค (,) หรือขึ้นบรรทัดใหม่\nเช่น: นายสมชาย ใจดี, นางสาวสมหญิง จริงใจ"
        )

        uploaded_pdf = st.file_uploader("แนบเอกสารขออนุมัติ (PDF)", type=["pdf"])
        
        submitted = st.form_submit_button("💾 บันทึกข้อมูล", use_container_width=True)
        
        if submitted:
            # --- รวมรายชื่อจากทั้ง 2 ช่อง ---
            final_staff_list = list(selected_staff) # เริ่มต้นด้วยคนที่เลือกจาก Dropdown
            
            if extra_staff_text:
                # แปลงข้อความที่พิมพ์มา เป็น List (รองรับทั้ง , และ ขึ้นบรรทัดใหม่)
                # 1. แทนที่ newline ด้วย comma
                cleaned_text = extra_staff_text.replace("\n", ",")
                # 2. แยกด้วย comma
                manual_names = cleaned_text.split(",")
                # 3. ตัดช่องว่างหน้าหลัง และเลือกเฉพาะที่ไม่ว่างเปล่า
                manual_names = [n.strip() for n in manual_names if n.strip()]
                
                # รวมเข้าไปในลิสต์หลัก
                final_staff_list.extend(manual_names)
            
            # ตัดชื่อซ้ำออก (เผื่อเลือกซ้ำ)
            final_staff_list = sorted(list(set(final_staff_list)))

            # --- Validation ---
            errors = validate_travel_data(final_staff_list, project, location, d_start, d_end)
            
            if errors:
                for error in errors:
                    st.error(error)
            else:
                try:
                    show_progress("กำลังตรวจสอบข้อมูล...", 20)
                    time.sleep(0.5)
                    
                    # Upload file
                    link = "-"
                    if uploaded_pdf:
                        show_progress("กำลังอัปโหลดไฟล์...", 40)
                        f_id = get_or_create_folder(ATTACHMENT_FOLDER_NAME, FOLDER_ID)
                        if f_id:
                            f_name = f"TRAVEL_{dt.datetime.now().strftime('%Y%m%d_%H%M')}_{len(final_staff_list)}pax.pdf"
                            link = upload_pdf_to_drive(uploaded_pdf, f_name, f_id)
                    
                    show_progress("กำลังบันทึกข้อมูล...", 70)
                    
                    # Prepare data
                    new_rows = []
                    ts = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    days = count_weekdays(d_start, d_end)
                    
                    # วนลูปบันทึกตามรายชื่อที่รวมมาแล้ว (final_staff_list)
                    for person in final_staff_list:
                        new_rows.append({
                            "Timestamp": ts,
                            "กลุ่มงาน": group_job,
                            "ชื่อ-สกุล": person,
                            "เรื่อง/กิจกรรม": project,
                            "สถานที่": location,
                            "วันที่เริ่ม": pd.to_datetime(d_start),
                            "วันที่สิ้นสุด": pd.to_datetime(d_end),
                            "จำนวนวัน": days,
                            "ไฟล์แนบ": link
                        })
                    
                    # Save
                    show_progress("กำลังสำรองข้อมูล...", 85)
                    backup_excel(FILE_TRAVEL, df_travel)
                    
                    new_df = pd.DataFrame(new_rows)
                    df_updated = pd.concat([df_travel, new_df], ignore_index=True)
                    
                    if write_excel_to_drive(FILE_TRAVEL, df_updated):
                        show_progress("สำเร็จ!", 100)
                        time.sleep(0.5)
                        clear_progress()
                        st.success(f"✅ บันทึกข้อมูลสำเร็จ! (จำนวน {len(final_staff_list)} ท่าน)")
                        time.sleep(1)
                        st.rerun()
                    else:
                        clear_progress()
                        st.error("❌ เกิดข้อผิดพลาดในการบันทึก")
                        
                except Exception as e:
                    clear_progress()
                    logger.error(f"Travel form error: {e}")
                    st.error(f"เกิดข้อผิดพลาด: {e}")

    st.divider()
    st.subheader("📋 ประวัติการบันทึกล่าสุด")
    if not df_travel.empty:
        display_cols = ["Timestamp", "ชื่อ-สกุล", "เรื่อง/กิจกรรม", "สถานที่", "วันที่เริ่ม", "วันที่สิ้นสุด"]
        available_cols = [col for col in display_cols if col in df_travel.columns]
        st.dataframe(df_travel[available_cols].tail(5), use_container_width=True)
    else:
        st.info("ยังไม่มีข้อมูล")
        
elif menu == "🕒 บันทึกการลา":
    # ===========================
    # 🕒 บันทึกการลา
    # ===========================
    with st.spinner("กำลังโหลดข้อมูล..."):
        df_leave = read_excel_from_drive(FILE_LEAVE)
        df_travel = read_excel_from_drive(FILE_TRAVEL)
        df_att = read_excel_from_drive(FILE_ATTEND)
        
        df_leave, df_travel, df_att = preprocess_dataframes(df_leave, df_travel, df_att)
        ALL_NAMES_SORTED = get_all_names(df_leave, df_travel, df_att)
    
    st.header("📝 แบบฟอร์มบันทึกการลา")
    
    with st.form("form_leave"):
        c1, c2 = st.columns(2)
        with c1:
            l_name = st.selectbox("ชื่อ-สกุล", ALL_NAMES_SORTED)
            l_group = st.selectbox("กลุ่มงาน", STAFF_GROUPS)
            l_type = st.selectbox("ประเภทการลา", LEAVE_TYPES)
        with c2:
            l_start = st.date_input("วันที่เริ่มลา", value=dt.date.today())
            l_end = st.date_input("ถึงวันที่", value=dt.date.today())
            l_reason = st.text_area("เหตุผลการลา", placeholder="ระบุเหตุผลอย่างน้อย 5 ตัวอักษร")
            
        l_file = st.file_uploader("แนบใบลา (PDF)", type=["pdf"])
        l_submit = st.form_submit_button("💾 บันทึกการลา", use_container_width=True)
        
        if l_submit:
            # Validation
            errors = validate_leave_data(l_name, l_start, l_end, l_reason, df_leave)
            
            if errors:
                for error in errors:
                    st.error(error)
            else:
                try:
                    show_progress("กำลังตรวจสอบข้อมูล...", 20)
                    time.sleep(0.5)
                    
                    # Upload file
                    link = "-"
                    if l_file:
                        show_progress("กำลังอัปโหลดไฟล์...", 40)
                        f_id = get_or_create_folder(ATTACHMENT_FOLDER_NAME, FOLDER_ID)
                        if f_id:
                            f_name = f"LEAVE_{l_name}_{dt.datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
                            link = upload_pdf_to_drive(l_file, f_name, f_id)
                    
                    show_progress("กำลังบันทึกข้อมูล...", 70)
                    
                    days = count_weekdays(l_start, l_end)
                    new_record = {
                        "Timestamp": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "ชื่อ-สกุล": l_name,
                        "กลุ่มงาน": l_group,
                        "ประเภทการลา": l_type,
                        "วันที่เริ่ม": pd.to_datetime(l_start),
                        "วันที่สิ้นสุด": pd.to_datetime(l_end),
                        "จำนวนวันลา": days,
                        "เหตุผล": l_reason,
                        "ไฟล์แนบ": link
                    }
                    
                    show_progress("กำลังสำรองข้อมูล...", 85)
                    backup_excel(FILE_LEAVE, df_leave)
                    
                    df_upd = pd.concat([df_leave, pd.DataFrame([new_record])], ignore_index=True)
                    
                    if write_excel_to_drive(FILE_LEAVE, df_upd):
                        show_progress("สำเร็จ!", 100)
                        time.sleep(0.5)
                        clear_progress()
                        st.success("✅ บันทึกเรียบร้อย")
                        time.sleep(1)
                        st.rerun()
                    else:
                        clear_progress()
                        st.error("❌ เกิดข้อผิดพลาดในการบันทึก")
                        
                except Exception as e:
                    clear_progress()
                    logger.error(f"Leave form error: {e}")
                    st.error(f"เกิดข้อผิดพลาด: {e}")

    st.divider()
    st.subheader("📋 ประวัติการลาล่าสุด")
    if not df_leave.empty:
        display_cols = ["Timestamp", "ชื่อ-สกุล", "ประเภทการลา", "วันที่เริ่ม", "วันที่สิ้นสุด", "จำนวนวันลา"]
        available_cols = [col for col in display_cols if col in df_leave.columns]
        st.dataframe(df_leave[available_cols].tail(5), use_container_width=True)
    else:
        st.info("ยังไม่มีข้อมูล")

elif menu == "⚙️ ผู้ดูแลระบบ":
    # ===========================
    # ⚙️ ผู้ดูแลระบบ
    # ===========================
    st.header("🔒 ส่วนจัดการข้อมูล (Admin Only)")
    
    password = st.text_input("🔑 ใส่รหัสผ่าน Admin", type="password")
    
    if password and check_admin_password(password):
        st.success("✅ Access Granted")
        
        # Load all data
        with st.spinner("กำลังโหลดข้อมูล..."):
            df_leave = read_excel_from_drive(FILE_LEAVE)
            df_travel = read_excel_from_drive(FILE_TRAVEL)
            df_att = read_excel_from_drive(FILE_ATTEND)
        
        tab1, tab2, tab3, tab4 = st.tabs([
            "📂 จัดการไฟล์ลา", 
            "📂 จัดการไฟล์ราชการ", 
            "📂 จัดการไฟล์สแกนนิ้ว",
            "📊 Export รายงาน"
        ])
        
        def admin_panel(df, filename, tab_obj):
            with tab_obj:
                st.subheader(f"ไฟล์: {filename}")
                
                if df.empty:
                    st.warning("⚠️ ไม่มีข้อมูล")
                else:
                    st.dataframe(df.head(20), use_container_width=True)
                    st.caption(f"แถวทั้งหมด: {len(df)} | คอลัมน์: {len(df.columns)}")
                
                col_d1, col_d2 = st.columns(2)
                
                with col_d1:
                    # Download
                    if not df.empty:
                        buffer = io.BytesIO()
                        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                            df.to_excel(writer, index=False)
                        st.download_button(
                            f"⬇️ ดาวน์โหลด {filename}", 
                            buffer.getvalue(), 
                            filename,
                            use_container_width=True
                        )
                
                with col_d2:
                    # CSV Download
                    if not df.empty:
                        csv = df.to_csv(index=False).encode('utf-8-sig')
                        st.download_button(
                            f"⬇️ ดาวน์โหลด CSV",
                            csv,
                            f"{filename.replace('.xlsx', '.csv')}",
                            "text/csv",
                            use_container_width=True
                        )
                
                st.markdown("---")
                st.warning("⚠️ การอัปโหลดจะเขียนทับข้อมูลเดิมทั้งหมด")
                up_file = st.file_uploader(
                    f"อัปโหลดทับ {filename}", 
                    type=["xlsx"], 
                    key=f"upload_{filename}"
                )
                
                if up_file:
                    try:
                        new_df = pd.read_excel(up_file)
                        st.info(f"📄 ไฟล์ที่อัปโหลด: {len(new_df)} แถว, {len(new_df.columns)} คอลัมน์")
                        st.dataframe(new_df.head(5))
                        
                        if st.button(f"✅ ยืนยันอัปโหลด {filename}", key=f"confirm_{filename}"):
                            with st.spinner("กำลังอัปโหลด..."):
                                backup_excel(filename, df)
                                if write_excel_to_drive(filename, new_df):
                                    st.success("✅ อัปเดตไฟล์สำเร็จ!")
                                    time.sleep(1)
                                    st.rerun()
                    except Exception as e:
                        st.error(f"❌ ไม่สามารถอ่านไฟล์ได้: {e}")

        admin_panel(df_leave, FILE_LEAVE, tab1)
        admin_panel(df_travel, FILE_TRAVEL, tab2)
        admin_panel(df_att, FILE_ATTEND, tab3)
        
        # Export Tab
        with tab4:
            st.subheader("📊 Export รายงานสรุป")
            
            col_e1, col_e2 = st.columns(2)
            
            with col_e1:
                export_month = st.selectbox(
                    "เลือกเดือนที่ต้องการ Export",
                    pd.date_range(start='2024-01-01', end='2025-12-31', freq='MS').strftime("%Y-%m").tolist(),
                    index=0
                )
            
            with col_e2:
                st.write("")
                st.write("")
                if st.button("📥 สร้างรายงาน Excel", use_container_width=True):
                    with st.spinner("กำลังสร้างรายงาน..."):
                        try:
                            # Filter data by month
                            month_start = pd.to_datetime(export_month + "-01")
                            month_end = month_start + pd.offsets.MonthEnd(0)
                            
                            df_leave_month = df_leave[
                                (df_leave["วันที่เริ่ม"] >= month_start) & 
                                (df_leave["วันที่เริ่ม"] <= month_end)
                            ] if not df_leave.empty else pd.DataFrame()
                            
                            df_travel_month = df_travel[
                                (df_travel["วันที่เริ่ม"] >= month_start) & 
                                (df_travel["วันที่เริ่ม"] <= month_end)
                            ] if not df_travel.empty else pd.DataFrame()
                            
                            # Create Excel
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                if not df_leave_month.empty:
                                    df_leave_month.to_excel(writer, sheet_name='การลา', index=False)
                                if not df_travel_month.empty:
                                    df_travel_month.to_excel(writer, sheet_name='ไปราชการ', index=False)
                                
                                # Summary sheet
                                summary_data = {
                                    "รายการ": ["จำนวนการลา", "จำนวนไปราชการ"],
                                    "จำนวน": [len(df_leave_month), len(df_travel_month)]
                                }
                                pd.DataFrame(summary_data).to_excel(writer, sheet_name='สรุป', index=False)
                            
                            st.download_button(
                                "⬇️ ดาวน์โหลดรายงาน",
                                output.getvalue(),
                                f"Monthly_Report_{export_month}.xlsx",
                                use_container_width=True
                            )
                            st.success("✅ สร้างรายงานสำเร็จ!")
                            
                        except Exception as e:
                            st.error(f"❌ เกิดข้อผิดพลาด: {e}")
        
    elif password:
        st.error("❌ รหัสผ่านไม่ถูกต้อง")
        st.info("💡 หากต้องการเปลี่ยนรหัสผ่าน กรุณาติดต่อผู้พัฒนาระบบ")

