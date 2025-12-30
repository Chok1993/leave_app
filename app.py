# ====================================================
# 📋 โปรแกรมติดตามการลาและไปราชการ (สคร.9)
# ✅ Final Complete Build - Fully Debugged
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
# 🔐 1. System Configuration & Auth
# ===========================
st.set_page_config(
    page_title="สคร.9 - ระบบติดตามการปฏิบัติงาน",
    page_icon="📋",
    layout="wide"
)

# ตรวจสอบ Secrets เพื่อป้องกัน App Crash หากไม่ได้ตั้งค่า
if "gcp_service_account" not in st.secrets:
    st.error("❌ Critical Error: ไม่พบข้อมูล 'gcp_service_account' ใน secrets.toml")
    st.stop()

# เชื่อมต่อ Google Drive API
try:
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    service = build("drive", "v3", credentials=creds)
except Exception as e:
    st.error(f"❌ Connection Error: ไม่สามารถเชื่อมต่อ Google Drive ได้ ({e})")
    st.stop()

ADMIN_PASSWORD = st.secrets.get("admin_password", "admin123")

# ===========================
# 🗂️ 2. Constants & Drive Config
# ===========================
FOLDER_ID = "1YFJZvs59ahRHmlRrKcQwepWJz6A-4B7d"  # Folder หลัก (Leave_App_Data)
FILE_ATTEND = "attendance_report.xlsx"
FILE_LEAVE  = "leave_report.xlsx"
FILE_TRAVEL = "travel_report.xlsx"
ATTACHMENT_FOLDER_NAME = "Attachments_Leave_App"

# รายชื่อกลุ่มงาน
STAFF_GROUPS = [
    "กลุ่มบริหารทั่วไป", "กลุ่มบริหารทั่วไป (งานธุรการ)", "กลุ่มบริหารทั่วไป (งานการเงินและบัญชี)",
    "กลุ่มบริหารทั่วไป (งานการเจ้าหน้าที่)", "กลุ่มบริหารทั่วไป (งานพัสดุและยานพาหนะ (งานพัสดุ))",
    "กลุ่มบริหารทั่วไป (งานพัสดุและยานพาหนะ (งานยานพาหนะ))", "กลุ่มบริหารทั่วไป (งานพัสดุและยานพาหนะ (งานอาคารสถานที่))",
    "กลุ่มยุทธศาสตร์และแผนงาน", "กลุ่มระบาดวิทยาและตอบโต้ภาวะฉุกเฉินทางสาธารณสุข",
    "กลุ่มโรคติดต่อ", "กลุ่มโรคไม่ติดต่อ", "กลุ่มโรคติดต่อเรื้อรัง", "กลุ่มโรคติดต่อนำโดยแมลง",
    "กลุ่มโรคติดต่อนำโดยแมลง (ศตม. 9.1 จ.ชัยภูมิ)", "กลุ่มโรคติดต่อนำโดยแมลง (ศตม. 9.2 จ.บุรีรัมย์)",
    "กลุ่มโรคติดต่อนำโดยแมลง (ศตม. 9.3 จ.สุรินทร์)", "กลุ่มโรคติดต่อนำโดยแมลง (ศตม. 9.4 อ.ปากช่อง)",
    "กลุ่มโรคจากการประกอบอาชีพและสิ่งแวดล้อม", "กลุ่มห้องปฏิบัติการทางการแพทย์ด้านควบคุมโรค",
    "กลุ่มสื่อสารความเสี่ยงโรคและภัยสุขภาพ", "กลุ่มพัฒนานวัตกรรมและวิจัย", "กลุ่มพัฒนาองค์กร",
    "ศูนย์ฝึกอบรมนักระบาดวิทยาภาคสนาม", "ศูนย์บริการเวชศาสตร์ป้องกัน", "งานกฎหมาย",
    "งานเภสัชกรรม", "ด่านควบคุมโรคติดต่อระหว่างประเทศ", "อื่นๆ"
]

LEAVE_TYPES = ["ลาป่วย", "ลากิจส่วนตัว", "ลาพักผ่อน", "ลาคลอดบุตร", "ลาอุปสมบท", "ลาช่วยเหลือภริยาที่คลอดบุตร"]

# ===========================
# 🔧 3. Helper Functions (Core Logic)
# ===========================

def get_file_id(filename: str, parent_id=FOLDER_ID):
    """ค้นหา File ID จาก Google Drive"""
    try:
        q = f"name='{filename}' and '{parent_id}' in parents and trashed=false"
        res = service.files().list(
            q=q, fields="files(id, name)", supportsAllDrives=True, includeItemsFromAllDrives=True
        ).execute()
        files = res.get("files", [])
        return files[0]["id"] if files else None
    except Exception as e:
        st.sidebar.error(f"Error finding file {filename}: {e}")
        return None

def get_or_create_folder(folder_name: str, parent_id: str):
    """ค้นหาโฟลเดอร์เก็บไฟล์แนบ ถ้าไม่มีให้สร้างใหม่"""
    try:
        q = f"name='{folder_name}' and '{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        res = service.files().list(q=q, fields="files(id)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
        folder = res.get("files", [])
        if folder:
            return folder[0]["id"]
        else:
            file_metadata = {'name': folder_name, 'parents': [parent_id], 'mimeType': 'application/vnd.google-apps.folder'}
            new_folder = service.files().create(body=file_metadata, supportsAllDrives=True, fields='id').execute()
            return new_folder.get('id')
    except Exception as e:
        st.error(f"Error creating folder: {e}")
        return None

@st.cache_data(ttl=300)
def read_excel_from_drive(filename: str) -> pd.DataFrame:
    """อ่านไฟล์ Excel จาก Drive พร้อม Error Handling"""
    file_id = get_file_id(filename)
    if not file_id:
        return pd.DataFrame()
    
    try:
        req = service.files().get_media(fileId=file_id, supportsAllDrives=True)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)
        
        # อ่านไฟล์และจัดการ Header
        try:
            df = pd.read_excel(fh, engine="openpyxl")
        except:
            return pd.DataFrame()

        return df
    except Exception as e:
        st.error(f"Error reading {filename}: {e}")
        return pd.DataFrame()

def write_excel_to_drive(filename: str, df: pd.DataFrame):
    """บันทึกไฟล์ลง Drive (Update หรือ Create)"""
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
            file_metadata = {"name": filename, "parents": [FOLDER_ID]}
            service.files().create(body=file_metadata, media_body=media, supportsAllDrives=True).execute()
        
        st.cache_data.clear() # Clear Cache เพื่อให้ข้อมูลใหม่แสดงผลทันที
    except Exception as e:
        st.error(f"Error saving file: {e}")

def backup_excel(filename: str, current_df: pd.DataFrame):
    """สำรองไฟล์ก่อนแก้ไข"""
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
    except Exception:
        pass # Backup failed, but we continue

def upload_pdf_to_drive(uploaded_file, new_filename, folder_id):
    """อัปโหลดไฟล์ PDF"""
    try:
        file_metadata = {'name': new_filename, 'parents': [folder_id]}
        media = MediaIoBaseUpload(io.BytesIO(uploaded_file.getvalue()), mimetype='application/pdf', resumable=True)
        created_file = service.files().create(body=file_metadata, media_body=media, supportsAllDrives=True, fields='id, webViewLink').execute()
        
        # แชร์ให้ทุกคนที่มีลิงก์ดูได้
        file_id = created_file.get('id')
        service.permissions().create(fileId=file_id, body={'type': 'anyone', 'role': 'reader'}, supportsAllDrives=True).execute()
        
        return created_file.get('webViewLink')
    except Exception as e:
        st.error(f"Upload failed: {e}")
        return "-"

def count_weekdays(start_date, end_date):
    """นับวันทำการ (จ-ศ)"""
    if not start_date or not end_date: return 0
    if isinstance(start_date, dt.datetime): start_date = start_date.date()
    if isinstance(end_date, dt.datetime): end_date = end_date.date()
    return np.busday_count(start_date, end_date + dt.timedelta(days=1))

# ===========================
# 📥 4. Data Loading & Cleaning
# ===========================
df_att = read_excel_from_drive(FILE_ATTEND)
df_leave = read_excel_from_drive(FILE_LEAVE)
df_travel = read_excel_from_drive(FILE_TRAVEL)

# --- Preprocessing & Normalization (จุดสำคัญป้องกัน Bug) ---
# 1. จัดการวันที่ให้เป็น Datetime เสมอ และ Normalize เป็น 00:00:00
def normalize_date_col(df, col_name):
    if col_name in df.columns:
        df[col_name] = pd.to_datetime(df[col_name], errors='coerce').dt.normalize()
    return df

df_leave = normalize_date_col(df_leave, "วันที่เริ่ม")
df_leave = normalize_date_col(df_leave, "วันที่สิ้นสุด")
df_travel = normalize_date_col(df_travel, "วันที่เริ่ม")
df_travel = normalize_date_col(df_travel, "วันที่สิ้นสุด")
if not df_att.empty and "วันที่" in df_att.columns:
    df_att = normalize_date_col(df_att, "วันที่")

# 2. จัดการชื่อ (Trim spaces)
def clean_names(df, col_name):
    if col_name in df.columns:
        df[col_name] = df[col_name].astype(str).str.strip()
    return df

df_leave = clean_names(df_leave, "ชื่อ-สกุล")
df_travel = clean_names(df_travel, "ชื่อ-สกุล")
att_name_col = next((c for c in ["ชื่อ-สกุล", "ชื่อพนักงาน", "ชื่อ"] if c in df_att.columns), "ชื่อพนักงาน")
if not df_att.empty:
    df_att = clean_names(df_att, att_name_col)

# 3. รวมรายชื่อบุคลากรทั้งหมด
all_names = set()
if not df_leave.empty: all_names.update(df_leave["ชื่อ-สกุล"].unique())
if not df_travel.empty: all_names.update(df_travel["ชื่อ-สกุล"].unique())
if not df_att.empty: all_names.update(df_att[att_name_col].unique())
ALL_NAMES_SORTED = sorted([n for n in all_names if n and n.lower() != 'nan'])

# ===========================
# 🖥️ 5. UI & Main Logic
# ===========================
st.markdown("### 🏥 ระบบติดตามการลา ไปราชการ และการปฏิบัติงาน (สคร.9)")
menu = st.sidebar.radio("📌 เมนูใช้งาน", ["หน้าหลัก", "📊 Dashboard", "📅 ตรวจสอบการมาปฏิบัติงาน", "🧭 บันทึกไปราชการ", "🕒 บันทึกการลา", "⚙️ ผู้ดูแลระบบ"])

# ---------------------------
# 🏠 หน้าหลัก
# ---------------------------
if menu == "หน้าหลัก":
    st.info("👋 ยินดีต้อนรับเข้าสู่ระบบ HR Tracking System")
    st.markdown("""
    **ระบบนี้รองรับการทำงานดังนี้:**
    * ✅ **บันทึกการลา:** ลาป่วย, ลากิจ, ลาพักผ่อน พร้อมแนบไฟล์ PDF
    * ✅ **บันทึกไปราชการ:** บันทึกแบบรายบุคคลหรือหมู่คณะ พร้อมคำนวณวันทำการ
    * ✅ **ติดตามการมาปฏิบัติงาน:** ตรวจสอบข้อมูลสแกนนิ้ว เปรียบเทียบกับการลาและไปราชการ
    * ✅ **Dashboard:** ดูภาพรวมสถิติของหน่วยงาน
    """)
    st.image("https://ddc.moph.go.th/uploads/files/11120210817094038.jpg", use_container_width=True)

# ---------------------------
# 📊 Dashboard
# ---------------------------
elif menu == "📊 Dashboard":
    st.header("📊 สรุปภาพรวมบุคลากร")
    
    # KPIs
    c1, c2, c3 = st.columns(3)
    with c1: st.metric("📋 จำนวนครั้งการลา", len(df_leave))
    with c2: st.metric("🚗 จำนวนครั้งไปราชการ", len(df_travel))
    with c3: st.metric("fingerprint ข้อมูลสแกน (รายการ)", len(df_att))
    
    st.divider()
    
    col_chart1, col_chart2 = st.columns(2)
    
    # Chart 1: Leave by Group
    with col_chart1:
        st.subheader("สถิติวันลาแยกตามกลุ่มงาน")
        if not df_leave.empty and "กลุ่มงาน" in df_leave.columns:
            df_chart_leave = df_leave.groupby("กลุ่มงาน", as_index=False)["จำนวนวันลา"].sum().sort_values("จำนวนวันลา", ascending=False)
            chart = alt.Chart(df_chart_leave).mark_bar().encode(
                x=alt.X("จำนวนวันลา", title="รวมจำนวนวันลา"),
                y=alt.Y("กลุ่มงาน", sort="-x", title=""),
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
            df_chart_travel = df_travel["ชื่อ-สกุล"].value_counts().nlargest(5).reset_index()
            df_chart_travel.columns = ["ชื่อ-สกุล", "จำนวนครั้ง"]
            chart2 = alt.Chart(df_chart_travel).mark_bar().encode(
                x=alt.X("จำนวนครั้ง", title="จำนวนครั้ง"),
                y=alt.Y("ชื่อ-สกุล", sort="-x", title=""),
                color=alt.value("#0ea5e9"),
                tooltip=["ชื่อ-สกุล", "จำนวนครั้ง"]
            ).properties(height=350)
            st.altair_chart(chart2, use_container_width=True)
        else:
            st.info("ไม่มีข้อมูลไปราชการ")

# ---------------------------
# 📅 ตรวจสอบการมาปฏิบัติงาน (Complex Logic)
# ---------------------------
elif menu == "📅 ตรวจสอบการมาปฏิบัติงาน":
    st.header("📅 รายงานการปฏิบัติงานรายบุคคล")
    
    if df_att.empty:
        st.warning("⚠️ ยังไม่มีข้อมูลสแกนนิ้วในระบบ (ไฟล์ attendance_report.xlsx ว่างหรือไม่อยู่)")
    else:
        # Filter Settings
        df_att["เดือน_str"] = df_att["วันที่"].dt.strftime("%Y-%m")
        avail_months = sorted(df_att["เดือน_str"].dropna().unique())
        
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            selected_month = st.selectbox("เลือกเดือน", avail_months, index=len(avail_months)-1 if avail_months else 0)
        with col_f2:
            selected_person = st.selectbox("เลือกรายชื่อ", ALL_NAMES_SORTED)

        if selected_month and selected_person:
            # Generate Date Range
            curr_month_dt = pd.to_datetime(selected_month + "-01")
            days_in_month = pd.date_range(curr_month_dt, curr_month_dt + pd.offsets.MonthEnd(0), freq='D')
            
            # Prepare Dataframes for lookup
            # 1. Leave
            user_leave = df_leave[df_leave["ชื่อ-สกุล"] == selected_person].copy() if not df_leave.empty else pd.DataFrame()
            # 2. Travel
            user_travel = df_travel[df_travel["ชื่อ-สกุล"] == selected_person].copy() if not df_travel.empty else pd.DataFrame()
            # 3. Attendance
            user_att = df_att[(df_att[att_name_col] == selected_person) & (df_att["เดือน_str"] == selected_month)].copy()

            report_data = []
            
            for d in days_in_month:
                date_only = d.date()
                status = ""
                note = ""
                t_in_show = ""
                t_out_show = ""
                
                # Check 1: Leave
                is_leave = False
                if not user_leave.empty:
                    match_leave = user_leave[(user_leave["วันที่เริ่ม"] <= d) & (user_leave["วันที่สิ้นสุด"] >= d)]
                    if not match_leave.empty:
                        is_leave = True
                        l_type = match_leave.iloc[0]["ประเภทการลา"]
                        status = f"ลา ({l_type})"
                
                # Check 2: Travel
                is_travel = False
                if not is_leave and not user_travel.empty:
                    match_travel = user_travel[(user_travel["วันที่เริ่ม"] <= d) & (user_travel["วันที่สิ้นสุด"] >= d)]
                    if not match_travel.empty:
                        is_travel = True
                        status = "ไปราชการ"

                # Check 3: Weekend
                is_weekend = d.weekday() >= 5
                
                # Check 4: Attendance Scan
                scan_row = user_att[user_att["วันที่"] == d]
                
                has_scan = False
                if not scan_row.empty:
                    has_scan = True
                    row_data = scan_row.iloc[0]
                    
                    # Parse Time Logic
                    raw_in = row_data.get("เวลาเข้า")
                    raw_out = row_data.get("เวลาออก")
                    
                    # Helper to convert to time object
                    def parse_time(val):
                        if pd.isna(val): return None
                        if isinstance(val, dt.time): return val
                        try: return pd.to_datetime(str(val)).time()
                        except: return None

                    t_in = parse_time(raw_in)
                    t_out = parse_time(raw_out)
                    
                    t_in_show = t_in.strftime("%H:%M") if t_in else "-"
                    t_out_show = t_out.strftime("%H:%M") if t_out else "-"
                    note = row_data.get("หมายเหตุ", "")

                    # Status Determination
                    WORK_START = dt.time(8, 30)
                    WORK_END = dt.time(16, 30)
                    
                    if not status: # ถ้าไม่ได้ลา หรือ ไปราชการ
                        if is_weekend:
                            status = "มาทำโอที" if (t_in or t_out) else "วันหยุด"
                        else:
                            if not t_in and not t_out:
                                status = "ขาดงาน"
                            elif t_in and t_in > WORK_START:
                                status = "มาสาย"
                                if t_out and t_out < WORK_END:
                                    status += "+ออกก่อน"
                            elif t_out and t_out < WORK_END:
                                status = "ออกก่อน"
                            else:
                                status = "มาปกติ"
                
                # Final Status Cleanup
                if not status:
                    status = "วันหยุด" if is_weekend else "ขาดงาน"

                report_data.append({
                    "วันที่": date_only,
                    "สถานะ": status,
                    "เวลาเข้า": t_in_show,
                    "เวลาออก": t_out_show,
                    "หมายเหตุ": note
                })

            # Display Report
            df_report = pd.DataFrame(report_data)
            
            # Styling
            def color_row(row):
                s = row["สถานะ"]
                if "มาสาย" in s or "ออกก่อน" in s: return ["background-color: #fef08a"] * len(row) # Yellow
                if "ขาดงาน" in s: return ["background-color: #fca5a5"] * len(row) # Red
                if "ลา" in s: return ["background-color: #bfdbfe"] * len(row) # Blue
                if "ราชการ" in s: return ["background-color: #bbf7d0"] * len(row) # Green
                return [""] * len(row)

            st.dataframe(df_report.style.apply(color_row, axis=1), use_container_width=True, height=500)
            
            # Download Button
            csv = df_report.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 ดาวน์โหลดรายงาน (CSV)", csv, f"Report_{selected_person}_{selected_month}.csv", "text/csv")

# ---------------------------
# 🧭 บันทึกไปราชการ
# ---------------------------
elif menu == "🧭 บันทึกไปราชการ":
    st.header("📝 แบบฟอร์มขออนุมัติเดินทางไปราชการ")
    
    with st.form("form_travel"):
        c1, c2 = st.columns(2)
        with c1:
            group_job = st.selectbox("กลุ่มงาน", STAFF_GROUPS)
            project = st.text_input("ชื่อโครงการ/กิจกรรม")
            location = st.text_input("สถานที่")
        with c2:
            d_start = st.date_input("วันที่เริ่มเดินทาง")
            d_end = st.date_input("วันที่สิ้นสุดเดินทาง")
            budget = st.number_input("งบประมาณ (บาท)", min_value=0.0, step=100.0)
        
        staff_list = st.multiselect("เลือกผู้เดินทาง (ได้หลายคน)", ALL_NAMES_SORTED)
        uploaded_pdf = st.file_uploader("แนบเอกสารขออนุมัติ (PDF)", type=["pdf"])
        
        submitted = st.form_submit_button("💾 บันทึกข้อมูล")
        
        if submitted:
            if not staff_list or not project:
                st.error("❌ กรุณากรอกชื่อโครงการและเลือกผู้เดินทางอย่างน้อย 1 คน")
            elif d_start > d_end:
                st.error("❌ วันที่เริ่มต้องน้อยกว่าวันที่สิ้นสุด")
            else:
                with st.spinner("กำลังบันทึกข้อมูล..."):
                    # 1. Upload File
                    link = "-"
                    if uploaded_pdf:
                        f_id = get_or_create_folder(ATTACHMENT_FOLDER_NAME, FOLDER_ID)
                        f_name = f"TRAVEL_{dt.datetime.now().strftime('%Y%m%d_%H%M')}_{len(staff_list)}pax"
                        link = upload_pdf_to_drive(uploaded_pdf, f_name, f_id)
                    
                    # 2. Prepare Data
                    new_rows = []
                    ts = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    days = count_weekdays(d_start, d_end)
                    
                    for person in staff_list:
                        new_rows.append({
                            "Timestamp": ts,
                            "กลุ่มงาน": group_job,
                            "ชื่อ-สกุล": person,
                            "เรื่อง/กิจกรรม": project,
                            "สถานที่": location,
                            "วันที่เริ่ม": pd.to_datetime(d_start),
                            "วันที่สิ้นสุด": pd.to_datetime(d_end),
                            "จำนวนวัน": days,
                            "งบประมาณ": budget,
                            "ไฟล์แนบ": link
                        })
                    
                    # 3. Save
                    new_df = pd.DataFrame(new_rows)
                    backup_excel(FILE_TRAVEL, df_travel)
                    df_updated = pd.concat([df_travel, new_df], ignore_index=True)
                    write_excel_to_drive(FILE_TRAVEL, df_updated)
                    
                    st.success("✅ บันทึกข้อมูลสำเร็จ!")
                    st.rerun()

    st.subheader("📋 ประวัติการบันทึกล่าสุด")
    st.dataframe(df_travel.tail(5), use_container_width=True)

# ---------------------------
# 🕒 บันทึกการลา
# ---------------------------
elif menu == "🕒 บันทึกการลา":
    st.header("📝 แบบฟอร์มบันทึกการลา")
    
    with st.form("form_leave"):
        c1, c2 = st.columns(2)
        with c1:
            l_name = st.selectbox("ชื่อ-สกุล", ALL_NAMES_SORTED)
            l_group = st.selectbox("กลุ่มงาน", STAFF_GROUPS)
            l_type = st.selectbox("ประเภทการลา", LEAVE_TYPES)
        with c2:
            l_start = st.date_input("วันที่เริ่มลา")
            l_end = st.date_input("ถึงวันที่")
            l_reason = st.text_area("เหตุผลการลา")
            
        l_file = st.file_uploader("แนบใบลา (PDF)", type=["pdf"])
        l_submit = st.form_submit_button("💾 บันทึกการลา")
        
        if l_submit:
            if l_start > l_end:
                st.error("❌ วันที่ผิดพลาด")
            else:
                with st.spinner("กำลังบันทึก..."):
                    link = "-"
                    if l_file:
                        f_id = get_or_create_folder(ATTACHMENT_FOLDER_NAME, FOLDER_ID)
                        f_name = f"LEAVE_{l_name}_{dt.datetime.now().strftime('%Y%m%d')}"
                        link = upload_pdf_to_drive(l_file, f_name, f_id)
                    
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
                    
                    backup_excel(FILE_LEAVE, df_leave)
                    df_upd = pd.concat([df_leave, pd.DataFrame([new_record])], ignore_index=True)
                    write_excel_to_drive(FILE_LEAVE, df_upd)
                    
                    st.success("✅ บันทึกเรียบร้อย")
                    st.rerun()

    st.subheader("📋 ประวัติการลาล่าสุด")
    st.dataframe(df_leave.tail(5), use_container_width=True)

# ---------------------------
# ⚙️ ผู้ดูแลระบบ
# ---------------------------
elif menu == "⚙️ ผู้ดูแลระบบ":
    st.header("🔒 ส่วนจัดการข้อมูล (Admin Only)")
    password = st.text_input("🔑 ใส่รหัสผ่าน Admin", type="password")
    
    if password == ADMIN_PASSWORD:
        st.success("Access Granted")
        
        tab1, tab2, tab3 = st.tabs(["📂 จัดการไฟล์ลา", "📂 จัดการไฟล์ราชการ", "📂 จัดการไฟล์สแกนนิ้ว"])
        
        def admin_panel(df, filename, tab_obj):
            with tab_obj:
                st.subheader(f"ไฟล์: {filename}")
                st.dataframe(df.head(10))
                st.caption(f"แถวทั้งหมด: {len(df)}")
                
                # Download
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    df.to_excel(writer, index=False)
                st.download_button(f"⬇️ ดาวน์โหลด {filename}", buffer.getvalue(), filename)
                
                st.markdown("---")
                st.warning("⚠️ การอัปโหลดจะเขียนทับข้อมูลเดิมทั้งหมด")
                up_file = st.file_uploader(f"อัปโหลดทับ {filename}", type=["xlsx"], key=filename)
                
                if up_file:
                    if st.button(f"ยืนยันอัปโหลด {filename}"):
                        new_df = pd.read_excel(up_file)
                        backup_excel(filename, df)
                        write_excel_to_drive(filename, new_df)
                        st.success("✅ อัปเดตไฟล์สำเร็จ! กรุณารีเฟรช")

        admin_panel(df_leave, FILE_LEAVE, tab1)
        admin_panel(df_travel, FILE_TRAVEL, tab2)
        admin_panel(df_att, FILE_ATTEND, tab3)
        
    elif password:
        st.error("รหัสผ่านไม่ถูกต้อง")
