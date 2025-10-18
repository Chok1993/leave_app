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
    "กลุ่มอำนวยการ", "กลุ่มยุทธศาสตร์และแผนงาน", "กลุ่มระบาดวิทยา",
    "กลุ่มโรคติดต่อ", "กลุ่มโรคไม่ติดต่อ", "กลุ่มวัณโรค", "กลุ่มโรคเอดส์", "กลุ่มโรคจากการประกอบอาชีพและสิ่งแวดล้อม",
    "กลุ่มพัฒนาองค์กร", "อื่นๆ"
]
leave_types = ["ลาป่วย", "ลากิจส่วนตัว", "ลาพักผ่อน", "ลาคลอดบุตร", "ลาอุปสมบท"]
## ---> END: ADDED THIS SECTION

# ====================================================
# 🧭 Interface
# ====================================================
st.markdown("##### **สำนักงานป้องกันควบคุมโรคที่ 9 จังหวัดนครราชสีมา**")
st.title("📋 ระบบติดตามการลา ไปราชการ และการปฏิบัติงาน")

menu = st.sidebar.radio("เลือกเมนู", ["หน้าหลัก", "📊 Dashboard", "📅 การมาปฏิบัติงาน", "🧭 การไปราชการ", "🕒 การลา", "🧑‍💼 ผู้ดูแลระบบ"])

# ===========================
# 🏠 หน้าหลัก
# ===========================
if menu == "หน้าหลัก":
    st.info("💡 ระบบนี้ใช้สำหรับบันทึกและสรุปข้อมูลบุคลากรใน สคร.9\n"
            "ได้แก่ การลา การไปราชการ และการมาปฏิบัติงาน พร้อมแนบไฟล์เอกสาร PDF ได้โดยตรง")
    st.image("https://ddc.moph.go.th/uploads/files/11120210817094038.jpg", caption="สำนักงานป้องกันควบคุมโรคที่ 9 นครราชสีมา", use_container_width=True)

# ===========================
# 📊 Dashboard
# ===========================
elif menu == "📊 Dashboard":
    st.header("📊 Dashboard ภาพรวมและข้อมูลเชิงลึก")
    st.markdown("#### **ภาพรวมสะสม**")
    col1, col2, col3 = st.columns(3)
    col1.metric("เดินทางราชการ (ครั้ง)", len(df_travel))
    col2.metric("การลา (ครั้ง)", len(df_leave))
    col3.metric("ข้อมูลสแกน (แถว)", len(df_att))
    st.markdown("---")

    col_chart1, col_chart2 = st.columns(2)
    with col_chart1:
        st.markdown("##### **การลาแยกตามกลุ่มงาน**")
        if not df_leave.empty and 'กลุ่มงาน' in df_leave.columns and 'จำนวนวันลา' in df_leave.columns:
            leave_by_group = df_leave.groupby('กลุ่มงาน')['จำนวนวันลา'].sum().sort_values(ascending=False).reset_index()
            st.altair_chart(alt.Chart(leave_by_group).mark_bar().encode(
                x=alt.X('จำนวนวันลา:Q', title='รวมวันลา'),
                y=alt.Y('กลุ่มงาน:N', sort='-x', title='กลุ่มงาน'),
                tooltip=['กลุ่มงาน', 'จำนวนวันลา']
            ).properties(height=300), use_container_width=True)

    with col_chart2:
        st.markdown("##### **ผู้เดินทางราชการบ่อยที่สุด (Top 5)**")
        if not df_travel.empty and 'ชื่อ-สกุล' in df_travel.columns:
            top_travelers = df_travel['ชื่อ-สกุล'].value_counts().nlargest(5).reset_index()
            top_travelers.columns = ['ชื่อ-สกุล', 'จำนวนครั้ง']
            st.altair_chart(alt.Chart(top_travelers).mark_bar(color='#ff8c00').encode(
                x=alt.X('จำนวนครั้ง:Q', title='จำนวนครั้ง'),
                y=alt.Y('ชื่อ-สกุล:N', sort='-x', title='ชื่อ-สกุล'),
                tooltip=['ชื่อ-สกุล', 'จำนวนครั้ง']
            ).properties(height=300), use_container_width=True)

    st.markdown("---")
    st.markdown("#### **มุมมองสำหรับ HR**")

    hr_col1, hr_col2 = st.columns(2)

    with hr_col1:
        st.markdown("##### **สัดส่วนประเภทการลา**")
        if not df_leave.empty and 'ประเภทการลา' in df_leave.columns and 'จำนวนวันลา' in df_leave.columns:
            df_leave['จำนวนวันลา'] = pd.to_numeric(df_leave['จำนวนวันลา'], errors='coerce')
            df_leave_cleaned = df_leave.dropna(subset=['จำนวนวันลา'])
            leave_type_dist = df_leave_cleaned.groupby('ประเภทการลา')['จำนวนวันลา'].sum().reset_index()
            chart_leave_type = alt.Chart(leave_type_dist).mark_arc(innerRadius=50).encode(
                theta=alt.Theta(field="จำนวนวันลา", type="quantitative"),
                color=alt.Color(field="ประเภทการลา", type="nominal", title="ประเภทการลา"),
                tooltip=['ประเภทการลา', 'จำนวนวันลา']
            ).properties(height=300)
            st.altair_chart(chart_leave_type, use_container_width=True)

    with hr_col2:
        st.markdown("##### **แนวโน้มการลารายเดือน**")
        if not df_leave.empty and 'วันที่เริ่ม' in df_leave.columns and 'จำนวนวันลา' in df_leave.columns:
            df_leave_copy = df_leave.copy()
            df_leave_copy['วันที่เริ่ม'] = pd.to_datetime(df_leave_copy['วันที่เริ่ม'], errors='coerce')
            df_leave_copy['จำนวนวันลา'] = pd.to_numeric(df_leave_copy['จำนวนวันลา'], errors='coerce')
            df_leave_copy.dropna(subset=['วันที่เริ่ม', 'จำนวนวันลา'], inplace=True)
            df_leave_copy['เดือน'] = df_leave_copy['วันที่เริ่ม'].dt.strftime('%Y-%m')
            monthly_leave = df_leave_copy.groupby('เดือน')['จำนวนวันลา'].sum().reset_index()
            chart_monthly_trend = alt.Chart(monthly_leave).mark_line(point=True, strokeWidth=3).encode(
                x=alt.X('เดือน:T', title='เดือน'),
                y=alt.Y('จำนวนวันลา:Q', title='จำนวนวันลาสะสม'),
                tooltip=['เดือน', 'จำนวนวันลา']
            ).properties(height=300)
            st.altair_chart(chart_monthly_trend, use_container_width=True)

    st.markdown("##### **Top 5 พนักงานที่มาสาย/ขาด/ออกก่อนเวลา (เดือนล่าสุด)**")
    if not df_att.empty:
        # --- เตรียมข้อมูลวันที่ให้เป็น Datetime ---
        df_att_copy = df_att.copy()
        df_att_copy["วันที่"] = pd.to_datetime(df_att_copy["วันที่"], errors="coerce")
        df_leave['วันที่เริ่ม'] = pd.to_datetime(df_leave['วันที่เริ่ม'], errors='coerce')
        df_leave['วันที่สิ้นสุด'] = pd.to_datetime(df_leave['วันที่สิ้นสุด'], errors='coerce')
        # ✅ **เพิ่มการแปลงวันที่ของ df_travel**
        df_travel['วันที่เริ่ม'] = pd.to_datetime(df_travel['วันที่เริ่ม'], errors='coerce')
        df_travel['วันที่สิ้นสุด'] = pd.to_datetime(df_travel['วันที่สิ้นสุด'], errors='coerce')

        latest_month_str = df_att_copy["วันที่"].dt.strftime("%Y-%m").max()
        df_month = df_att_copy[df_att_copy["วันที่"].dt.strftime("%Y-%m") == latest_month_str]
        name_col = next((c for c in ["ชื่อ-สกุล", "ชื่อพนักงาน", "ชื่อ"] if c in df_month.columns), None)

        if name_col and latest_month_str:
            records = []
            WORK_START = dt.time(8, 30)
            all_days_in_month = pd.date_range(start=latest_month_str, end=pd.to_datetime(latest_month_str) + pd.offsets.MonthEnd(0), freq='D')
            
            for name in all_names:
                for day in all_days_in_month:
                    if day.weekday() >= 5: continue

                    # --- ตรวจสอบสถานะ ลา และ ไปราชการ ---
                    is_on_leave = not df_leave.empty and not df_leave[
                        (df_leave["ชื่อ-สกุล"] == name) & (df_leave["วันที่เริ่ม"] <= day) & (df_leave["วันที่สิ้นสุด"] >= day)
                    ].empty
                    
                    # ✅ **เพิ่มการตรวจสอบสถานะ "ไปราชการ"**
                    is_on_travel = not df_travel.empty and not df_travel[
                        (df_travel["ชื่อ-สกุล"] == name) & (df_travel["วันที่เริ่ม"] <= day) & (df_travel["วันที่สิ้นสุด"] >= day)
                    ].empty
                    
                    # ถ้าลา หรือ ไปราชการ ให้ข้ามไปเลย
                    if is_on_leave or is_on_travel: 
                        continue

                    # --- ตรวจสอบการมาทำงานจากข้อมูลสแกน ---
                    att_record = df_month[(df_month[name_col] == name) & (df_month['วันที่'].dt.date == day.date())]
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
                top_issues = df_issues['ชื่อพนักงาน'].value_counts().nlargest(5).reset_index()
                top_issues.columns = ['ชื่อพนักงาน', 'จำนวนครั้ง']
                chart_top_issues = alt.Chart(top_issues).mark_bar(color='indianred').encode(
                    x=alt.X('จำนวนครั้ง:Q', title='จำนวนครั้ง (สาย/ขาด)'),
                    y=alt.Y('ชื่อพนักงาน:N', sort='-x', title='ชื่อพนักงาน'),
                    tooltip=['ชื่อพนักงาน', 'จำนวนครั้ง']
                ).properties(height=300)
                st.altair_chart(chart_top_issues, use_container_width=True)
            else:
                st.info("ไม่พบข้อมูลการมาสายหรือขาดงานในเดือนล่าสุด")
                
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


