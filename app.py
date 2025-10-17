# ====================================================
# 📋 โปรแกรมติดตามการลาและไปราชการ (สคร.9)
# ✅ Final Version: โค้ดฉบับสมบูรณ์
# ====================================================

import io
import mimetypes
import altair as alt
import datetime as dt
import pandas as pd
import numpy as np
import streamlit as st

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload, MediaFileUpload

# ===========================
# 🔐 Auth & App Config
# ===========================
st.set_page_config(page_title="สคร.9 - ติดตามการลา/ราชการ/สแกน", layout="wide")

creds = service_account.Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=["https://www.googleapis.com/auth/drive"]
)
ADMIN_PASSWORD = st.secrets.get("admin_password", "admin123")

# ===========================
# 🗂️ Shared Drive Config
# ===========================
FOLDER_ID = "1YFJZvs59ahRHmlRrKcQwepWJz6A-4B7d"
ATTACHMENT_FOLDER_NAME = "เอกสารแนบ_ไปราชการ"
FILE_ATTEND = "scan_report.xlsx"
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

def get_file_id(filename: str, parent_id=FOLDER_ID):
    """หา ID ของไฟล์หรือโฟลเดอร์ใน Parent ที่กำหนด"""
    q = f"name='{filename}' and '{parent_id}' in parents and trashed=false"
    res = service.files().list(q=q, fields="files(id,name)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    files = res.get("files", [])
    return files[0]["id"] if files else None

def write_excel_to_drive(filename: str, df: pd.DataFrame):
    """บันทึก DataFrame กลับไปยังไฟล์ Excel บน Shared Drive"""
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
    """สร้างไฟล์สำรองข้อมูลพร้อมประทับเวลา"""
    if df.empty: return
    now = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    backup_filename = f"backup_{now}_{original_filename}"
    write_excel_to_drive(backup_filename, df)

@st.cache_resource
def get_or_create_folder(folder_name, parent_folder_id):
    """หา ID ของโฟลเดอร์ ถ้าไม่มีให้สร้างใหม่"""
    folder_id = get_file_id(folder_name, parent_id=parent_folder_id)
    if folder_id:
        return folder_id
    else:
        file_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_folder_id]
        }
        folder = service.files().create(body=file_metadata, fields='id', supportsAllDrives=True).execute()
        return folder.get('id')

def upload_pdf_to_drive(file_object, filename, folder_id):
    """อัปโหลดไฟล์ PDF และคืนค่าเป็น View Link"""
    if file_object is None:
        return "-"
    
    file_object.seek(0)
    media = MediaIoBaseUpload(file_object, mimetype='application/pdf', resumable=True)
    file_metadata = {'name': filename, 'parents': [folder_id]}
    
    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id, webViewLink',
        supportsAllDrives=True
    ).execute()
    
    permission = {'type': 'anyone', 'role': 'reader'}
    service.permissions().create(fileId=file.get('id'), body=permission, supportsAllDrives=True).execute()
    
    return file.get('webViewLink')

def count_weekdays(start_date, end_date):
    """นับจำนวนวันทำการ (จันทร์-ศุกร์) ระหว่างวันที่สองวัน แบบรวมวันสิ้นสุด"""
    if start_date and end_date and start_date <= end_date:
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

all_names = sorted(list(set(pd.concat([
    df_leave['ชื่อ-สกุล'] if 'ชื่อ-สกุล' in df_leave.columns else pd.Series(dtype='str'),
    df_travel['ชื่อ-สกุล'] if 'ชื่อ-สกุล' in df_travel.columns else pd.Series(dtype='str'),
    df_att['ชื่อ-สกุล'] if 'ชื่อ-สกุล' in df_att.columns else pd.Series(dtype='str')
]).dropna())))

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
    st.info("💡 ระบบนี้ใช้สำหรับบันทึกและสรุปข้อมูลบุคลากรใน สคร.9\n"
            "ได้แก่ การลา การไปราชการ และการมาปฏิบัติงาน พร้อมแนบไฟล์เอกสาร PDF ได้โดยตรง")
    st.image("https://ddc.moph.go.th/uploads/files/11120210817094038.jpg",
             caption="สำนักงานป้องกันควบคุมโรคที่ 9 นครราชสีมา", use_container_width=True)

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
            st.altair_chart(alt.Chart(leave_by_group).mark_bar().encode(x=alt.X('จำนวนวันลา:Q', title='รวมวันลา'), y=alt.Y('กลุ่มงาน:N', sort='-x', title='กลุ่มงาน'), tooltip=['กลุ่มงาน', 'จำนวนวันลา']).properties(height=300), use_container_width=True)
    with col_chart2:
        st.markdown("##### **ผู้เดินทางราชการบ่อยที่สุด (Top 5)**")
        if not df_travel.empty and 'ชื่อ-สกุล' in df_travel.columns:
            top_travelers = df_travel['ชื่อ-สกุล'].value_counts().nlargest(5).reset_index()
            top_travelers.columns = ['ชื่อ-สกุล', 'จำนวนครั้ง']
            st.altair_chart(alt.Chart(top_travelers).mark_bar(color='#ff8c00').encode(x=alt.X('จำนวนครั้ง:Q', title='จำนวนครั้ง'), y=alt.Y('ชื่อ-สกุล:N', sort='-x', title='ชื่อ-สกุล'), tooltip=['ชื่อ-สกุล', 'จำนวนครั้ง']).properties(height=300), use_container_width=True)

import streamlit as st
import pandas as pd
import datetime as dt
import io

# ----------------------------
# 🧩 ฟังก์ชันอ่านไฟล์ Excel (เลือกชีตอัตโนมัติ)
# ----------------------------
def read_scan_report(file_path):
    try:
        xls = pd.ExcelFile(file_path)
        valid_sheets = []
        for sheet in xls.sheet_names:
            df_temp = pd.read_excel(xls, sheet_name=sheet, nrows=5)
            if df_temp.dropna(how="all").shape[0] > 0:
                valid_sheets.append(sheet)
        if not valid_sheets:
            st.error("❌ ไม่พบชีตที่มีข้อมูลในไฟล์ scan_report.xlsx")
            return pd.DataFrame()
        df = pd.read_excel(xls, sheet_name=valid_sheets[0])
        st.success(f"📄 ใช้ชีต: {valid_sheets[0]}")
        return df
    except Exception as e:
        st.error(f"⚠️ อ่านไฟล์ scan_report.xlsx ไม่สำเร็จ: {e}")
        return pd.DataFrame()

# ====================================================
# 📂 โหลดไฟล์หลัก (Smart Cache + ปุ่มรีเฟรช + เวลา Sync)
# ====================================================

import os
import shutil
import datetime as dt

LOCAL_CACHE_DIR = "cached_files"
os.makedirs(LOCAL_CACHE_DIR, exist_ok=True)

# 🔁 ปุ่มรีเฟรชข้อมูลจาก Drive
st.sidebar.markdown("---")
if st.sidebar.button("🔁 รีเฟรชข้อมูลจาก Drive (อัปเดตล่าสุด)"):
    try:
        shutil.rmtree(LOCAL_CACHE_DIR)
        os.makedirs(LOCAL_CACHE_DIR, exist_ok=True)
        st.sidebar.success("✅ ล้าง cache เรียบร้อย กำลังโหลดข้อมูลใหม่...")
        st.experimental_rerun()
    except Exception as e:
        st.sidebar.error(f"⚠️ ไม่สามารถล้าง cache ได้: {e}")

# 🕒 ฟังก์ชันจัดการเวลาซิงก์ล่าสุด
def update_sync_time():
    sync_file = os.path.join(LOCAL_CACHE_DIR, "last_sync.txt")
    with open(sync_file, "w", encoding="utf-8") as f:
        f.write(dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

def get_sync_time():
    sync_file = os.path.join(LOCAL_CACHE_DIR, "last_sync.txt")
    if os.path.exists(sync_file):
        with open(sync_file, "r", encoding="utf-8") as f:
            return f.read().strip()
    return "— ยังไม่เคยซิงก์ข้อมูล —"

def load_excel_smart_cache(filename, from_drive=True):
    """
    โหลดไฟล์ Excel แบบอัจฉริยะ:
    1️⃣ ถ้ามีไฟล์ใน cache local → ใช้เลย
    2️⃣ ถ้าไม่มี → ดึงจาก Shared Drive
    3️⃣ ถ้าดึงสำเร็จ → เก็บสำเนา local และอัปเดตเวลาซิงก์
    """
    local_path = os.path.join(LOCAL_CACHE_DIR, filename)

    # ✅ 1. โหลดจาก cache ถ้ามี
    if os.path.exists(local_path):
        st.success(f"📄 โหลดไฟล์ {filename} จาก cache local สำเร็จ ✅")
        return pd.read_excel(local_path)

    # ✅ 2. โหลดจาก Drive ถ้าไม่มีใน local
    elif from_drive:
        st.info(f"🔄 ไม่พบไฟล์ {filename} ในเครื่อง — กำลังดึงจาก Shared Drive...")
        df = read_excel_from_drive(filename)
        if df.empty:
            st.error(f"❌ ไม่พบไฟล์ {filename} ใน Shared Drive")
            return pd.DataFrame()

        # ✅ 3. เก็บสำเนา cache และอัปเดตเวลาซิงก์
        try:
            df.to_excel(local_path, index=False)
            update_sync_time()
            st.success(f"✅ โหลดไฟล์ {filename} จาก Shared Drive และเก็บ cache ไว้ใน '{LOCAL_CACHE_DIR}'")
        except Exception as e:
            st.warning(f"⚠️ โหลดสำเร็จแต่บันทึก cache ไม่ได้: {e}")
        return df

    # ❌ 4. ไม่พบทั้งสองที่
    else:
        st.error(f"❌ ไม่พบไฟล์ {filename} ทั้งใน local และ Shared Drive")
        return pd.DataFrame()

# ✅ โหลดข้อมูลทั้งสามชุด
df_att = load_excel_smart_cache("scan_report.xlsx")
df_leave = load_excel_smart_cache("leave_report.xlsx")
df_travel = load_excel_smart_cache("travel_report.xlsx")

# 🕒 แสดงเวลาซิงก์ล่าสุด
st.sidebar.caption(f"🕒 ซิงก์ข้อมูลล่าสุด: {get_sync_time()}")

# ----------------------------
# เมนูหลัก
# ----------------------------
menu = st.sidebar.radio("เลือกเมนู", ["📅 การมาปฏิบัติงาน"])

# ----------------------------
# 📅 การมาปฏิบัติงาน (ตรวจไขว้ 3 ฐาน + เวลาเข้าออกจริง)
# ----------------------------
if menu == "📅 การมาปฏิบัติงาน":
    st.header("📅 สรุปการมาปฏิบัติงานรายวัน (ตรวจจากสแกน + ลา + ราชการ + เวลาเข้า-ออกจริง)")

    if df_att.empty:
        st.warning("⚠️ ยังไม่มีข้อมูลสแกนเข้า-ออกในระบบ หรือไม่พบไฟล์ scan_report.xlsx")
        st.stop()

    # ✅ ตรวจหาชื่อคอลัมน์บุคลากร
    name_candidates = [c for c in df_att.columns if "ชื่อ" in str(c)]
    if len(name_candidates) == 0:
        st.error("❌ ไม่พบคอลัมน์ชื่อบุคลากร (เช่น 'ชื่อพนักงาน')")
        st.stop()
    name_col = name_candidates[0]
    df_att[name_col] = df_att[name_col].astype(str).str.strip()
    valid_names = sorted([n for n in df_att[name_col].unique() if n and n.lower() != "nan"])

    # ✅ แปลงวันที่อย่างยืดหยุ่น
    if "วันที่" not in df_att.columns:
        st.error("❌ ไม่พบคอลัมน์ 'วันที่' ในไฟล์ scan_report.xlsx")
        st.stop()

    st.write("🧩 ตัวอย่างข้อมูลวันที่ (ก่อนแปลง):", df_att["วันที่"].head(10).tolist())

    def parse_date_flex(x):
        for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y"):
            try:
                return pd.to_datetime(x, format=fmt)
            except Exception:
                continue
        return pd.NaT

    df_att["วันที่"] = df_att["วันที่"].apply(parse_date_flex)
    st.write("✅ ตรวจสอบวันที่หลังแปลง:", df_att["วันที่"].head(10))
    st.write("📘 dtype วันที่:", df_att["วันที่"].dtype)

    # ✅ แปลงเวลาเข้า-ออก
    for col in ["เวลาเข้า", "เวลาออก"]:
        if col in df_att.columns:
            df_att[col] = pd.to_datetime(df_att[col], errors="coerce").dt.time

    # ✅ แปลงวันที่ในใบลา/ไปราชการ
    for df in [df_leave, df_travel]:
        for c in ["วันที่เริ่ม", "วันที่สิ้นสุด"]:
            if c in df.columns:
                df[c] = pd.to_datetime(df[c], errors="coerce")

    # ✅ ตัวเลือกเดือนและชื่อ
    df_att["เดือน"] = df_att["วันที่"].dt.strftime("%Y-%m")
    months = sorted(df_att["เดือน"].dropna().unique())
    if not months:
        st.warning("⚠️ ไม่พบข้อมูลวันที่ในไฟล์ scan_report.xlsx")
        st.stop()

    selected_month = st.selectbox("📅 เลือกเดือนที่ต้องการดู", months, index=len(months)-1)
    selected_names = st.multiselect("👤 เลือกชื่อบุคลากร", valid_names, max_selections=5)

    df_month = df_att[df_att["เดือน"] == selected_month].copy()
    if selected_names:
        df_month = df_month[df_month[name_col].isin(selected_names)]
    if df_month.empty:
        st.info("ℹ️ ไม่มีข้อมูลสแกนในเดือนที่เลือก")
        st.stop()

    # ✅ ช่วงวันที่ในเดือน
    month_start = df_month["วันที่"].min().date().replace(day=1)
    month_end = (month_start + pd.offsets.MonthEnd(0)).date()
    date_range = pd.date_range(month_start, month_end, freq="D")

    records = []
    for name in df_month[name_col].unique():
        for d in date_range:
            rec = {"ชื่อพนักงาน": name, "วันที่": d.date(), "เวลาเข้า": "", "เวลาออก": "", "สถานะ": "", "หมายเหตุ": ""}
            att = df_month[(df_month[name_col] == name) & (df_month["วันที่"].dt.date == d.date())]

            if not att.empty:
                rec["เวลาเข้า"] = att.iloc[0].get("เวลาเข้า", "")
                rec["เวลาออก"] = att.iloc[0].get("เวลาออก", "")
                note = str(att.iloc[0].get("หมายเหตุ", "")).strip()
                rec["หมายเหตุ"] = note

                if d.weekday() >= 5:
                    rec["สถานะ"] = "วันหยุด"
                else:
                    try:
                        t_in = att.iloc[0].get("เวลาเข้า", None)
                        t_out = att.iloc[0].get("เวลาออก", None)
                        if not t_in and not t_out:
                            rec["สถานะ"] = "ขาดงาน"
                        elif t_in and (t_in > dt.time(8, 30)):
                            rec["สถานะ"] = "มาสาย"
                        elif t_out and (t_out < dt.time(16, 30)):
                            rec["สถานะ"] = "ออกก่อน"
                        else:
                            rec["สถานะ"] = "มาปกติ"
                    except Exception:
                        rec["สถานะ"] = "ข้อมูลไม่สมบูรณ์"
            else:
                # ❌ ไม่มีข้อมูลสแกน — ตรวจลา/ราชการ
                in_leave = (
                    not df_leave.empty and
                    (df_leave["ชื่อ-สกุล"] == name).any() and
                    (df_leave[
                        (df_leave["ชื่อ-สกุล"] == name) &
                        (df_leave["วันที่เริ่ม"] <= d) &
                        (df_leave["วันที่สิ้นสุด"] >= d)
                    ].shape[0] > 0)
                )
                in_travel = (
                    not df_travel.empty and
                    (df_travel["ชื่อ-สกุล"] == name).any() and
                    (df_travel[
                        (df_travel["ชื่อ-สกุล"] == name) &
                        (df_travel["วันที่เริ่ม"] <= d) &
                        (df_travel["วันที่สิ้นสุด"] >= d)
                    ].shape[0] > 0)
                )

                if in_leave:
                    leave_type = df_leave.loc[
                        (df_leave["ชื่อ-สกุล"] == name) &
                        (df_leave["วันที่เริ่ม"] <= d) &
                        (df_leave["วันที่สิ้นสุด"] >= d),
                        "ประเภทการลา"
                    ].iloc[0]
                    rec["สถานะ"] = f"ลา ({leave_type})"
                elif in_travel:
                    rec["สถานะ"] = "ไปราชการ"
                elif d.weekday() >= 5:
                    rec["สถานะ"] = "วันหยุด"
                else:
                    rec["สถานะ"] = "ขาดงาน"

            records.append(rec)

    df_daily = pd.DataFrame(records).sort_values(["ชื่อพนักงาน", "วันที่"])

    # 🎨 สีสถานะ
    def color_status(val):
        colors = {
            "มาปกติ": "background-color:#d4edda",
            "มาสาย": "background-color:#fff3cd",
            "ออกก่อน": "background-color:#f8d7da",
            "ลา": "background-color:#d1ecf1",
            "ไปราชการ": "background-color:#cce5ff",
            "วันหยุด": "background-color:#e2e3e5",
            "ขาดงาน": "background-color:#f5c6cb",
        }
        for key in colors:
            if key in str(val):
                return colors[key]
        return ""

    st.markdown("### 📋 ตารางสรุปสถานะรายวัน")
    st.dataframe(df_daily.style.applymap(color_status, subset=["สถานะ"]), use_container_width=True, height=600)

    # 📤 ดาวน์โหลดรายวัน
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_daily.to_excel(writer, index=False, sheet_name="รายวัน")
    output.seek(0)
    st.download_button(
        "📥 ดาวน์โหลดรายงานรายวัน (Excel)",
        data=output,
        file_name=f"รายงานมาปฏิบัติงาน_{selected_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # --------------------------------------------------
    # 📊 สรุปสถิติรวมต่อเดือน
    # --------------------------------------------------
    st.markdown("---")
    st.subheader("📊 สรุปสถิติรวมต่อเดือนต่อคน")

    def simplify_status(s):
        if isinstance(s, str):
            if s.startswith("ลา"):
                return "ลา"
            return s
        return s

    df_daily["สถานะย่อ"] = df_daily["สถานะ"].apply(simplify_status)

    summary = (
        df_daily.groupby(["ชื่อพนักงาน", "สถานะย่อ"])
        .size()
        .reset_index(name="จำนวนวัน")
        .pivot(index="ชื่อพนักงาน", columns="สถานะย่อ", values="จำนวนวัน")
        .fillna(0)
        .astype(int)
        .reset_index()
    )

    summary["รวมทั้งหมด"] = summary.drop(columns=["ชื่อพนักงาน"]).sum(axis=1)
    preferred_order = ["มาปกติ", "มาสาย", "ออกก่อน", "ลา", "ไปราชการ", "ขาดงาน", "วันหยุด", "รวมทั้งหมด"]
    cols_present = [c for c in preferred_order if c in summary.columns]
    summary = summary[["ชื่อพนักงาน"] + cols_present]

    st.dataframe(summary, use_container_width=True)

    excel_output = io.BytesIO()
    with pd.ExcelWriter(excel_output, engine="xlsxwriter") as writer:
        df_daily.to_excel(writer, index=False, sheet_name="รายวัน")
        summary.to_excel(writer, index=False, sheet_name="สรุปสถิติรวม")
    excel_output.seek(0)
    st.download_button(
        "📥 ดาวน์โหลดรายงานสรุป (รายวัน + รวมต่อเดือน)",
        data=excel_output,
        file_name=f"สรุปมาปฏิบัติงาน_{selected_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
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
        days = count_weekdays(st.session_state.travel_start_date, st.session_state.travel_end_date)
        if days > 0: st.caption(f"🗓️ รวมเฉพาะวันทำการ {days} วัน")
        selected_names = st.multiselect("เลือกชื่อเจ้าหน้าที่ (เลือกได้หลายคน)", options=all_names, disabled=st.session_state.submitted)
        new_names_str = st.text_area("เพิ่มชื่อเจ้าหน้าที่ใหม่ (กรณีไม่มีในตัวเลือก)", placeholder="ใส่ 1 ชื่อต่อ 1 บรรทัด", disabled=st.session_state.submitted)
        uploaded_file = st.file_uploader("แนบไฟล์คำสั่ง/เอกสารอนุมัติ (PDF)", type="pdf", disabled=st.session_state.submitted)
        submitted_travel = st.form_submit_button("💾 บันทึกข้อมูล", on_click=callback_submit, disabled=st.session_state.submitted)

    if submitted_travel:
        final_names = list(set(selected_names + [name.strip() for name in new_names_str.split('\n') if name.strip()]))
        if not final_names:
            st.warning("กรุณาเลือกหรือกรอก 'ชื่อ-สกุล' อย่างน้อย 1 คน")
            st.session_state.submitted = False
        elif common_data_ui["วันที่เริ่ม"] > common_data_ui["วันที่สิ้นสุด"]:
            st.error("'วันที่เริ่ม' ต้องมาก่อน 'วันที่สิ้นสุด'")
            st.session_state.submitted = False
        else:
            with st.spinner('⏳ กำลังบันทึกและอัปโหลดไฟล์...'):
                attachment_folder_id = get_or_create_folder(ATTACHMENT_FOLDER_NAME, FOLDER_ID)
                file_link = "-"
                if uploaded_file:
                    pdf_filename = f'{common_data_ui["วันที่เริ่ม"].strftime("%Y-%m-%d")}_{common_data_ui["กิจกรรม"].replace(" ", "_")}_{final_names[0].replace(" ", "_")}.pdf'
                    file_link = upload_pdf_to_drive(uploaded_file, pdf_filename, attachment_folder_id)
                backup_excel(FILE_TRAVEL, df_travel)
                new_records = []
                timestamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
                num_days = count_weekdays(common_data_ui["วันที่เริ่ม"], common_data_ui["วันที่สิ้นสุด"])
                for name in final_names:
                    fellow_travelers = ", ".join([other for other in final_names if other != name])
                    new_records.append({**common_data_ui, "ชื่อ-สกุล": name, "จำนวนวัน": num_days, "ผู้ร่วมเดินทาง": fellow_travelers or "-", "ลิงก์เอกสาร": file_link, "last_update": timestamp})
                if new_records:
                    df_travel_new = pd.concat([df_travel, pd.DataFrame(new_records)], ignore_index=True)
                    write_excel_to_drive(FILE_TRAVEL, df_travel_new)
                    st.success(f"✅ บันทึกข้อมูลสำเร็จ!")
                    st.session_state.submitted = False
                    st.rerun()
    st.markdown("---")
    st.markdown("### 🔍 ค้นหาข้อมูลรายบุคคล")
    search_name_travel = st.text_input("พิมพ์ชื่อ-สกุลเพื่อค้นหา (ไปราชการ)", "")
    df_display_travel = df_travel[df_travel['ชื่อ-สกุล'].str.contains(search_name_travel, case=False, na=False)] if search_name_travel else df_travel
    st.dataframe(df_display_travel.astype(str).sort_values('วันที่เริ่ม', ascending=False), column_config={"ลิงก์เอกสาร": st.column_config.LinkColumn("เอกสารแนบ", display_text="🔗 เปิดไฟล์")})

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
        days = count_weekdays(st.session_state.leave_start_date, st.session_state.leave_end_date)
        if days > 0: st.caption(f"🗓️ รวมเฉพาะวันทำการ {days} วัน")
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
    df_display_leave = df_leave[df_leave['ชื่อ-สกุล'].str.contains(search_name_leave, case=False, na=False)] if search_name_leave else df_leave
    st.dataframe(df_display_leave.astype(str).sort_values('วันที่เริ่ม', ascending=False))

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










