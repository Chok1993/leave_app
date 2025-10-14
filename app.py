# ====================================================
# 📋 โปรแกรมติดตามการลาและไปราชการ (สคร.9)
# ✅ Shared Drive + Admin + Dashboard + Attendance รวม (เวอร์ชันปรับปรุง)
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
# 👉 ID ของโฟลเดอร์ใน Google Drive ที่เก็บไฟล์ข้อมูล
FOLDER_ID = "1YFJZvs59ahRHmlRrKcQwepWJz6A-4B7d"

# ชื่อไฟล์มาตรฐาน
FILE_ATTEND = "attendance_report.xlsx" # สแกนนิ้ว
FILE_LEAVE  = "leave_report.xlsx"      # การลา
FILE_TRAVEL = "travel_report.xlsx"     # ไปราชการ (แก้ไขจาก scan_report.xlsx)

service = build("drive", "v3", credentials=creds)

# ===========================
# 🔧 Drive Helpers
# ===========================
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

def read_excel_from_drive(filename: str) -> pd.DataFrame:
    """อ่านไฟล์ Excel จาก Shared Drive; ถ้าไม่มีไฟล์ จะคืนค่า DataFrame ว่าง"""
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
    if file_id: # ถ้ามีไฟล์อยู่แล้ว ให้อัปเดต
        service.files().update(
            fileId=file_id, media_body=media, supportsAllDrives=True
        ).execute()
    else: # ถ้าไม่มี ให้สร้างไฟล์ใหม่
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
    """แปลงข้อมูลเป็น Date Object (รองรับหลาย format)"""
    if pd.isna(s): return pd.NaT
    try:
        return pd.to_datetime(s).date()
    except (ValueError, TypeError):
        return pd.NaT

def to_time(s):
    """แปลงข้อมูลเป็น Time Object (รองรับหลาย format)"""
    if pd.isna(s): return None
    try:
        return pd.to_datetime(str(s)).time()
    except (ValueError, TypeError):
        return None

# --- โหลดข้อมูลสแกน (Attendance) ---
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

# --- โหลดข้อมูลการลา (Leave) ---
df_leave = read_excel_from_drive(FILE_LEAVE)
if not df_leave.empty:
    for c in ['วันที่เริ่ม', 'วันที่สิ้นสุด']:
        if c in df_leave.columns:
            df_leave[c] = df_leave[c].apply(to_date)
else:
    df_leave = pd.DataFrame(columns=['ชื่อ-สกุล','กลุ่มงาน','ประเภทการลา','วันที่เริ่ม','วันที่สิ้นสุด','จำนวนวันลา','หมายเหตุ'])

# --- โหลดข้อมูลไปราชการ (Travel) ---
df_travel = read_excel_from_drive(FILE_TRAVEL)
if not df_travel.empty:
    for c in ['วันที่เริ่ม', 'วันที่สิ้นสุด']:
        if c in df_travel.columns:
            df_travel[c] = df_travel[c].apply(to_date)
else:
    df_travel = pd.DataFrame(columns=['ชื่อ-สกุล','กลุ่มงาน','กิจกรรม','สถานที่','วันที่เริ่ม','วันที่สิ้นสุด','จำนวนวัน','หมายเหตุ'])


# =================================================================
# 🧪 Helpers: ขยายช่วงวันที่ (Leave/Travel) ให้เป็นรายวัน
# =================================================================
def expand_date_range(df, start_col='วันที่เริ่ม', end_col='วันที่สิ้นสุด'):
    """ขยาย DataFrame ที่มีช่วงวันที่ ให้กลายเป็นข้อมูลรายวัน"""
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

daily_leave = expand_date_range(df_leave)
daily_travel = expand_date_range(df_travel)
daily_status = pd.concat([daily_leave, daily_travel]).drop_duplicates(subset=['ชื่อ-สกุล', 'วันที่'], keep='first')

# =================================================================
# 🧩 รวมข้อมูลทั้งหมด (Attendance + Leave + Travel)
# =================================================================
def determine_status(row, status_map):
    """(Helper) จัดลำดับความสำคัญและกำหนดสถานะการทำงานในแต่ละวัน"""
    # 1. ตรวจสอบจากข้อมูล ลา/ราชการ ก่อน
    status = status_map.get((row['ชื่อ-สกุล'], row['วันที่']))
    if status:
        return status
    # 2. ตรวจสอบว่าเป็นวันหยุดหรือไม่ (จากหมายเหตุ)
    if 'เสาร์' in str(row.get('หมายเหตุ', '')) or 'อาทิตย์' in str(row.get('หมายเหตุ', '')):
        return 'วันหยุด'
    # 3. ตรวจสอบสถานะการมาสาย
    is_late = str(row.get('สาย', '')).strip() not in ['', '0', '0:00', '00:00']
    if is_late:
        return 'สาย'
    # 4. ถ้ามีข้อมูลสแกน แต่ไม่เข้าเงื่อนไขอื่น = มาปกติ
    if pd.notna(row.get('เวลาเข้า')) or pd.notna(row.get('เวลาออก')):
        return 'มาปกติ'
    # 5. กรณีอื่นๆ
    return 'ไม่พบข้อมูล'

def build_attendance_view(month: int, year: int):
    """สร้างมุมมองการมาปฏิบัติงานรายเดือน"""
    start_date = dt.date(year, month, 1)
    end_date = (start_date + dt.timedelta(days=32)).replace(day=1) - dt.timedelta(days=1)

    # กรองข้อมูลตามเดือนที่เลือก
    att_m = df_att[(df_att['วันที่'] >= start_date) & (df_att['วันที่'] <= end_date)].copy() if not df_att.empty else df_att.copy()
    status_m = daily_status[(daily_status['วันที่'] >= start_date) & (daily_status['วันที่'] <= end_date)]

    # สร้าง status map เพื่อการค้นหาที่รวดเร็ว
    status_map = { (r['ชื่อ-สกุล'], r['วันที่']): r['สถานะ'] for _, r in status_m.iterrows() }

    # กำหนดสถานะ
    att_m['สถานะ'] = att_m.apply(determine_status, args=(status_map,), axis=1)
    att_m = att_m.sort_values(['ชื่อ-สกุล', 'วันที่'])

    # สร้างตารางสรุป
    summary = (att_m.groupby(['ชื่อ-สกุล', 'สถานะ'], dropna=False)
               .size().reset_index(name='จำนวนวัน'))
    pivot = summary.pivot_table(index='ชื่อ-สกุล', columns='สถานะ', values='จำนวนวัน', aggfunc='sum', fill_value=0).reset_index()

    # (ปรับปรุง) นับจำนวนครั้งที่มาสายจากคอลัมน์ 'สถานะ' โดยตรง
    if 'สาย' in pivot.columns:
        pivot = pivot.rename(columns={'สาย': 'จำนวนครั้งมาสาย'})
    else:
        pivot['จำนวนครั้งมาสาย'] = 0

    return att_m, pivot


# ====================================================
# 🎯 UI Constants
# ====================================================
staff_groups = sorted([
    "กลุ่มโรคติดต่อ", "กลุ่มระบาดวิทยาและตอบโต้ภาวะฉุกเฉินทางสาธารณสุข",
    "กลุ่มพัฒนาองค์กร", "กลุ่มบริหารทั่วไป", "กลุ่มโรคไม่ติดต่อ",
    "กลุ่มห้องปฏิบัติการทางการแพทย์", "กลุ่มพัฒนานวัตกรรมและวิจัย",
    "กลุ่มโรคติดต่อเรื้อรัง", "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.1 จ.ชัยภูมิ",
    "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.2 จ.บุรีรัมย์",
    "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.3 จ.สุรินทร์",
    "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.4 ปากช่อง",
    "ด่านควบคุมโรคช่องจอม จ.สุรินทร์", "ศูนย์บริการเวชศาสตร์ป้องกัน",
    "กลุ่มสื่อสารความเสี่ยง", "กลุ่มโรคจากการประกอบอาชีพและสิ่งแวดล้อม"
])
leave_types = ["ลาป่วย", "ลากิจ", "ลาพักผ่อน", "อื่นๆ"]

# ====================================================
# 🧭 Navigation & App Body
# ====================================================
st.title("📋 ระบบติดตามการลา ไปราชการ และการมาปฏิบัติงาน (สคร.9)")

menu = st.sidebar.radio(
    "เลือกเมนู",
    ["📊 Dashboard", "📅 การมาปฏิบัติงาน", "🧭 การไปราชการ", "🕒 การลา", "🧑‍💼 ผู้ดูแลระบบ"]
)

# --- 📊 Dashboard ---
if menu == "📊 Dashboard":
    st.header("📊 ภาพรวมข้อมูล")
    col1, col2, col3 = st.columns(3)
    col1.metric("รายการไปราชการทั้งหมด", len(df_travel))
    col2.metric("รายการลาทั้งหมด", len(df_leave))
    col3.metric("ข้อมูลสแกน (แถว)", len(df_att))
    st.markdown("---")

    if not df_leave.empty and 'ประเภทการลา' in df_leave.columns:
        st.subheader("📌 สรุปประเภทการลา")
        leave_counts = df_leave['ประเภทการลา'].value_counts().reset_index()
        leave_counts.columns = ['ประเภทการลา', 'จำนวนครั้ง']
        chart = alt.Chart(leave_counts).mark_bar().encode(
            x=alt.X('ประเภทการลา:N', sort=None, title='ประเภทการลา'),
            y=alt.Y('จำนวนครั้ง:Q', title='จำนวนครั้ง'),
            color='ประเภทการลา:N',
            tooltip=['ประเภทการลา', 'จำนวนครั้ง']
        ).properties(title='จำนวนการลาแยกตามประเภท')
        st.altair_chart(chart, use_container_width=True)

    if not df_att.empty and not df_att['วันที่'].isna().all():
        st.subheader("📈 ปริมาณข้อมูลสแกนรายวัน")
        att_counts = df_att.dropna(subset=['วันที่']).groupby('วันที่').size().reset_index(name='จำนวนแถวข้อมูล')
        line_chart = alt.Chart(att_counts).mark_line(point=True).encode(
            x=alt.X('วันที่:T', title='วันที่'),
            y=alt.Y('จำนวนแถวข้อมูล:Q', title='จำนวนแถวข้อมูล'),
            tooltip=['วันที่:T', 'จำนวนแถวข้อมูล:Q']
        ).properties(title='แนวโน้มข้อมูลการสแกนลายนิ้วมือ')
        st.altair_chart(line_chart, use_container_width=True)

# --- 📅 Attendance View ---
elif menu == "📅 การมาปฏิบัติงาน":
    st.header("📅 สรุปการมาปฏิบัติงานรายเดือน")
    today = dt.date.today()
    colf1, colf2 = st.columns([1, 1])
    sel_month = colf1.selectbox("เลือกเดือน", range(1, 13), index=today.month-1, format_func=lambda m: f"{m:02d}")
    sel_year = colf2.number_input("เลือกปี (ค.ศ.)", value=today.year, min_value=2020, max_value=2050)

    att_month, summary = build_attendance_view(sel_month, sel_year)

    st.subheader("📊 สรุปต่อบุคคล (จำนวนวัน/สถานะ)")
    st.dataframe(summary, use_container_width=True)

    with st.expander("แสดงข้อมูลรายวัน (Daily View)"):
        st.dataframe(att_month.astype(str), use_container_width=True, height=420)

    if not summary.empty:
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
            summary.to_excel(writer, sheet_name="Summary", index=False)
            att_month.to_excel(writer, sheet_name="Daily", index=False)
        out.seek(0)
        st.download_button(
            "⬇️ ดาวน์โหลดสรุป (Excel)", data=out,
            file_name=f"attendance_summary_{sel_year}_{sel_month:02d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# --- 🧭 Travel Form ---
elif menu == "🧭 การไปราชการ":
    st.header("🧭 บันทึกการไปราชการ")
    with st.form("form_travel", clear_on_submit=True):
        data = {
            "ชื่อ-สกุล": st.text_input("ชื่อ-สกุล"),
            "กลุ่มงาน": st.selectbox("กลุ่มงาน", staff_groups),
            "กิจกรรม": st.text_input("กิจกรรม/โครงการ"),
            "สถานที่": st.text_input("สถานที่"),
            "วันที่เริ่ม": st.date_input("วันที่เริ่ม", dt.date.today()),
            "วันที่สิ้นสุด": st.date_input("วันที่สิ้นสุด", dt.date.today())
        }
        submitted = st.form_submit_button("💾 บันทึกข้อมูล")

    if submitted:
        if not data["ชื่อ-สกุล"]:
            st.warning("กรุณากรอก 'ชื่อ-สกุล'")
        elif data["วันที่เริ่ม"] > data["วันที่สิ้นสุด"]:
            st.error("'วันที่เริ่ม' ต้องมาก่อน 'วันที่สิ้นสุด'")
        else:
            data["จำนวนวัน"] = (data["วันที่สิ้นสุด"] - data["วันที่เริ่ม"]).days + 1
            df_travel_new = pd.concat([df_travel, pd.DataFrame([data])], ignore_index=True)
            write_excel_to_drive(FILE_TRAVEL, df_travel_new)
            st.success("✅ บันทึกข้อมูลไปราชการเรียบร้อยแล้ว!")

    st.markdown("--- \n ### 📋 ข้อมูลปัจจุบัน")
    st.dataframe(df_travel.astype(str).sort_values('วันที่เริ่ม', ascending=False), use_container_width=True, height=420)

# --- 🕒 Leave Form ---
elif menu == "🕒 การลา":
    st.header("🕒 บันทึกการลา")
    with st.form("form_leave", clear_on_submit=True):
        data = {
            "ชื่อ-สกุล": st.text_input("ชื่อ-สกุล"),
            "กลุ่มงาน": st.selectbox("กลุ่มงาน", staff_groups),
            "ประเภทการลา": st.selectbox("ประเภทการลา", leave_types),
            "วันที่เริ่ม": st.date_input("วันที่เริ่ม", dt.date.today()),
            "วันที่สิ้นสุด": st.date_input("วันที่สิ้นสุด", dt.date.today())
        }
        submitted = st.form_submit_button("💾 บันทึกข้อมูล")

    if submitted:
        if not data["ชื่อ-สกุล"]:
            st.warning("กรุณากรอก 'ชื่อ-สกุล'")
        elif data["วันที่เริ่ม"] > data["วันที่สิ้นสุด"]:
            st.error("'วันที่เริ่ม' ต้องมาก่อน 'วันที่สิ้นสุด'")
        else:
            data["จำนวนวันลา"] = (data["วันที่สิ้นสุด"] - data["วันที่เริ่ม"]).days + 1
            df_leave_new = pd.concat([df_leave, pd.DataFrame([data])], ignore_index=True)
            write_excel_to_drive(FILE_LEAVE, df_leave_new)
            st.success("✅ บันทึกข้อมูลการลาเรียบร้อยแล้ว!")

    st.markdown("--- \n ### 📋 ข้อมูลปัจจุบัน")
    st.dataframe(df_leave.astype(str).sort_values('วันที่เริ่ม', ascending=False), use_container_width=True, height=420)

# --- 🧑‍💼 Admin Panel ---
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
        edited_leave = st.data_editor(df_leave, num_rows="dynamic", use_container_width=True, key="ed_leave")
        if st.button("💾 บันทึกข้อมูลการลา", key="save_leave"):
            write_excel_to_drive(FILE_LEAVE, edited_leave)
            st.success("✅ บันทึกข้อมูลการลาเรียบร้อย")
            st.rerun()

    with tabB:
        st.caption("แก้ไขตารางด้านล่างได้โดยตรง (เพิ่ม/ลบ/แก้ไข) แล้วกดปุ่มบันทึก")
        edited_travel = st.data_editor(df_travel, num_rows="dynamic", use_container_width=True, key="ed_travel")
        if st.button("💾 บันทึกข้อมูลไปราชการ", key="save_travel"):
            write_excel_to_drive(FILE_TRAVEL, edited_travel)
            st.success("✅ บันทึกข้อมูลไปราชการเรียบร้อย")
            st.rerun()

    with tabC:
        st.caption("ข้อมูลสแกนมีขนาดใหญ่ แนะนำให้แก้ไขเท่าที่จำเป็น (เช่น เติมหมายเหตุ)")
        edited_att = st.data_editor(df_att, num_rows="dynamic", use_container_width=True, key="ed_att")
        if st.button("💾 บันทึกข้อมูลสแกน", key="save_att"):
            write_excel_to_drive(FILE_ATTEND, edited_att)
            st.success("✅ บันทึกข้อมูลสแกนเรียบร้อย")
            st.rerun()
