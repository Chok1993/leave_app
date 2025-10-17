# ====================================================
# 📋 โปรแกรมติดตามการลาและไปราชการ (สคร.9)
# ✅ Shared Drive + Admin + Dashboard + Attendance รวม (เวอร์ชันแก้ไข Heatmap)
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
FILE_LEAVE  = "leave_report.xlsx"
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
    df_travel = pd.DataFrame(columns=['ชื่อ-สกุล','กลุ่มงาน','กิจกรรม','สถานที่','วันที่เริ่ม','วันที่สิ้นสุด','จำนวนวัน','หมายเหตุ'])

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
    if status:
        return status
    if 'เสาร์' in str(row.get('หมายเหตุ', '')) or 'อาทิตย์' in str(row.get('หมายเหตุ', '')):
        return 'วันหยุด'
    is_late = str(row.get('สาย', '')).strip() not in ['', '0', '0:00', '00:00', 'None']
    if is_late:
        return 'สาย'
    if pd.notna(row.get('เวลาเข้า')) or pd.notna(row.get('เวลาออก')):
        return 'มาปกติ'
    return 'ไม่พบข้อมูล'

def build_attendance_view(month: int, year: int):
    start_date = dt.date(year, month, 1)
    end_date = (start_date + dt.timedelta(days=32)).replace(day=1) - dt.timedelta(days=1)

    att_m = df_att[(df_att['วันที่'] >= start_date) & (df_att['วันที่'] <= end_date)].copy() if not df_att.empty else df_att.copy()
    status_m = daily_status[(daily_status['วันที่'] >= start_date) & (daily_status['วันที่'] <= end_date)]

    status_map = { (r['ชื่อ-สกุล'], r['วันที่']): r['สถานะ'] for _, r in status_m.iterrows() }

    att_m['สถานะ'] = att_m.apply(determine_status, args=(status_map,), axis=1)
    att_m = att_m.sort_values(['ชื่อ-สกุล', 'วันที่'])

    summary = (att_m.groupby(['ชื่อ-สกุล', 'สถานะ'], dropna=False)
               .size().reset_index(name='จำนวนวัน'))
    pivot = summary.pivot_table(index='ชื่อ-สกุล', columns='สถานะ', values='จำนวนวัน', aggfunc='sum', fill_value=0).reset_index()

    if 'สาย' in pivot.columns:
        pivot = pivot.rename(columns={'สาย': 'จำนวนครั้งมาสาย'})
    else:
        pivot['จำนวนครั้งมาสาย'] = 0

    return att_m, pivot

# ====================================================
# 🎯 UI Constants & Main App
# ====================================================
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

    col_chart1, col_chart2 = st.columns(2)
    with col_chart1:
        st.markdown("##### **การลาแยกตามกลุ่มงาน**")
        if not df_leave.empty and 'กลุ่มงาน' in df_leave.columns and 'จำนวนวันลา' in df_leave.columns:
            leave_by_group = df_leave.groupby('กลุ่มงาน')['จำนวนวันลา'].sum().sort_values(ascending=False).reset_index()
            chart_group_leave = alt.Chart(leave_by_group).mark_bar().encode(
                x=alt.X('จำนวนวันลา:Q', title='รวมจำนวนวันลา'),
                y=alt.Y('กลุ่มงาน:N', sort='-x', title='กลุ่มงาน'),
                tooltip=['กลุ่มงาน', 'จำนวนวันลา']
            ).properties(height=300)
            st.altair_chart(chart_group_leave, use_container_width=True)
        else:
            st.info("ไม่มีข้อมูลการลาเพียงพอที่จะแสดงผล")

    with col_chart2:
        st.markdown("##### **ผู้เดินทางราชการบ่อยที่สุด (Top 5)**")
        if not df_travel.empty and 'ชื่อ-สกุล' in df_travel.columns:
            top_travelers = df_travel['ชื่อ-สกุล'].value_counts().nlargest(5).reset_index()
            top_travelers.columns = ['ชื่อ-สกุล', 'จำนวนครั้ง']
            chart_top_travel = alt.Chart(top_travelers).mark_bar(color='#ff8c00').encode(
                x=alt.X('จำนวนครั้ง:Q', title='จำนวนครั้งไปราชการ'),
                y=alt.Y('ชื่อ-สกุล:N', sort='-x', title='ชื่อ-สกุล'),
                tooltip=['ชื่อ-สกุล', 'จำนวนครั้ง']
            ).properties(height=300)
            st.altair_chart(chart_top_travel, use_container_width=True)
        else:
            st.info("ไม่มีข้อมูลการเดินทางราชการ")

    st.markdown("##### **แนวโน้มการลา (รายเดือน)**")
    if not daily_status.empty:
        daily_leave_only = daily_status[daily_status['สถานะ'].str.contains("ลา", na=False)].copy()
        if not daily_leave_only.empty:
            daily_leave_only['เดือน'] = pd.to_datetime(daily_leave_only['วันที่']).dt.strftime('%Y-%m')
            leave_trend = daily_leave_only.groupby('เดือน').size().reset_index(name='จำนวนวันลา')
            chart_trend = alt.Chart(leave_trend).mark_line(point=True, strokeWidth=3).encode(
                x=alt.X('เดือน:T', title='เดือน'),
                y=alt.Y('จำนวนวันลา:Q', title='รวมจำนวนวันลา (ทุกประเภท)'),
                tooltip=['เดือน', 'จำนวนวันลา']
            ).properties(title='จำนวนวันลาทั้งหมดในแต่ละเดือน')
            st.altair_chart(chart_trend, use_container_width=True)
        else:
            st.info("ไม่มีข้อมูลการลาสำหรับแสดงแนวโน้ม")
    st.markdown("---")
    
    # --- 4. Heatmap Calendar ---
    st.markdown("#### **ปฏิทิน Heatmap (สรุปการลาและไปราชการ)**")
    today = dt.date.today()
    colh1, colh2 = st.columns([1,2])
    sel_month_h = colh1.selectbox("เลือกเดือน (สำหรับ Heatmap)", range(1, 13), index=today.month-1, format_func=lambda m: f"{m:02d}", key="hm_month")
    sel_year_h = colh1.number_input("เลือกปี (ค.ศ.)", value=today.year, min_value=2020, max_value=2050, key="hm_year")

    start_date_h = dt.date(sel_year_h, sel_month_h, 1)
    end_date_h = (start_date_h + dt.timedelta(days=32)).replace(day=1) - dt.timedelta(days=1)
    
    monthly_status = daily_status[(daily_status['วันที่'] >= start_date_h) & (daily_status['วันที่'] <= end_date_h)]
    
    if not monthly_status.empty:
        heatmap_data = monthly_status.groupby('วันที่').size().reset_index(name='จำนวนคน')
        
        # --- ‼️ CODE HAS BEEN FIXED HERE ‼️ ---
        heatmap = alt.Chart(heatmap_data).mark_rect().encode(
            x=alt.X('date(วันที่):O', title='วันที่'),
            y=alt.Y('day(วันที่):O', title='วันในสัปดาห์', sort='descending'),
            color=alt.Color('จำนวนคน:Q', scale=alt.Scale(scheme='lighttealblue'), title='จำนวนคน'),
            # FIX: Changed 'utchmonthdate' to a standard temporal format
            tooltip=[
                alt.Tooltip('วันที่:T', title='วันที่', format='%A, %B %d, %Y'),
                alt.Tooltip('จำนวนคน:Q', title='จำนวนคน (ลา/ราชการ)')
            ]
        ).properties(
            title=f"ภาพรวมกำลังคน เดือน {sel_month_h}/{sel_year_h}"
        )
        
        text = heatmap.mark_text(baseline='middle').encode(
            text='date(วันที่):O',
            color=alt.condition(
                alt.datum.จำนวนคน > 5,
                alt.value('white'),
                alt.value('black')
            )
        )
        st.altair_chart(heatmap + text, use_container_width=True)
    else:
        st.info(f"ไม่พบข้อมูลการลาหรือไปราชการในเดือน {sel_month_h}/{sel_year_h}")

# The rest of the app code remains the same...

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
            st.rerun()

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
            st.rerun()

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
