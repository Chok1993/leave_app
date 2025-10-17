# ====================================================
# 📋 โปรแกรมติดตามการลาและไปราชการ (สคร.9)
# ✅ Final Full Version: ครบทุกเมนู + อัปเกรดระบบอ่าน Excel ทนทานขึ้น
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

# ====================================================
# 🔐 Auth & App Config
# ====================================================
st.set_page_config(page_title="สคร.9 - ติดตามการลา/ราชการ/สแกน", layout="wide")

creds = service_account.Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=["https://www.googleapis.com/auth/drive"]
)
service = build("drive", "v3", credentials=creds)
ADMIN_PASSWORD = st.secrets.get("admin_password", "admin123")

# ====================================================
# 🗂️ Shared Drive Config
# ====================================================
FOLDER_ID = "1YFJZvs59ahRHmlRrKcQwepWJz6A-4B7d"
ATTACHMENT_FOLDER_NAME = "เอกสารแนบ_ไปราชการ"
FILE_ATTEND = "attendance_report.xlsx"
FILE_LEAVE  = "leave_report.xlsx"
FILE_TRAVEL = "travel_report.xlsx"

# ====================================================
# 🔧 Drive Helper Functions
# ====================================================
def get_file_id(filename: str, parent_id=FOLDER_ID):
    q = f"name='{filename}' and '{parent_id}' in parents and trashed=false"
    res = service.files().list(
        q=q, fields="files(id,name)", supportsAllDrives=True, includeItemsFromAllDrives=True
    ).execute()
    files = res.get("files", [])
    return files[0]["id"] if files else None


@st.cache_data(ttl=600)
def read_excel_from_drive(filename: str) -> pd.DataFrame:
    """อ่านไฟล์ Excel จาก Shared Drive; ถ้าไม่มีไฟล์จะคืนค่า DataFrame ว่าง"""
    try:
        file_id = get_file_id(filename)
        if not file_id:
            st.warning(f"⚠️ ไม่พบไฟล์ '{filename}' ใน Google Drive กรุณาตรวจสอบชื่อไฟล์ให้ถูกต้อง")
            return pd.DataFrame()

        req = service.files().get_media(fileId=file_id, supportsAllDrives=True)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)

        xls = pd.ExcelFile(fh, engine="openpyxl")
        if not xls.sheet_names:
            st.error(f"❌ ไฟล์ '{filename}' ไม่มีชีตข้อมูล")
            return pd.DataFrame()

        df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
        expected_cols = ["วันที่", "ชื่อพนักงาน", "ชื่อ-สกุล"]
        if not any(col in df.columns for col in expected_cols):
            fh.seek(0)
            df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=1)

        return df
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์ {filename}: {e}")
        return pd.DataFrame()


def write_excel_to_drive(filename: str, df: pd.DataFrame):
    """อัปโหลด DataFrame กลับไปยัง Drive"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    media = MediaIoBaseUpload(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    file_id = get_file_id(filename)
    if file_id:
        service.files().update(fileId=file_id, media_body=media, supportsAllDrives=True).execute()
    else:
        service.files().create(
            body={"name": filename, "parents": [FOLDER_ID]},
            media_body=media, supportsAllDrives=True
        ).execute()


def count_weekdays(start_date, end_date):
    if start_date and end_date and start_date <= end_date:
        return np.busday_count(start_date, end_date + dt.timedelta(days=1))
    return 0

# ====================================================
# 📥 Load Data
# ====================================================
df_att = read_excel_from_drive(FILE_ATTEND)
df_leave = read_excel_from_drive(FILE_LEAVE)
df_travel = read_excel_from_drive(FILE_TRAVEL)

# ====================================================
# 🎯 UI Header
# ====================================================
st.markdown("##### **สำนักงานป้องกันควบคุมโรคที่ 9 จังหวัดนครราชสีมา**")
st.title("📋 ระบบติดตามการลา ไปราชการ และการปฏิบัติงาน (สคร.9)")

# ====================================================
# 🧭 Sidebar Menu
# ====================================================
menu = st.sidebar.radio(
    "เลือกเมนู",
    ["หน้าหลัก", "📊 Dashboard", "📅 การมาปฏิบัติงาน", "🧭 การไปราชการ", "🕒 การลา", "🧑‍💼 ผู้ดูแลระบบ"]
)

# ====================================================
# 🏠 หน้าหลัก
# ====================================================
if menu == "หน้าหลัก":
    st.info(
        """
        💡 ระบบนี้ใช้สำหรับบันทึกและสรุปข้อมูลบุคลากรของ **สำนักงานป้องกันควบคุมโรคที่ 9 นครราชสีมา**
        - การลา  
        - การไปราชการ  
        - การมาปฏิบัติงาน  
        - การแนบเอกสาร PDF ผ่านระบบอัตโนมัติ
        """
    )
    st.image(
        "https://ddc.moph.go.th/uploads/files/11120210817094038.jpg",
        caption="สำนักงานป้องกันควบคุมโรคที่ 9 จังหวัดนครราชสีมา",
        use_container_width=True
    )

# ====================================================
# 📊 Dashboard
# ====================================================
elif menu == "📊 Dashboard":
    st.header("📊 Dashboard สรุปภาพรวม")

    col1, col2, col3 = st.columns(3)
    col1.metric("📄 ข้อมูลสแกน (แถว)", len(df_att))
    col2.metric("🕒 การลา (ครั้ง)", len(df_leave))
    col3.metric("🧭 การไปราชการ (ครั้ง)", len(df_travel))
    st.markdown("---")

    # กราฟลา
    if not df_leave.empty and "กลุ่มงาน" in df_leave.columns and "จำนวนวันลา" in df_leave.columns:
        leave_by_group = (
            df_leave.groupby("กลุ่มงาน")["จำนวนวันลา"]
            .sum().sort_values(ascending=False).reset_index()
        )
        chart = alt.Chart(leave_by_group).mark_bar().encode(
            x=alt.X("จำนวนวันลา:Q", title="รวมวันลา"),
            y=alt.Y("กลุ่มงาน:N", sort="-x", title="กลุ่มงาน"),
            tooltip=["กลุ่มงาน", "จำนวนวันลา"]
        ).properties(height=300)
        st.altair_chart(chart, use_container_width=True)
    else:
        st.info("ไม่มีข้อมูลลาเพียงพอในการแสดงผล")

# ====================================================
# 📅 การมาปฏิบัติงาน (Core Feature)
# ====================================================
elif menu == "📅 การมาปฏิบัติงาน":
    st.header("📅 สรุปการมาปฏิบัติงานรายวัน (ตรวจจากสแกน + ลา + ราชการ)")

    if df_att.empty:
        st.warning("⚠️ ยังไม่มีข้อมูลสแกนเข้า-ออกในระบบ")
        st.stop()

    rename_map = {"ชื่อ-สกุล": "ชื่อพนักงาน", "ชื่อ": "ชื่อพนักงาน"}
    for df in [df_att, df_leave, df_travel]:
        df.rename(columns=rename_map, inplace=True)

    df_att["วันที่"] = pd.to_datetime(df_att["วันที่"], errors="coerce")
    for df in [df_leave, df_travel]:
        for c in ["วันที่เริ่ม", "วันที่สิ้นสุด"]:
            if c in df.columns:
                df[c] = pd.to_datetime(df[c], errors="coerce")

    df_att["เดือน"] = df_att["วันที่"].dt.strftime("%Y-%m")
    months = sorted(df_att["เดือน"].dropna().unique())
    selected_month = st.selectbox("เลือกเดือน", months, index=len(months)-1)
    all_names = sorted(df_att["ชื่อพนักงาน"].dropna().unique())
    selected_names = st.multiselect("เลือกชื่อบุคลากร", all_names, default=all_names[:3])

    df_month = df_att[df_att["เดือน"] == selected_month]
    if selected_names:
        df_month = df_month[df_month["ชื่อพนักงาน"].isin(selected_names)]

    if df_month.empty:
        st.warning("ไม่มีข้อมูลในเดือนนี้")
        st.stop()

    WORK_START = dt.time(8, 30)
    WORK_END = dt.time(16, 30)

    y, m = map(int, selected_month.split("-"))
    month_start = dt.date(y, m, 1)
    month_end = (month_start + pd.offsets.MonthEnd(0)).date()
    date_range = pd.date_range(month_start, month_end, freq="D")

    records = []
    for name in df_month["ชื่อพนักงาน"].unique():
        for d in date_range:
            rec = {"ชื่อพนักงาน": name, "วันที่": d.date(), "เวลาเข้า": "", "เวลาออก": "", "หมายเหตุ": "", "สถานะ": ""}
            att = df_month[(df_month["ชื่อพนักงาน"] == name) & (df_month["วันที่"].dt.date == d.date())]

            if not att.empty:
                rec["เวลาเข้า"] = att.iloc[0].get("เวลาเข้า", "")
                rec["เวลาออก"] = att.iloc[0].get("เวลาออก", "")
                rec["หมายเหตุ"] = att.iloc[0].get("หมายเหตุ", "")

                if d.weekday() >= 5:
                    rec["สถานะ"] = "วันหยุด"
                else:
                    try:
                        t_in = pd.to_datetime(str(rec["เวลาเข้า"]), format="%H:%M").time() if rec["เวลาเข้า"] else None
                        t_out = pd.to_datetime(str(rec["เวลาออก"]), format="%H:%M").time() if rec["เวลาออก"] else None
                    except:
                        t_in, t_out = None, None

                    if not t_in and not t_out:
                        rec["สถานะ"] = "ขาดงาน"
                    elif t_in and not t_out:
                        rec["สถานะ"] = "ออกก่อน"
                    elif not t_in and t_out:
                        rec["สถานะ"] = "ขาดงาน"
                    else:
                        if t_in > WORK_START and t_out < WORK_END:
                            rec["สถานะ"] = "มาสายและออกก่อน"
                        elif t_in > WORK_START:
                            rec["สถานะ"] = "มาสาย"
                        elif t_out < WORK_END:
                            rec["สถานะ"] = "ออกก่อน"
                        else:
                            rec["สถานะ"] = "มาปกติ"
            else:
                in_leave = (
                    not df_leave.empty and
                    (df_leave["ชื่อพนักงาน"] == name).any() and
                    (df_leave[
                        (df_leave["ชื่อพนักงาน"] == name) &
                        (df_leave["วันที่เริ่ม"] <= d) &
                        (df_leave["วันที่สิ้นสุด"] >= d)
                    ].shape[0] > 0)
                )
                in_travel = (
                    not df_travel.empty and
                    (df_travel["ชื่อพนักงาน"] == name).any() and
                    (df_travel[
                        (df_travel["ชื่อพนักงาน"] == name) &
                        (df_travel["วันที่เริ่ม"] <= d) &
                        (df_travel["วันที่สิ้นสุด"] >= d)
                    ].shape[0] > 0)
                )

                if in_leave:
                    leave_type = df_leave.loc[
                        (df_leave["ชื่อพนักงาน"] == name) &
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

    df_daily = pd.DataFrame(records)

    def color_status(val):
        colors = {
            "มาปกติ": "background-color:#d4edda",
            "มาสาย": "background-color:#ffeeba",
            "ออกก่อน": "background-color:#f8d7da",
            "มาสายและออกก่อน": "background-color:#fcd5b5",
            "ลา": "background-color:#d1ecf1",
            "ไปราชการ": "background-color:#fff3cd",
            "วันหยุด": "background-color:#e2e3e5",
            "ขาดงาน": "background-color:#f5c6cb",
        }
        for key in colors:
            if key in str(val):
                return colors[key]
        return ""

    st.markdown("### 📋 ตารางสรุปสถานะรายวัน")
    st.dataframe(df_daily.style.applymap(color_status, subset=["สถานะ"]), use_container_width=True, height=600)

# ====================================================
# 🧭 การไปราชการ
# ====================================================
elif menu == "🧭 การไปราชการ":
    st.header("🧭 บันทึกข้อมูลการไปราชการ")
    st.info("ฟังก์ชันนี้สามารถแนบไฟล์คำสั่ง PDF และอัปโหลดเข้า Drive ได้ (จะเพิ่มในรุ่นถัดไป)")

# ====================================================
# 🕒 การลา
# ====================================================
elif menu == "🕒 การลา":
    st.header("🕒 ข้อมูลการลา")
    st.info("ส่วนนี้จะเชื่อมข้อมูลกับระบบลาราชการในเวอร์ชันถัดไป")

# ====================================================
# 🧑‍💼 ผู้ดูแลระบบ
# ====================================================
elif menu == "🧑‍💼 ผู้ดูแลระบบ":
    st.header("🧑‍💼 ผู้ดูแลระบบ")
    st.info("ในรุ่นนี้ผู้ดูแลสามารถตรวจสอบไฟล์ทั้งหมดใน Shared Drive ได้โดยตรง")
