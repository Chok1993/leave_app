# ====================================================
# 📋 โปรแกรมติดตามการลาและไปราชการ (สคร.9)
# ✅ Full Production Version: แก้ Dashboard + เพิ่มเมนูครบ + ป้องกัน Error
# ====================================================

import io
import altair as alt
import datetime as dt
import pandas as pd
import numpy as np
import streamlit as st
import re
import requests

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
ADMIN_PASSWORD = st.secrets.get("admin_password", "admin123")

# ===========================
# 🗂️ Shared Drive Config
# ===========================
FOLDER_ID = "1YFJZvs59ahRHmlRrKcQwepWJz6A-4B7d"
ATTACHMENT_FOLDER_NAME = "เอกสารแนบ_ไปราชการ"
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
    if df.empty:
        return
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    media = MediaIoBaseUpload(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    file_id = get_file_id(filename)
    if file_id:
        service.files().update(fileId=file_id, media_body=media, supportsAllDrives=True).execute()
    else:
        service.files().create(body={"name": filename, "parents": [FOLDER_ID]},
                               media_body=media, fields="id", supportsAllDrives=True).execute()

def backup_excel(original_filename: str, df: pd.DataFrame):
    """สำรองข้อมูลทุกครั้งก่อนบันทึก"""
    if df.empty:
        return
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
        meta = {'name': folder_name, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_folder_id]}
        folder = service.files().create(body=meta, fields='id', supportsAllDrives=True).execute()
        return folder.get('id')

def upload_pdf_to_drive(file_object, filename, folder_id):
    """อัปโหลด PDF คืนค่า ViewLink"""
    if file_object is None:
        return "-"
    file_object.seek(0)
    media = MediaIoBaseUpload(file_object, mimetype='application/pdf', resumable=True)
    meta = {'name': filename, 'parents': [folder_id]}
    file = service.files().create(body=meta, media_body=media, fields='id, webViewLink', supportsAllDrives=True).execute()
    permission = {'type': 'domain', 'role': 'reader', 'domain': 'ddc.mail.go.th'}
    service.permissions().create(fileId=file.get('id'), body=permission, supportsAllDrives=True).execute()
    return file.get('webViewLink')

def count_weekdays(start_date, end_date):
    """นับเฉพาะวันทำการ"""
    if start_date and end_date and start_date <= end_date:
        return np.busday_count(start_date, end_date + dt.timedelta(days=1))
    return 0

# ===========================
# 📥 Load Data
# ===========================
df_att = read_excel_from_drive(FILE_ATTEND)
df_leave = read_excel_from_drive(FILE_LEAVE)
df_travel = read_excel_from_drive(FILE_TRAVEL)

# ===========================
# 🎯 UI Constants & Main App
# ===========================
st.title("📋 ระบบติดตามการลา ไปราชการ และการปฏิบัติงาน (สคร.9)")
st.caption("สำนักงานป้องกันควบคุมโรคที่ 9 จังหวัดนครราชสีมา")

if 'submitted' not in st.session_state:
    st.session_state.submitted = False

def callback_submit():
    st.session_state.submitted = True

staff_groups = sorted([
    "กลุ่มโรคติดต่อ", "กลุ่มระบาดวิทยาและตอบโต้ภาวะฉุกเฉินทางสาธารณสุข", "กลุ่มพัฒนาองค์กร",
    "กลุ่มบริหารทั่วไป", "กลุ่มโรคไม่ติดต่อ", "กลุ่มห้องปฏิบัติการทางการแพทย์", "กลุ่มพัฒนานวัตกรรมและวิจัย",
    "กลุ่มโรคติดต่อเรื้อรัง", "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.1 จ.ชัยภูมิ", "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.2 จ.บุรีรัมย์",
    "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.3 จ.สุรินทร์", "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.4 ปากช่อง",
    "ด่านควบคุมโรคช่องจอม จ.สุรินทร์", "ศูนย์บริการเวชศาสตร์ป้องกัน", "กลุ่มสื่อสารความเสี่ยง",
    "กลุ่มโรคจากการประกอบอาชีพและสิ่งแวดล้อม"
])
leave_types = ["ลาป่วย", "ลากิจ", "ลาพักผ่อน", "อื่นๆ"]

all_names = sorted(list(set(pd.concat([
    df_leave.get('ชื่อ-สกุล', pd.Series(dtype=str)),
    df_travel.get('ชื่อ-สกุล', pd.Series(dtype=str)),
    df_att.get('ชื่อ-สกุล', pd.Series(dtype=str))
]).dropna())))

menu = st.sidebar.radio("เลือกเมนู", [
    "หน้าหลัก", "📊 Dashboard", "📅 การมาปฏิบัติงาน", "🧭 การไปราชการ", "🕒 การลา", "🧑‍💼 ผู้ดูแลระบบ"
])

# ----------------------------
# 🏠 หน้าหลัก
# ----------------------------
if menu == "หน้าหลัก":
    st.info("💡 ระบบนี้ใช้สำหรับบันทึกและสรุปข้อมูลบุคลากรใน สคร.9\n"
            "ได้แก่ การลา การไปราชการ และการมาปฏิบัติงาน พร้อมแนบไฟล์เอกสาร PDF ได้โดยตรง")
    st.image("https://ddc.moph.go.th/uploads/files/11120210817094038.jpg",
             caption="สำนักงานป้องกันควบคุมโรคที่ 9 นครราชสีมา", use_container_width=True)

# ----------------------------
# 📊 Dashboard
# ----------------------------
elif menu == "📊 Dashboard":
    st.header("📊 Dashboard ภาพรวมและข้อมูลเชิงลึก")
    col1, col2, col3 = st.columns(3)
    col1.metric("เดินทางราชการ (ครั้ง)", len(df_travel))
    col2.metric("การลา (ครั้ง)", len(df_leave))
    col3.metric("ข้อมูลสแกน (แถว)", len(df_att))
    st.markdown("---")

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("📅 การลาแยกตามกลุ่มงาน")
        if not df_leave.empty and "กลุ่มงาน" in df_leave and "จำนวนวันลา" in df_leave:
            leave_by_group = df_leave.groupby("กลุ่มงาน")["จำนวนวันลา"].sum().reset_index().sort_values("จำนวนวันลา", ascending=False)
            chart = alt.Chart(leave_by_group).mark_bar(color="#4C9A2A").encode(
                x="จำนวนวันลา:Q", y=alt.Y("กลุ่มงาน:N", sort="-x"), tooltip=["กลุ่มงาน", "จำนวนวันลา"]
            )
            st.altair_chart(chart, use_container_width=True)
        else:
            st.info("ไม่มีข้อมูลการลาเพียงพอ")

    with c2:
        st.subheader("🧭 ผู้เดินทางราชการบ่อยที่สุด (Top 5)")
        if not df_travel.empty and "ชื่อ-สกุล" in df_travel:
            top5 = df_travel["ชื่อ-สกุล"].value_counts().nlargest(5).reset_index()
            top5.columns = ["ชื่อ-สกุล", "จำนวนครั้ง"]
            chart = alt.Chart(top5).mark_bar(color="#E67E22").encode(
                x="จำนวนครั้ง:Q", y=alt.Y("ชื่อ-สกุล:N", sort="-x"), tooltip=["ชื่อ-สกุล", "จำนวนครั้ง"]
            )
            st.altair_chart(chart, use_container_width=True)
        else:
            st.info("ไม่มีข้อมูลไปราชการ")

    # ----------------------------
# 📅 การมาปฏิบัติงาน (เวอร์ชันสมบูรณ์ ตรวจเวลาจริง + ไขว้ 3 ฐาน)
# ----------------------------
elif menu == "📅 การมาปฏิบัติงาน":
    st.header("📅 สรุปการมาปฏิบัติงานรายวัน (ตรวจจากสแกน + ลา + ราชการ)")

    if df_att.empty:
        st.warning("ยังไม่มีข้อมูลสแกนเข้า-ออกในระบบ")
        st.stop()

    # ✅ ปรับชื่อคอลัมน์ให้เป็นชื่อเดียวกัน
    rename_map = {"ชื่อ-สกุล": "ชื่อพนักงาน", "ชื่อ": "ชื่อพนักงาน"}
    for df in [df_att, df_leave, df_travel]:
        df.rename(columns=rename_map, inplace=True)

    if "ชื่อพนักงาน" not in df_att.columns:
        st.error("⚠️ ไม่พบคอลัมน์ชื่อพนักงานในไฟล์สแกน")
        st.stop()

    # ✅ แปลงวันที่ทั้งหมด
    df_att["วันที่"] = pd.to_datetime(df_att["วันที่"], errors="coerce")
    for df in [df_leave, df_travel]:
        for c in ["วันที่เริ่ม", "วันที่สิ้นสุด"]:
            if c in df.columns:
                df[c] = pd.to_datetime(df[c], errors="coerce")

    # ✅ สร้างคอลัมน์เดือน
    df_att["เดือน"] = df_att["วันที่"].dt.strftime("%Y-%m")
    months = sorted(df_att["เดือน"].dropna().unique())
    if not months:
        st.warning("ไม่พบข้อมูลเดือนในไฟล์สแกน")
        st.stop()

    selected_month = st.selectbox("เลือกเดือนที่ต้องการดู", months, index=len(months)-1)
    selected_names = st.multiselect("เลือกชื่อบุคลากร", sorted(df_att["ชื่อพนักงาน"].unique()), max_selections=5)

    df_month = df_att[df_att["เดือน"] == selected_month].copy()
    if selected_names:
        df_month = df_month[df_month["ชื่อพนักงาน"].isin(selected_names)]

    if df_month.empty:
        st.info("ไม่มีข้อมูลสแกนในเดือนที่เลือก")
        st.stop()

    # ✅ กำหนดช่วงเวลาทำงานราชการ
    WORK_START = dt.time(8, 30)
    WORK_END = dt.time(16, 30)

    # ✅ สร้างช่วงวันที่ทั้งเดือนจากเดือนที่เลือก
    y, m = map(int, selected_month.split("-"))
    month_start = dt.date(y, m, 1)
    month_end = (month_start + pd.offsets.MonthEnd(0)).date()
    date_range = pd.date_range(month_start, month_end, freq="D")

    # ✅ ประมวลผลข้อมูลแต่ละวัน
    records = []
    for name in df_month["ชื่อพนักงาน"].unique():
        for d in date_range:
            rec = {"ชื่อพนักงาน": name, "วันที่": d.date(), "เวลาเข้า": "", "เวลาออก": "", "หมายเหตุ": "", "สถานะ": ""}
            att = df_month[(df_month["ชื่อพนักงาน"] == name) & (df_month["วันที่"].dt.date == d.date())]

            if not att.empty:
                rec["เวลาเข้า"] = att.iloc[0].get("เวลาเข้า", "")
                rec["เวลาออก"] = att.iloc[0].get("เวลาออก", "")
                rec["หมายเหตุ"] = att.iloc[0].get("หมายเหตุ", "")

                # 🧮 ตรวจวันหยุดเสาร์–อาทิตย์ก่อน
                if d.weekday() >= 5:
                    rec["สถานะ"] = "วันหยุด"
                else:
                    try:
                        t_in = pd.to_datetime(str(rec["เวลาเข้า"]), format="%H:%M").time() if rec["เวลาเข้า"] else None
                        t_out = pd.to_datetime(str(rec["เวลาออก"]), format="%H:%M").time() if rec["เวลาออก"] else None
                    except Exception:
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
                # ❌ ไม่มีสแกน → ตรวจลา / ราชการ
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

    # ✅ สร้างตารางแสดงผล
    df_daily = pd.DataFrame(records).sort_values(["ชื่อพนักงาน", "วันที่"])

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

    # 📥 ดาวน์โหลดรายวัน
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_daily.to_excel(writer, index=False, sheet_name="รายวัน")
    output.seek(0)
    st.download_button(
        "📥 ดาวน์โหลดรายงานรายวัน (Excel)",
        data=output,
        file_name=f"รายงานการมาปฏิบัติงาน_{selected_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # 📊 สรุปสถิติรวมต่อเดือน
    st.markdown("---")
    st.subheader("📊 สรุปสถิติรวมต่อเดือนต่อคน")

    df_daily["สถานะย่อ"] = df_daily["สถานะ"].apply(lambda x: "ลา" if str(x).startswith("ลา") else x)
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

    preferred = ["มาปกติ", "มาสาย", "ออกก่อน", "มาสายและออกก่อน", "ลา", "ไปราชการ", "ขาดงาน", "วันหยุด", "รวมทั้งหมด"]
    cols = [c for c in preferred if c in summary.columns]
    summary = summary[["ชื่อพนักงาน"] + cols]

    st.dataframe(summary, use_container_width=True)


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
        upload = st.file_uploader("แนบไฟล์ขออนุมัติ (PDF)", type="pdf")

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

# ----------------------------
# 🧑‍💼 ผู้ดูแลระบบ
# ----------------------------
elif menu == "🧑‍💼 ผู้ดูแลระบบ":
    st.header("🔐 แผงผู้ดูแลระบบ")
    pwd = st.text_input("รหัสผ่าน Admin", type="password")
    if pwd == ADMIN_PASSWORD:
        st.success("เข้าสู่โหมดผู้ดูแลระบบเรียบร้อย ✅")
        st.markdown("### 🔧 ตัวเลือกผู้ดูแล")
        colA, colB = st.columns(2)
        with colA:
            if st.button("📤 ดาวน์โหลดข้อมูลทั้งหมด"):
                with st.spinner("กำลังดาวน์โหลด..."):
                    st.download_button("⬇️ ดาวน์โหลดไฟล์ลา", df_leave.to_csv(index=False), file_name="leave_report.csv")
                    st.download_button("⬇️ ดาวน์โหลดไฟล์ไปราชการ", df_travel.to_csv(index=False), file_name="travel_report.csv")
        with colB:
            if st.button("🧹 เคลียร์ Cache ระบบ"):
                st.cache_data.clear()
                st.cache_resource.clear()
                st.success("ล้าง Cache เรียบร้อยแล้ว")
    elif pwd:
        st.error("รหัสผ่านไม่ถูกต้อง ❌")
    else:
        st.info("กรุณาใส่รหัสผ่านเพื่อเข้าถึงฟังก์ชันผู้ดูแลระบบ")









