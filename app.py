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

    # ✅ แปลงวันที่ทั้งหมด
    df_att["วันที่"] = pd.to_datetime(df_att["วันที่"], errors="coerce")
    for df in [df_leave, df_travel]:
        for c in ["วันที่เริ่ม", "วันที่สิ้นสุด"]:
            if c in df.columns:
                df[c] = pd.to_datetime(df[c], errors="coerce")

    # ✅ สร้างคอลัมน์เดือน
    df_att["เดือน"] = df_att["วันที่"].dt.strftime("%Y-%m")
    months = sorted(df_att["เดือน"].dropna().unique())
    selected_month = st.selectbox("เลือกเดือนที่ต้องการดู", months, index=len(months)-1)

    selected_names = st.multiselect("เลือกชื่อบุคลากร (ว่าง=ทุกคน)", all_names)
    df_month = df_att[df_att["เดือน"] == selected_month].copy()

    WORK_START = dt.time(8, 30)
    WORK_END = dt.time(16, 30)
    month_start = pd.to_datetime(selected_month + "-01").date()
    month_end = (month_start + pd.offsets.MonthEnd(0)).date()
    date_range = pd.date_range(month_start, month_end, freq="D")

    records = []
    names_to_process = selected_names if selected_names else all_names

    for name in names_to_process:
        for d in date_range:
            rec = {"ชื่อพนักงาน": name, "วันที่": d.date(), "เวลาเข้า": "", "เวลาออก": "", "หมายเหตุ": "", "สถานะ": ""}

            att = df_month[(df_month[name_col] == name) & (df_month["วันที่"].dt.date == d.date())]
            in_leave = not df_leave.empty and (df_leave[(df_leave["ชื่อ-สกุล"] == name) & (df_leave["วันที่เริ่ม"] <= d) & (df_leave["วันที่สิ้นสุด"] >= d)].shape[0] > 0)
            in_travel = not df_travel.empty and (df_travel[(df_travel["ชื่อ-สกุล"] == name) & (df_travel["วันที่เริ่ม"] <= d) & (df_travel["วันที่สิ้นสุด"] >= d)].shape[0] > 0)

            if in_leave:
                leave_type = df_leave.loc[(df_leave["ชื่อ-สกุล"] == name) & (df_leave["วันที่เริ่ม"] <= d) & (df_leave["วันที่สิ้นสุด"] >= d), "ประเภทการลา"].iloc[0]
                rec["สถานะ"] = f"ลา ({leave_type})"
            elif in_travel:
                rec["สถานะ"] = "ไปราชการ"
            elif not att.empty:
                row = att.iloc[0]
                rec["เวลาเข้า"] = row.get("เวลาเข้า", "")
                rec["เวลาออก"] = row.get("เวลาออก", "")
                rec["หมายเหตุ"] = row.get("หมายเหตุ", "")
                if d.weekday() >= 5:
                    rec["สถานะ"] = "วันหยุด"
                else:
                    try:
                        t_in = pd.to_datetime(str(rec["เวลาเข้า"])).time() if rec["เวลาเข้า"] else None
                        t_out = pd.to_datetime(str(rec["เวลาออก"])).time() if rec["เวลาออก"] else None
                    except Exception:
                        t_in, t_out = None, None
                    if not t_in and not t_out:
                        rec["สถานะ"] = "ขาดงาน"
                    elif t_in > WORK_START and (not t_out or t_out < WORK_END):
                        rec["สถานะ"] = "มาสายและออกก่อน"
                    elif t_in > WORK_START:
                        rec["สถานะ"] = "มาสาย"
                    elif not t_out or t_out < WORK_END:
                        rec["สถานะ"] = "ออกก่อน"
                    else:
                        rec["สถานะ"] = "มาปกติ"
            else:
                rec["สถานะ"] = "วันหยุด" if d.weekday() >= 5 else "ขาดงาน"
            records.append(rec)

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







