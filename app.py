import streamlit as st
import pandas as pd
import datetime as dt
import altair as alt
import matplotlib.pyplot as plt
from fpdf import FPDF
import tempfile
from io import BytesIO

# ===== ไฟล์เก็บข้อมูล =====
FILE_SCAN = "scan_report.xlsx"
FILE_REPORT = "leave_report.xlsx"
ADMIN_PASSWORD = "admin2568"  # 🔐 รหัสผ่านผู้ดูแลระบบ

# ===== โหลดไฟล์ =====
try:
    df_scan = pd.read_excel(FILE_SCAN)
except:
    df_scan = pd.DataFrame()

try:
    df_report = pd.read_excel(FILE_REPORT)
except:
    df_report = pd.DataFrame()

# ===== รายชื่อกลุ่มงาน =====
staff_groups = [
    "กลุ่มโรคติดต่อ",
    "กลุ่มระบาดวิทยาและตอบโต้ภาวะฉุกเฉินทางสาธารณสุข",
    "กลุ่มพัฒนาองค์กร",
    "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.1 จ.ชัยภูมิ",
    "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.2 จ.บุรีรัมย์",
    "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.3 จ.สุรินทร์",
    "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.4 ปากช่อง",
    "ด่านควบคุมโรคติดต่อระหว่างประเทศพรมแดนช่องจอม จังหวัดสุรินทร์",
    "กลุ่มโรคไม่ติดต่อ",
    "งานควบคุมโรคเขตเมือง",
    "งานกฎหมาย",
    "กลุ่มโรคติดต่อเรื้อรัง",
    "กลุ่มห้องปฏิบัติการทางการแพทย์ด้านควบคุมโรค",
    "กลุ่มสื่อสารความเสี่ยงโรคและภัยสุขภาพ",
    "กลุ่มโรคจากการประกอบอาชีพและสิ่งแวดล้อม",
    "ศูนย์บริการเวชศาสตร์ป้องกัน",
    "กลุ่มบริหารทั่วไป",
    "ศูนย์ฝึกอบรมนักระบาดวิทยาภาคสนาม",
    "กลุ่มพัฒนานวัตกรรมและวิจัย"
]

# ===== เมนูหลัก =====
st.title("📋 โปรแกรมบันทึกข้อมูล (สคร.9)")
main_menu = st.sidebar.radio(
    "เลือกหน้าเมนู",
    ["📊 Dashboard รวม", "🧭 การไปราชการ", "🕒 การลา", "🛠️ ผู้ดูแลระบบ (Admin)"]
)

# ===== Dashboard รวม =====
if main_menu == "📊 Dashboard รวม":
    st.header("📈 Dashboard สรุปข้อมูลทั้งหมด")

    total_travel = len(df_scan)
    total_leave = len(df_report)
    total_travel_days = df_scan["จำนวนวัน"].sum() if "จำนวนวัน" in df_scan.columns else 0
    total_leave_days = df_report["จำนวนวันลา"].sum() if "จำนวนวันลา" in df_report.columns else 0

    col1, col2 = st.columns(2)
    col1.metric("จำนวนผู้ไปราชการ", f"{total_travel} คน")
    col1.metric("จำนวนวันไปราชการรวม", f"{total_travel_days} วัน")
    col2.metric("จำนวนผู้ลา", f"{total_leave} คน")
    col2.metric("จำนวนวันลารวม", f"{total_leave_days} วัน")

    # --- สรุปข้อมูลการลาเพิ่มเติม ---
    if not df_report.empty:
        st.subheader("🧾 Key Metrics การลา")

        unique_persons = df_report["ชื่อ-สกุล"].nunique()
        avg_leave = round(df_report["จำนวนวันลา"].mean(), 2) if "จำนวนวันลา" in df_report.columns else 0

        c1, c2 = st.columns(2)
        c1.metric("จำนวนบุคลากรที่ลา", f"{unique_persons} คน")
        c2.metric("เฉลี่ยวันลาต่อคน", f"{avg_leave} วัน")

        # --- กราฟ 1: แสดงจำนวนวันลาตามประเภท ---
        st.markdown("### 📊 จำนวนวันลาตามประเภท")
        leave_by_type = df_report.groupby("ประเภทการลา")["จำนวนวันลา"].sum().reset_index()
        if not leave_by_type.empty:
            chart1 = alt.Chart(leave_by_type).mark_bar().encode(
                x=alt.X("ประเภทการลา:N", title="ประเภทการลา"),
                y=alt.Y("จำนวนวันลา:Q", title="จำนวนวันลา"),
                color="ประเภทการลา"
            )
            st.altair_chart(chart1, use_container_width=True)

        # --- กราฟ 2: แสดงจำนวนวันลาตามกลุ่มงาน ---
        st.markdown("### 🏢 จำนวนวันลาตามกลุ่มงาน")
        leave_by_group = df_report.groupby("กลุ่มงาน")["จำนวนวันลา"].sum().reset_index()
        if not leave_by_group.empty:
            chart2 = alt.Chart(leave_by_group).mark_bar().encode(
                x=alt.X("กลุ่มงาน:N", sort="-y", title="กลุ่มงาน"),
                y=alt.Y("จำนวนวันลา:Q", title="จำนวนวันลา"),
                color="กลุ่มงาน"
            )
            st.altair_chart(chart2, use_container_width=True)

        st.markdown("### 📋 ตารางข้อมูลการลา (ดิบ)")
        st.dataframe(df_report.astype(str), use_container_width=True)

# ===== ฟอร์มบันทึกการไปราชการ =====
elif main_menu == "🧭 การไปราชการ":
    st.header("🧾 แบบฟอร์มบันทึกการไปราชการ")
    with st.form("scan_form"):
        data = {}
        data["ชื่อ-สกุล"] = st.text_input("👤 ชื่อ-สกุล")
        data["กลุ่มงาน"] = st.selectbox("🏢 กลุ่มงาน", staff_groups)
        data["ปี พ.ศ."] = st.number_input("📅 ปี พ.ศ.", min_value=2560, max_value=2600, value=2568)
        data["กิจกรรม"] = st.text_input("🏢 กิจกรรม (พิมพ์ชื่อกิจกรรม)")
        data["สถานที่"] = st.text_input("📍 สถานที่")
        data["ผู้จัด"] = st.text_input("👥 ผู้จัด")
        data["วันที่เริ่ม"] = st.date_input("📅 วันที่เริ่ม", dt.date.today())
        data["วันที่สิ้นสุด"] = st.date_input("📅 วันที่สิ้นสุด", dt.date.today())
        data["จำนวนวัน"] = (data["วันที่สิ้นสุด"] - data["วันที่เริ่ม"]).days + 1
        st.write(f"📆 รวม {data['จำนวนวัน']} วัน")

        # ผู้ร่วมเดินทาง
        st.markdown("### 👨‍👩‍👧‍👦 รายชื่อผู้ร่วมเดินทาง (หากไปเป็นหมู่คณะ)")
        num_people = st.number_input("จำนวนผู้ร่วมเดินทาง", min_value=0, max_value=20, step=1)
        companions = []
        for i in range(int(num_people)):
            name = st.text_input(f"ชื่อ-สกุลผู้ร่วมเดินทางคนที่ {i+1}")
            group = st.selectbox(f"กลุ่มงานของผู้ร่วมเดินทางคนที่ {i+1}", staff_groups, key=f"group_{i}")
            if name:
                companions.append({"ชื่อ-สกุล": name, "กลุ่มงาน": group})
        data["ผู้ร่วมเดินทาง"] = ", ".join([p["ชื่อ-สกุล"] for p in companions])
        data["หมายเหตุ"] = st.text_area("📝 หมายเหตุ")

        submitted = st.form_submit_button("✅ บันทึกข้อมูล")

    if submitted:
        if not data["ชื่อ-สกุล"]:
            st.error("⚠️ กรุณากรอกชื่อ-นามสกุลของผู้ไปราชการ")
        elif not data["กิจกรรม"]:
            st.error("⚠️ กรุณากรอกชื่อกิจกรรม")
        elif int(num_people) > 0 and len(companions) < int(num_people):
            st.error("⚠️ กรุณากรอกชื่อ-นามสกุลผู้ร่วมเดินทางให้ครบ")
        else:
            df_scan = df_scan.dropna(how='all')
            df_scan = pd.concat([df_scan, pd.DataFrame([data])], ignore_index=True)
            df_scan.to_excel(FILE_SCAN, index=False)
            st.success("✅ บันทึกข้อมูลการไปราชการเรียบร้อย")

    if not df_scan.empty:
        st.markdown("## 📊 สรุปข้อมูลการไปราชการ")
        st.write(f"📋 ทั้งหมด {len(df_scan)} รายการ รวม {df_scan['จำนวนวัน'].sum()} วัน")
        st.dataframe(df_scan.astype(str), use_container_width=True)

# ===== ฟอร์มบันทึกการลา =====
elif main_menu == "🕒 การลา":
    st.header("📝 แบบฟอร์มบันทึกการลา")
    with st.form("leave_form"):
        data = {}
        data["ชื่อ-สกุล"] = st.text_input("👤 ชื่อ-สกุล")
        data["กลุ่มงาน"] = st.selectbox("🏢 กลุ่มงาน", staff_groups)
        data["ประเภทการลา"] = st.selectbox("📌 ประเภทการลา", ["ลาป่วย", "ลากิจ", "ลาพักผ่อน", "อื่นๆ"])
        data["วันที่เริ่ม"] = st.date_input("📅 วันที่เริ่ม", dt.date.today())
        data["วันที่สิ้นสุด"] = st.date_input("📅 วันที่สิ้นสุด", dt.date.today())
        data["จำนวนวันลา"] = (data["วันที่สิ้นสุด"] - data["วันที่เริ่ม"]).days + 1
        st.write(f"🗓️ รวมลา {data['จำนวนวันลา']} วัน")
        data["หมายเหตุ"] = st.text_area("📝 หมายเหตุ")

        submitted = st.form_submit_button("✅ บันทึกข้อมูล")

    if submitted:
        if not data["ชื่อ-สกุล"]:
            st.error("⚠️ กรุณากรอกชื่อ-นามสกุลของผู้ลา")
        else:
            df_report = df_report.dropna(how='all')
            df_report = pd.concat([df_report, pd.DataFrame([data])], ignore_index=True)
            df_report.to_excel(FILE_REPORT, index=False)
            st.success("✅ บันทึกข้อมูลการลาเรียบร้อย")

    if not df_report.empty:
        st.markdown("## 📊 สรุปข้อมูลการลา")
        st.write(f"📋 ทั้งหมด {len(df_report)} รายการ รวมลา {df_report['จำนวนวันลา'].sum()} วัน")
        st.dataframe(df_report.astype(str), use_container_width=True)

# ===== โหมดผู้ดูแลระบบ =====
elif main_menu == "🛠️ ผู้ดูแลระบบ (Admin)":
    st.header("🧑‍💼 โหมดผู้ดูแลระบบ")
    password = st.text_input("🔐 กรุณากรอกรหัสผ่าน", type="password")

    if password == ADMIN_PASSWORD:
        st.success("✅ เข้าสู่ระบบสำเร็จ")
        tab1, tab2 = st.tabs(["📄 ข้อมูลการไปราชการ", "🕒 ข้อมูลการลา"])

        with tab1:
            st.subheader("📄 ข้อมูลการไปราชการทั้งหมด")
            st.dataframe(df_scan.astype(str), use_container_width=True)
            if st.button("🗑️ ล้างข้อมูลไปราชการทั้งหมด"):
                df_scan = pd.DataFrame()
                df_scan.to_excel(FILE_SCAN, index=False)
                st.warning("ลบข้อมูลการไปราชการทั้งหมดแล้ว ✅")

        with tab2:
            st.subheader("🕒 ข้อมูลการลาทั้งหมด")
            st.dataframe(df_report.astype(str), use_container_width=True)
            if st.button("🗑️ ล้างข้อมูลการลาทั้งหมด"):
                df_report = pd.DataFrame()
                df_report.to_excel(FILE_REPORT, index=False)
                st.warning("ลบข้อมูลการลาทั้งหมดแล้ว ✅")

    elif password:
        st.error("❌ รหัสผ่านไม่ถูกต้อง กรุณาลองใหม่")
    else:
        st.info("ℹ️ กรุณาใส่รหัสผ่านเพื่อเข้าสู่โหมดแอดมิน")
