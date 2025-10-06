import streamlit as st
import pandas as pd
import datetime as dt
import altair as alt
import tempfile
from fpdf import FPDF
import matplotlib.pyplot as plt
from io import BytesIO

# ===== ไฟล์เก็บข้อมูล =====
FILE_SCAN = "scan_report.xlsx"
FILE_REPORT = "leave_report.xlsx"

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
    "กลุ่มโรคติดต่อ", "กลุ่มระบาดวิทยาและตอบโต้ภาวะฉุกเฉินทางสาธารณสุข", "กลุ่มพัฒนาองค์กร",
    "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.1 จ.ชัยภูมิ", "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.2 จ.บุรีรัมย์",
    "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.3 จ.สุรินทร์", "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.4 ปากช่อง",
    "ด่านควบคุมโรคติดต่อระหว่างประเทศพรมแดนช่องจอม จังหวัดสุรินทร์", "กลุ่มโรคไม่ติดต่อ", "งานควบคุมโรคเขตเมือง",
    "งานกฎหมาย", "กลุ่มโรคติดต่อเรื้อรัง", "กลุ่มห้องปฏิบัติการทางการแพทย์ด้านควบคุมโรค",
    "กลุ่มสื่อสารความเสี่ยงโรคและภัยสุขภาพ", "กลุ่มโรคจากการประกอบอาชีพและสิ่งแวดล้อม",
    "ศูนย์บริการเวชศาสตร์ป้องกัน", "กลุ่มบริหารทั่วไป", "ศูนย์ฝึกอบรมนักระบาดวิทยาภาคสนาม", "กลุ่มพัฒนานวัตกรรมและวิจัย"
]

# ===== เมนูหลัก =====
st.title("📋 โปรแกรมบันทึกข้อมูล (สคร.9)")
main_menu = st.sidebar.radio("เลือกหน้าเมนู", ["📊 Dashboard รวม", "🧭 การไปราชการ", "🕒 การลา", "🛠️ ผู้ดูแลระบบ (Admin)"])

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

    if not df_report.empty:
        st.markdown("### 🧾 Key Metrics การลา")
        unique_persons = df_report["ชื่อ-สกุล"].nunique()
        avg_leave = round(df_report["จำนวนวันลา"].mean(), 2)
        c1, c2 = st.columns(2)
        c1.metric("จำนวนบุคลากรที่ลา", f"{unique_persons} คน")
        c2.metric("เฉลี่ยวันลาต่อคน", f"{avg_leave} วัน")

        st.markdown("### 📊 จำนวนวันลาตามประเภท")
        leave_by_type = df_report.groupby("ประเภทการลา")["จำนวนวันลา"].sum().reset_index()
        st.altair_chart(alt.Chart(leave_by_type).mark_bar().encode(x="ประเภทการลา", y="จำนวนวันลา", color="ประเภทการลา"), use_container_width=True)

        st.markdown("### 🏢 จำนวนวันลาตามกลุ่มงาน")
        leave_by_group = df_report.groupby("กลุ่มงาน")["จำนวนวันลา"].sum().reset_index()
        st.altair_chart(alt.Chart(leave_by_group).mark_bar().encode(x="กลุ่มงาน", y="จำนวนวันลา", color="กลุ่มงาน"), use_container_width=True)
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
            st.error("⚠️ กรุณากรอกชื่อผู้ไปราชการ")
        else:
            # ✅ เพิ่มข้อมูลผู้หลักและผู้ร่วมเดินทางทุกคนลงตาราง
            all_rows = [data.copy()]
            for c in companions:
                d = data.copy()
                d["ชื่อ-สกุล"] = c["ชื่อ-สกุล"]
                d["กลุ่มงาน"] = c["กลุ่มงาน"]
                d["ผู้ร่วมเดินทาง"] = ""
                all_rows.append(d)

            df_scan = pd.concat([df_scan, pd.DataFrame(all_rows)], ignore_index=True)
            df_scan.to_excel(FILE_SCAN, index=False)
            st.success("✅ บันทึกข้อมูลการไปราชการเรียบร้อย (รวมผู้ร่วมเดินทางแล้ว)")

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
        df_report = pd.concat([df_report, pd.DataFrame([data])], ignore_index=True)
        df_report.to_excel(FILE_REPORT, index=False)
        st.success("✅ บันทึกข้อมูลการลาเรียบร้อย")

    if not df_report.empty:
        st.markdown("## 📊 สรุปข้อมูลการลา")
        st.dataframe(df_report.astype(str), use_container_width=True)

# ===== ผู้ดูแลระบบ ===== 
elif main_menu == "🛠️ ผู้ดูแลระบบ (Admin)":
    st.header("🔐 เข้าสู่ระบบผู้ดูแล")
    ADMIN_PASSWORD = "admin999"

    password = st.text_input("กรอกรหัสผ่าน", type="password")
    if password == ADMIN_PASSWORD:
        st.success("✅ เข้าสู่ระบบสำเร็จ")

        tab1, tab2, tab3 = st.tabs(["🧭 ข้อมูลไปราชการ", "🕒 ข้อมูลการลา", "📈 Dashboard กลุ่มงาน"])

        # --- แท็บ 1 : ข้อมูลไปราชการ ---
        with tab1:
            if not df_scan.empty:
                st.dataframe(df_scan.astype(str), use_container_width=True)
                idx = st.number_input("เลือกแถว (0 ถึง ...)", 0, len(df_scan)-1)
                action = st.radio("เลือกการทำงาน", ["-", "📝 แก้ไข", "🗑️ ลบ"])
                if action == "📝 แก้ไข":
                    edit_data = {}
                    for col in df_scan.columns:
                        edit_data[col] = st.text_input(col, str(df_scan.loc[idx, col]))
                    if st.button("💾 บันทึกการแก้ไข"):
                        for col in df_scan.columns:
                            df_scan.loc[idx, col] = edit_data[col]
                        df_scan.to_excel(FILE_SCAN, index=False)
                        st.success("✅ แก้ไขข้อมูลเรียบร้อย")
                elif action == "🗑️ ลบ":
                    if st.button("❌ ยืนยันลบ"):
                        df_scan = df_scan.drop(index=idx).reset_index(drop=True)
                        df_scan.to_excel(FILE_SCAN, index=False)
                        st.success("✅ ลบข้อมูลสำเร็จ")

        # --- แท็บ 2 : ข้อมูลการลา ---
        with tab2:
            if not df_report.empty:
                st.dataframe(df_report.astype(str), use_container_width=True)
                idx = st.number_input("เลือกแถวการลา", 0, len(df_report)-1, key="leave_idx")
                action = st.radio("เลือกการทำงาน", ["-", "📝 แก้ไข", "🗑️ ลบ"], key="leave_action")
                if action == "📝 แก้ไข":
                    edit_data = {}
                    for col in df_report.columns:
                        edit_data[col] = st.text_input(col, str(df_report.loc[idx, col]), key=f"edit_{col}")
                    if st.button("💾 บันทึกการแก้ไข", key="save_leave"):
                        for col in df_report.columns:
                            df_report.loc[idx, col] = edit_data[col]
                        df_report.to_excel(FILE_REPORT, index=False)
                        st.success("✅ แก้ไขข้อมูลเรียบร้อย")
                elif action == "🗑️ ลบ":
                    if st.button("❌ ยืนยันลบ", key="delete_leave"):
                        df_report = df_report.drop(index=idx).reset_index(drop=True)
                        df_report.to_excel(FILE_REPORT, index=False)
                        st.success("✅ ลบข้อมูลสำเร็จ")

        # --- แท็บ 3 : Dashboard ---
        with tab3:
            st.subheader("📈 Dashboard สรุปข้อมูลตามกลุ่มงาน")
            if not df_scan.empty or not df_report.empty:
                st.bar_chart(df_scan.groupby("กลุ่มงาน")["จำนวนวัน"].sum())
                st.bar_chart(df_report.groupby("กลุ่มงาน")["จำนวนวันลา"].sum())
    elif password:
        st.error("❌ รหัสผ่านไม่ถูกต้อง กรุณาลองใหม่")
