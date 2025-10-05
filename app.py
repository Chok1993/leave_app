import streamlit as st
import pandas as pd
import datetime as dt

# ===== ไฟล์เก็บข้อมูล =====
FILE_SCAN = "scan_report.xlsx"
FILE_REPORT = "leave_report.xlsx"

# โหลดไฟล์ ถ้าไม่มีให้สร้าง DataFrame ว่าง
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
    "ด่านควบคุมโรคติดต่อระหว่างประเทศพรมแดนช่องจอม จ.สุรินทร์",
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
menu = st.radio("เลือกประเภทข้อมูลที่ต้องการบันทึก", ["การไปราชการ", "การลา"])

# ===== ฟอร์มบันทึกการไปราชการ =====
if menu == "การไปราชการ":
    with st.form("scan_form"):
        data = {}
        data["ชื่อ-สกุล"] = st.text_input("👤 ชื่อ-สกุล")
        data["กลุ่มงาน"] = st.selectbox("🏢 กลุ่มงาน", staff_groups)
        data["วันที่เริ่ม"] = st.date_input("📅 วันที่เริ่มต้น", dt.date.today())
        data["วันที่สิ้นสุด"] = st.date_input("📅 วันที่สิ้นสุด", dt.date.today())
        data["จำนวนวัน"] = (data["วันที่สิ้นสุด"] - data["วันที่เริ่ม"]).days + 1

        # ✅ เปลี่ยนจาก selectbox → text_input
        data["กิจกรรม"] = st.text_input("🏢 ชื่อกิจกรรมที่ไป", placeholder="เช่น อบรมหลักสูตรการพัฒนาศักยภาพบุคลากร")

        data["สถานที่"] = st.text_input("📍 สถานที่")
        data["ผู้จัด"] = st.text_input("👥 ผู้จัด")
        data["หมายเหตุ"] = st.text_input("📝 หมายเหตุ")

        # เพิ่มผู้ร่วมเดินทาง
        st.markdown("### 👨‍👩‍👧‍👦 รายชื่อผู้ร่วมเดินทาง (หากไปเป็นหมู่คณะ)")
        num_people = st.number_input("จำนวนผู้ร่วมเดินทาง", min_value=0, max_value=20, step=1)
        companions = []
        for i in range(int(num_people)):
            name = st.text_input(f"ชื่อ-สกุลผู้ร่วมเดินทางคนที่ {i+1}")
            group = st.selectbox(f"กลุ่มงานของผู้ร่วมเดินทางคนที่ {i+1}", staff_groups, key=f"group_{i}")
            if name:
                companions.append({"ชื่อ-สกุล": name, "กลุ่มงาน": group})
        data["ผู้ร่วมเดินทาง"] = ", ".join([p["ชื่อ-สกุล"] for p in companions])

        submitted = st.form_submit_button("✅ บันทึกข้อมูล")

    if submitted:
        df_scan = pd.concat([df_scan, pd.DataFrame([data])], ignore_index=True)
        df_scan.to_excel(FILE_SCAN, index=False)
        st.success("บันทึกข้อมูลการไปราชการเรียบร้อย ✅")

    # ===== Dashboard การไปราชการ =====
    st.subheader("📊 สรุปข้อมูลการไปราชการ")

    if not df_scan.empty:
        st.write(f"📋 จำนวนรายการทั้งหมด: {len(df_scan)} รายการ")
        if "จำนวนวัน" in df_scan.columns:
            st.write(f"🗓️ รวมจำนวนวันทั้งหมด: {df_scan['จำนวนวัน'].sum()} วัน")

        st.divider()

        # กราฟรายเดือน
        if "วันที่เริ่ม" in df_scan.columns:
            df_scan["เดือน"] = pd.to_datetime(df_scan["วันที่เริ่ม"]).dt.month
            st.markdown("### 📆 จำนวนการไปราชการรายเดือน")
            st.bar_chart(df_scan["เดือน"].value_counts().sort_index())

        # กราฟกิจกรรม
        if "กิจกรรม" in df_scan.columns:
            st.markdown("### 🏢 ประเภทกิจกรรม")
            st.bar_chart(df_scan["กิจกรรม"].value_counts())

        # กราฟกลุ่มงาน
        if "กลุ่มงาน" in df_scan.columns:
            st.markdown("### 🧩 กลุ่มงานที่ไปราชการมากที่สุด")
            st.bar_chart(df_scan["กลุ่มงาน"].value_counts().sort_values(ascending=True))

        st.markdown("### 📋 ข้อมูลทั้งหมด")
        st.dataframe(df_scan.astype(str), use_container_width=True)
    else:
        st.info("ยังไม่มีข้อมูลการไปราชการ")

# ===== ฟอร์มบันทึกการลา =====
elif menu == "การลา":
    with st.form("leave_form"):
        data = {}
        data["ชื่อ-สกุล"] = st.text_input("👤 ชื่อ-สกุล")
        data["กลุ่มงาน"] = st.selectbox("🏢 กลุ่มงาน", staff_groups)
        data["ประเภทการลา"] = st.selectbox("📌 ประเภทการลา", ["ลาป่วย", "ลากิจ", "ลาพักผ่อน", "ลาอื่นๆ"])
        data["วันที่เริ่ม"] = st.date_input("📅 วันที่เริ่มลา", dt.date.today())
        data["วันที่สิ้นสุด"] = st.date_input("📅 วันที่สิ้นสุดการลา", dt.date.today())
        data["จำนวนวันลา"] = (data["วันที่สิ้นสุด"] - data["วันที่เริ่ม"]).days + 1
        data["หมายเหตุ"] = st.text_area("📝 หมายเหตุ")

        submitted = st.form_submit_button("✅ บันทึกข้อมูล")

    if submitted:
        df_report = pd.concat([df_report, pd.DataFrame([data])], ignore_index=True)
        df_report.to_excel(FILE_REPORT, index=False)
        st.success("บันทึกข้อมูลการลาเรียบร้อย ✅")

    # ===== Dashboard การลา =====
    st.subheader("📈 สรุปข้อมูลการลา")

    if not df_report.empty:
        st.write(f"📋 จำนวนรายการการลา: {len(df_report)} รายการ")
        st.write(f"🗓️ รวมจำนวนวันลาทั้งหมด: {df_report['จำนวนวันลา'].sum()} วัน")

        st.divider()

        # กราฟประเภทการลา
        if "ประเภทการลา" in df_report.columns:
            st.markdown("### 📊 ประเภทการลา")
            st.bar_chart(df_report["ประเภทการลา"].value_counts())

        # กราฟรายเดือน
        if "วันที่เริ่ม" in df_report.columns:
            df_report["เดือน"] = pd.to_datetime(df_report["วันที่เริ่ม"]).dt.month
            st.markdown("### 📆 จำนวนการลารายเดือน")
            st.bar_chart(df_report["เดือน"].value_counts().sort_index())

        # กราฟกลุ่มงาน
        if "กลุ่มงาน" in df_report.columns:
            st.markdown("### 🧩 กลุ่มงานที่มีการลามากที่สุด")
            st.bar_chart(df_report["กลุ่มงาน"].value_counts().sort_values(ascending=True))

        st.markdown("### 📋 ข้อมูลทั้งหมด")
        st.dataframe(df_report.astype(str), use_container_width=True)
    else:
        st.info("ยังไม่มีข้อมูลการลาในระบบ")

