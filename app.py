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
    "กลุ่มพัฒนานวัตกรรมและวิจัย",
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
        data["ปี พ.ศ."] = st.number_input("📅 ปี พ.ศ.", min_value=2560, max_value=2600, value=2568)
        data["เดือน"] = st.selectbox("📌 เดือน", list(range(1, 13)))
        data["วัน"] = st.number_input("📆 วัน", min_value=1, max_value=31, value=dt.date.today().day)
        data["กิจกรรม"] = st.selectbox("🏢 กิจกรรม", ["อบรม", "สัมมนา", "อื่นๆ"])
        data["สถานที่"] = st.text_input("📍 สถานที่")
        data["ผู้จัด"] = st.text_input("👥 ผู้จัด")
        data["หมายเหตุ"] = st.text_input("📝 หมายเหตุ")
        
        # วันที่เริ่มและวันที่สิ้นสุด
        data["วันที่เริ่ม"] = st.date_input("📅 วันที่เริ่ม", dt.date.today())
        data["วันที่สิ้นสุด"] = st.date_input("📅 วันที่สิ้นสุด", dt.date.today())
        
        # คำนวณจำนวนวันที่ไป
        start_date = data["วันที่เริ่ม"]
        end_date = data["วันที่สิ้นสุด"]
        if start_date and end_date:
            days_count = (end_date - start_date).days + 1  # รวมวันเริ่มต้นด้วย
            st.write(f"📅 จำนวนวันที่ไป: {days_count} วัน")
        else:
            days_count = 0  # ถ้าไม่มีวันที่เริ่มหรือสิ้นสุด
        
        # ฟิลด์สำหรับการกรอกชื่อหมู่คณะ
        num_group_members = st.number_input("จำนวนสมาชิกในกลุ่ม", min_value=1, max_value=10, value=1, step=1)
        group_members = []
        
        for i in range(num_group_members):
            member_name = st.text_input(f"ชื่อ-นามสกุล สมาชิกที่ {i+1}", key=f"member_{i}")
            if member_name:
                group_members.append(member_name)
        
        data["สมาชิกกลุ่ม"] = ", ".join(group_members)  # รวมชื่อสมาชิกทั้งหมดไว้ในฟิลด์เดียว
        
        submitted = st.form_submit_button("✅ บันทึกข้อมูล")

    if submitted:
        # เพิ่มจำนวนวันที่ไปในการเก็บข้อมูล
        data["จำนวนวันที่ไป"] = days_count
        df_scan = pd.concat([df_scan, pd.DataFrame([data])], ignore_index=True)
        df_scan.to_excel(FILE_SCAN, index=False)
        st.success("บันทึกข้อมูลการไปราชการเรียบร้อย ✅")

    st.subheader("📊 ตารางข้อมูลการไปราชการ ล่าสุด")
    st.dataframe(df_scan.astype(str), use_container_width=True)


# ===== ฟอร์มบันทึกการลา =====
elif menu == "การลา":
    with st.form("leave_form"):
        data = {}
        data["ชื่อ-สกุล"] = st.text_input("👤 ชื่อ-สกุล")
        data["กลุ่มงาน"] = st.selectbox("🏢 กลุ่มงาน", staff_groups)
        data["ประเภทการลา"] = st.selectbox("📌 ประเภทการลา", ["ลาป่วย", "ลากิจ", "ลาพักผ่อน", "อื่นๆ"])
        data["วันที่เริ่ม"] = st.date_input("📅 วันที่เริ่ม", dt.date.today())
        data["วันที่สิ้นสุด"] = st.date_input("📅 วันที่สิ้นสุด", dt.date.today())
        data["หมายเหตุ"] = st.text_area("📝 หมายเหตุ")

        submitted = st.form_submit_button("✅ บันทึกข้อมูล")

    if submitted:
        df_report = pd.concat([df_report, pd.DataFrame([data])], ignore_index=True)
        df_report.to_excel(FILE_REPORT, index=False)
        st.success("บันทึกข้อมูลการลาเรียบร้อย ✅")

    st.subheader("📊 ตารางข้อมูลการลาล่าสุด")
    st.dataframe(df_report.astype(str), use_container_width=True)
