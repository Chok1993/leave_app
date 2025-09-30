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
            "ศูนย์บริการเวชศาสตร์ป้องกัน"
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
        
        submitted = st.form_submit_button("✅ บันทึกข้อมูล")

    if submitted:
        df_scan = pd.concat([df_scan, pd.DataFrame([data])], ignore_index=True)
        df_scan.to_excel(FILE_SCAN, index=False)
        st.success("บันทึกข้อมูลการไปราชการเรียบร้อย ✅")

    st.subheader("📊 ตารางข้อมูลการไปราชการ ล่าสุด")
    st.dataframe(df_scan.astype(str), use_container_width=True)

    # ===== ฟังก์ชันแก้ไข/ลบ =====
    if not df_scan.empty:
        st.write("✏️ จัดการข้อมูล")
        row_id = st.number_input("เลือกแถว (เริ่มจาก 0)", min_value=0, max_value=len(df_scan)-1, step=1)

        action = st.radio("เลือกการทำงาน", ["แก้ไข", "ลบ"], key="scan_action")

        if action == "แก้ไข":
            new_data = {}
            new_data["ชื่อ-สกุล"] = st.text_input("👤 ชื่อ-สกุล", value=df_scan.loc[row_id, "ชื่อ-สกุล"])
            new_data["กลุ่มงาน"] = st.selectbox("🏢 กลุ่มงาน", staff_groups, index=staff_groups.index(df_scan.loc[row_id, "กลุ่มงาน"]))
            new_data["ปี พ.ศ."] = st.number_input("📅 ปี พ.ศ.", value=int(df_scan.loc[row_id, "ปี พ.ศ."]))
            new_data["เดือน"] = st.number_input("📆 เดือน", value=int(df_scan.loc[row_id, "เดือน"]))
            new_data["วัน"] = st.number_input("📆 วัน", value=int(df_scan.loc[row_id, "วัน"]))
            new_data["กิจกรรม"] = st.selectbox("🏢 กิจกรรม", ["อบรม", "สัมมนา", "อื่นๆ"], index=["อบรม", "สัมมนา", "อื่นๆ"].index(df_scan.loc[row_id, "กิจกรรม"]))
            new_data["สถานที่"] = st.text_input("📍 สถานที่", value=df_scan.loc[row_id, "สถานที่"])
            new_data["ผู้จัด"] = st.text_input("👥 ผู้จัด", value=df_scan.loc[row_id, "ผู้จัด"])
            new_data["หมายเหตุ"] = st.text_input("📝 หมายเหตุ", value=df_scan.loc[row_id, "หมายเหตุ"])

            if st.button("💾 บันทึกการแก้ไข"):
                df_scan.loc[row_id] = new_data
                df_scan.to_excel(FILE_SCAN, index=False)
                st.success("แก้ไขข้อมูลเรียบร้อย ✅")

        elif action == "ลบ":
            if st.button("🗑️ ลบข้อมูล"):
                df_scan = df_scan.drop(row_id).reset_index(drop=True)
                df_scan.to_excel(FILE_SCAN, index=False)
                st.success("ลบข้อมูลเรียบร้อย ✅")


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

    # ===== ฟังก์ชันแก้ไข/ลบ =====
    if not df_report.empty:
        st.write("✏️ จัดการข้อมูล")
        row_id = st.number_input("เลือกแถว (เริ่มจาก 0)", min_value=0, max_value=len(df_report)-1, step=1, key="leave_row")

        action = st.radio("เลือกการทำงาน", ["แก้ไข", "ลบ"], key="leave_action")

        if action == "แก้ไข":
            new_data = {}
            new_data["ชื่อ-สกุล"] = st.text_input("👤 ชื่อ-สกุล", value=df_report.loc[row_id, "ชื่อ-สกุล"], key="edit_name")
            new_data["กลุ่มงาน"] = st.selectbox("🏢 กลุ่มงาน", staff_groups, index=staff_groups.index(df_report.loc[row_id, "กลุ่มงาน"]), key="edit_group")
            new_data["ประเภทการลา"] = st.selectbox("📌 ประเภทการลา", ["ลาป่วย", "ลากิจ", "ลาพักผ่อน", "อื่นๆ"], index=["ลาป่วย", "ลากิจ", "ลาพักผ่อน", "อื่นๆ"].index(df_report.loc[row_id, "ประเภทการลา"]), key="edit_leave")
            new_data["วันที่เริ่ม"] = st.date_input("📅 วันที่เริ่ม", value=pd.to_datetime(df_report.loc[row_id, "วันที่เริ่ม"]).date(), key="edit_start")
            new_data["วันที่สิ้นสุด"] = st.date_input("📅 วันที่สิ้นสุด", value=pd.to_datetime(df_report.loc[row_id, "วันที่สิ้นสุด"]).date(), key="edit_end")
            new_data["หมายเหตุ"] = st.text_area("📝 หมายเหตุ", value=df_report.loc[row_id, "หมายเหตุ"], key="edit_note")

            if st.button("💾 บันทึกการแก้ไข"):
                df_report.loc[row_id] = new_data
                df_report.to_excel(FILE_REPORT, index=False)
                st.success("แก้ไขข้อมูลเรียบร้อย ✅")

        elif action == "ลบ":
            if st.button("🗑️ ลบข้อมูล"):
                df_report = df_report.drop(row_id).reset_index(drop=True)
                df_report.to_excel(FILE_REPORT, index=False)
                st.success("ลบข้อมูลเรียบร้อย ✅")
