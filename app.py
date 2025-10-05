import streamlit as st
import pandas as pd
import datetime as dt
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
main_menu = st.sidebar.radio("เลือกหน้าเมนู", ["📊 Dashboard รวม", "🧭 การไปราชการ", "🕒 การลา", "👩‍💼 Admin"])

# ===== ฟังก์ชันกรองข้อมูลตามเดือน =====
def filter_by_month(df, start_col="วันที่เริ่ม"):
    if start_col in df.columns:
        df = df.copy()
        df[start_col] = pd.to_datetime(df[start_col], errors="coerce")
        return df
    return pd.DataFrame()

# ===== Dashboard รวม =====
if main_menu == "📊 Dashboard รวม":
    st.header("📈 Dashboard สรุปข้อมูลทั้งหมด")

    month_choice = st.selectbox("เลือกเดือน", list(range(1, 13)), format_func=lambda x: f"เดือน {x}")
    year_choice = st.number_input("ปี พ.ศ.", min_value=2560, max_value=2600, value=2568)

    df_scan_f = filter_by_month(df_scan)
    df_report_f = filter_by_month(df_report)

    # คำนวณสรุป
    total_travel = len(df_scan_f)
    total_leave = len(df_report_f)
    total_travel_days = df_scan_f["จำนวนวัน"].sum() if "จำนวนวัน" in df_scan_f.columns else 0
    total_leave_days = df_report_f["จำนวนวันลา"].sum() if "จำนวนวันลา" in df_report_f.columns else 0

    col1, col2 = st.columns(2)
    col1.metric("จำนวนผู้ไปราชการ", f"{total_travel} คน")
    col1.metric("จำนวนวันไปราชการรวม", f"{total_travel_days} วัน")
    col2.metric("จำนวนผู้ลา", f"{total_leave} คน")
    col2.metric("จำนวนวันลารวม", f"{total_leave_days} วัน")

    # กราฟรวม
    if not df_scan_f.empty or not df_report_f.empty:
        st.markdown("### 📊 เปรียบเทียบวันลา-ไปราชการต่อเดือน")
        df_chart = pd.DataFrame({
            "ประเภท": ["ไปราชการ", "ลา"],
            "จำนวนวัน": [total_travel_days, total_leave_days]
        }).set_index("ประเภท")
        st.bar_chart(df_chart)

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
            st.error("⚠️ กรุณากรอกชื่อ-นามสกุลของผู้ไปราชการ")
        elif int(num_people) > 0 and len(companions) < int(num_people):
            st.error("⚠️ กรุณากรอกชื่อ-นามสกุลผู้ร่วมเดินทางให้ครบ")
        else:
            df_scan = df_scan.dropna(how='all')
            df_scan = pd.concat([df_scan, pd.DataFrame([data])], ignore_index=True)
            df_scan.to_excel(FILE_SCAN, index=False)
            st.success("✅ บันทึกข้อมูลการไปราชการเรียบร้อย")

    # ===== Dashboard ย่อย =====
    if not df_scan.empty:
        st.markdown("## 📊 Dashboard การไปราชการ")
        total_records = len(df_scan)
        total_days = df_scan["จำนวนวัน"].sum()
        unique_people = df_scan["ชื่อ-สกุล"].nunique()
        avg_days = total_days / unique_people if unique_people > 0 else 0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("📋 จำนวนครั้งที่ไปราชการ", f"{total_records}")
        c2.metric("🗓️ จำนวนวันรวม", f"{total_days} วัน")
        c3.metric("👥 จำนวนบุคลากร", f"{unique_people} คน")
        c4.metric("⏱️ เฉลี่ยวันต่อคน", f"{avg_days:.2f} วัน")

        st.markdown("### 🏢 กราฟจำนวนวันไปราชการตามกลุ่มงาน")
        travel_group_chart = df_scan.groupby("กลุ่มงาน")["จำนวนวัน"].sum().sort_values(ascending=False)
        st.bar_chart(travel_group_chart)

        st.markdown("### 📄 ตารางข้อมูลการไปราชการ")
        st.dataframe(df_scan.astype(str), use_container_width=True)

        # ปุ่มดาวน์โหลดรายงาน
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_scan.to_excel(writer, sheet_name='ข้อมูลไปราชการ', index=False)
            summary = pd.DataFrame({
                "หัวข้อ": ["จำนวนครั้ง", "จำนวนวันรวม", "จำนวนบุคลากร", "เฉลี่ยวันต่อคน"],
                "ค่า": [total_records, total_days, unique_people, round(avg_days, 2)]
            })
            summary.to_excel(writer, sheet_name='สรุปภาพรวม', index=False)

        st.download_button(
            label="📥 ดาวน์โหลดรายงานไปราชการ (Excel)",
            data=output.getvalue(),
            file_name=f"รายงานไปราชการ_{dt.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

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

    # Dashboard การลา
    if not df_report.empty:
        st.markdown("## 📊 Dashboard การลา")
        total_records = len(df_report)
        total_days = df_report["จำนวนวันลา"].sum()
        unique_people = df_report["ชื่อ-สกุล"].nunique()
        avg_days = total_days / unique_people if unique_people > 0 else 0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("📋 จำนวนครั้งที่ลา", f"{total_records} ครั้ง")
        c2.metric("🗓️ จำนวนวันลารวม", f"{total_days} วัน")
        c3.metric("👥 จำนวนบุคลากรที่ลา", f"{unique_people} คน")
        c4.metric("⏱️ เฉลี่ยวันลาต่อคน", f"{avg_days:.2f} วัน")

        st.markdown("### 📈 กราฟจำนวนวันลาตามประเภท")
        leave_type_chart = df_report.groupby("ประเภทการลา")["จำนวนวันลา"].sum()
        st.bar_chart(leave_type_chart)

        st.markdown("### 🏢 กราฟจำนวนวันลาตามกลุ่มงาน")
        leave_group_chart = df_report.groupby("กลุ่มงาน")["จำนวนวันลา"].sum().sort_values(ascending=False)
        st.bar_chart(leave_group_chart)

        st.markdown("### 📄 ตารางข้อมูลการลา")
        st.dataframe(df_report.astype(str), use_container_width=True)

        # ปุ่มดาวน์โหลดรายงาน
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_report.to_excel(writer, sheet_name='ข้อมูลการลา', index=False)
            summary = pd.DataFrame({
                "หัวข้อ": ["จำนวนครั้งที่ลา", "จำนวนวันลารวม", "จำนวนบุคลากร", "เฉลี่ยวันลาต่อคน"],
                "ค่า": [total_records, total_days, unique_people, round(avg_days, 2)]
            })
            summary.to_excel(writer, sheet_name='สรุปภาพรวม', index=False)

        st.download_button(
            label="📥 ดาวน์โหลดรายงานการลา (Excel)",
            data=output.getvalue(),
            file_name=f"รายงานการลา_{dt.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ===== เมนู Admin =====
elif main_menu == "👩‍💼 Admin":
    st.header("🛠️ เมนูผู้ดูแลระบบ (หลังบ้าน)")
    st.info("หน้านี้สงวนสำหรับผู้ดูแลระบบเพื่อจัดการและตรวจสอบข้อมูล")
    st.write("ในอนาคตสามารถเพิ่มฟังก์ชัน เช่น ลบ/แก้ไขข้อมูล หรือดูสถิติแบบละเอียดได้ที่นี่")
