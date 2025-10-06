import streamlit as st
import pandas as pd
import datetime as dt
import altair as alt
import matplotlib.pyplot as plt
from fpdf import FPDF
import tempfile

# ===== รหัสผ่านผู้ดูแลระบบ =====
ADMIN_PASSWORD = "DDC9admin"

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

    if not df_report.empty:
        st.subheader("🧾 Key Metrics การลา")
        unique_persons = df_report["ชื่อ-สกุล"].nunique()
        avg_leave = round(df_report["จำนวนวันลา"].mean(), 2) if "จำนวนวันลา" in df_report.columns else 0

        c1, c2 = st.columns(2)
        c1.metric("จำนวนบุคลากรที่ลา", f"{unique_persons} คน")
        c2.metric("เฉลี่ยวันลาต่อคน", f"{avg_leave} วัน")

        st.markdown("### 📊 จำนวนวันลาตามประเภท")
        leave_by_type = df_report.groupby("ประเภทการลา")["จำนวนวันลา"].sum().reset_index()
        chart1 = alt.Chart(leave_by_type).mark_bar().encode(
            x=alt.X("ประเภทการลา:N", title="ประเภทการลา"),
            y=alt.Y("จำนวนวันลา:Q", title="จำนวนวันลา"),
            color="ประเภทการลา"
        )
        st.altair_chart(chart1, use_container_width=True)

        st.markdown("### 🏢 จำนวนวันลาตามกลุ่มงาน")
        leave_by_group = df_report.groupby("กลุ่มงาน")["จำนวนวันลา"].sum().reset_index()
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
            main_entry = pd.DataFrame([data])  # รายการหลัก

            # ✅ เพิ่ม record สำหรับผู้ร่วมเดินทางโดยอัตโนมัติ
            companion_entries = []
            for p in companions:
                companion_data = {
                    "ชื่อ-สกุล": p["ชื่อ-สกุล"],
                    "กลุ่มงาน": p["กลุ่มงาน"],
                    "กิจกรรม": data["กิจกรรม"] + " (ร่วม)",
                    "ผู้ร่วมเดินทาง": data["ชื่อ-สกุล"],  # แสดงชื่อหัวหน้าแทน
                    "วันที่เริ่ม": data["วันที่เริ่ม"],
                    "วันที่สิ้นสุด": data["วันที่สิ้นสุด"],
                    "จำนวนวัน": data["จำนวนวัน"],
                    "หมายเหตุ": "ร่วมเดินทางกับ " + data["ชื่อ-สกุล"]
                }
                companion_entries.append(companion_data)

            # รวมทั้งหมด (หัวหน้า + ผู้ร่วม)
            combined = pd.concat([main_entry, pd.DataFrame(companion_entries)], ignore_index=True)
            df_scan = pd.concat([df_scan, combined], ignore_index=True)

            # บันทึกลง Excel
            df_scan.to_excel(FILE_SCAN, index=False)
            st.success("✅ บันทึกข้อมูลการไปราชการเรียบร้อย (รวมผู้ร่วมเดินทางแล้ว)")



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

# ===== ผู้ดูแลระบบ ===== 
elif main_menu == "🛠️ ผู้ดูแลระบบ (Admin)":
    st.header("🔐 เข้าสู่ระบบผู้ดูแล")
    ADMIN_PASSWORD = "admin999"  # ✅ ตั้งรหัสผ่านตรงนี้ได้เลย

    password = st.text_input("กรอกรหัสผ่าน", type="password")

    if password == ADMIN_PASSWORD:
        st.success("✅ เข้าสู่ระบบสำเร็จแล้ว")

        tab1, tab2, tab3 = st.tabs(["🧭 ข้อมูลไปราชการ", "🕒 ข้อมูลการลา", "📈 Dashboard กลุ่มงาน"])

        # ========== TAB 1 : ข้อมูลการไปราชการ ==========
        with tab1:
            st.subheader("🧭 ข้อมูลการไปราชการทั้งหมด")

            if not df_scan.empty:
                st.dataframe(df_scan.astype(str), use_container_width=True)

                # ✅ เลือกแถวที่ต้องการแก้ไขหรือลบ
                selected_index = st.number_input("🔢 เลือกลำดับแถว (Row Index)", min_value=0, max_value=len(df_scan) - 1, step=1)
                selected_row = df_scan.iloc[selected_index]

                st.write("✏️ ข้อมูลที่เลือก:")
                st.json(selected_row.to_dict())

                edit_mode = st.radio("ต้องการทำอะไร", ["-", "📝 แก้ไขข้อมูล", "🗑️ ลบข้อมูล"])

                # ✅ แก้ไขข้อมูล
                if edit_mode == "📝 แก้ไขข้อมูล":
                    edit_data = {}
                    for col in df_scan.columns:
                        edit_data[col] = st.text_input(f"{col}", str(selected_row[col]))
                    if st.button("💾 บันทึกการแก้ไข"):
                        for col in df_scan.columns:
                            df_scan.at[selected_index, col] = edit_data[col]
                        df_scan.to_excel(FILE_SCAN, index=False)
                        st.success("✅ แก้ไขข้อมูลเรียบร้อยแล้ว")

                # ✅ ลบข้อมูล
                elif edit_mode == "🗑️ ลบข้อมูล":
                    if st.button("❌ ยืนยันการลบแถวนี้"):
                        df_scan = df_scan.drop(index=selected_index).reset_index(drop=True)
                        df_scan.to_excel(FILE_SCAN, index=False)
                        st.success("✅ ลบข้อมูลเรียบร้อยแล้ว")

            else:
                st.info("ยังไม่มีข้อมูลการไปราชการ")

        # ========== TAB 2 : ข้อมูลการลา ==========
        with tab2:
            st.subheader("🕒 ข้อมูลการลาทั้งหมด")

            if not df_report.empty:
                st.dataframe(df_report.astype(str), use_container_width=True)

                selected_index = st.number_input("🔢 เลือกลำดับแถว (Row Index)", min_value=0, max_value=len(df_report) - 1, step=1, key="leave_idx")
                selected_row = df_report.iloc[selected_index]
                st.write("✏️ ข้อมูลที่เลือก:")
                st.json(selected_row.to_dict())

                edit_mode = st.radio("ต้องการทำอะไร", ["-", "📝 แก้ไขข้อมูล", "🗑️ ลบข้อมูล"], key="leave_action")

                if edit_mode == "📝 แก้ไขข้อมูล":
                    edit_data = {}
                    for col in df_report.columns:
                        edit_data[col] = st.text_input(f"{col}", str(selected_row[col]), key=f"edit_{col}")
                    if st.button("💾 บันทึกการแก้ไข", key="save_leave_edit"):
                        for col in df_report.columns:
                            df_report.at[selected_index, col] = edit_data[col]
                        df_report.to_excel(FILE_REPORT, index=False)
                        st.success("✅ แก้ไขข้อมูลเรียบร้อยแล้ว")

                elif edit_mode == "🗑️ ลบข้อมูล":
                    if st.button("❌ ยืนยันการลบแถวนี้", key="del_leave_row"):
                        df_report = df_report.drop(index=selected_index).reset_index(drop=True)
                        df_report.to_excel(FILE_REPORT, index=False)
                        st.success("✅ ลบข้อมูลเรียบร้อยแล้ว")

            else:
                st.info("ยังไม่มีข้อมูลการลา")

        # ========== TAB 3 : Dashboard ==========
        with tab3:
            st.subheader("📈 Dashboard สรุปข้อมูลตามกลุ่มงาน")

            this_year = dt.date.today().year + 543
            year_choice = st.selectbox("เลือกปี พ.ศ.", list(range(this_year - 3, this_year + 1)), index=3)
            month_choice = st.selectbox("เลือกเดือน", list(range(1, 13)), format_func=lambda x: f"เดือน {x}")

            # ===== ฟังก์ชันกรองข้อมูล =====
            def filter_data_by_month(df, start_col, end_col):
                df = df.copy()
                df[start_col] = pd.to_datetime(df[start_col], errors="coerce")
                df[end_col] = pd.to_datetime(df[end_col], errors="coerce")
                df["ปี"] = df[start_col].dt.year + 543


    elif password:
        st.error("❌ รหัสผ่านไม่ถูกต้อง กรุณาลองใหม่")





