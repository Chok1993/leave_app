import streamlit as st
import pandas as pd
import datetime as dt

# ===== ไฟล์เก็บข้อมูล =====
FILE_SCAN = "scan_report.xlsx"
FILE_REPORT = "leave_report.xlsx"

# ===== โหลดไฟล์ =====
def load_excel(file_path):
    try:
        return pd.read_excel(file_path)
    except:
        return pd.DataFrame()

df_scan = load_excel(FILE_SCAN)
df_report = load_excel(FILE_REPORT)

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
    ["📊 Dashboard รวม", "🧭 การไปราชการ", "🕒 การลา", "🛠️ จัดการข้อมูล (Admin)"]
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

    # ===== กราฟสรุป =====
    if not df_scan.empty or not df_report.empty:
        st.markdown("### 📊 กราฟเปรียบเทียบจำนวนวันไปราชการและการลา (ต่อเดือน)")
        df_chart = pd.DataFrame({
            "เดือน": list(range(1, 13)),
            "วันไปราชการ": [0]*12,
            "วันลา": [0]*12
        })
        if "วันที่เริ่ม" in df_scan.columns:
            for _, row in df_scan.iterrows():
                if pd.notna(row["วันที่เริ่ม"]):
                    month = pd.to_datetime(row["วันที่เริ่ม"]).month
                    df_chart.loc[month-1, "วันไปราชการ"] += row.get("จำนวนวัน", 0)
        if "วันที่เริ่ม" in df_report.columns:
            for _, row in df_report.iterrows():
                if pd.notna(row["วันที่เริ่ม"]):
                    month = pd.to_datetime(row["วันที่เริ่ม"]).month
                    df_chart.loc[month-1, "วันลา"] += row.get("จำนวนวันลา", 0)
        st.bar_chart(df_chart.set_index("เดือน"))

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
            df_scan = pd.concat([df_scan, pd.DataFrame([data])], ignore_index=True)
            df_scan.to_excel(FILE_SCAN, index=False)
            st.success("✅ บันทึกข้อมูลการไปราชการเรียบร้อย")

    # Dashboard ย่อย
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
            df_report = pd.concat([df_report, pd.DataFrame([data])], ignore_index=True)
            df_report.to_excel(FILE_REPORT, index=False)
            st.success("✅ บันทึกข้อมูลการลาเรียบร้อย")

    # Dashboard ย่อย
    if not df_report.empty:
        st.markdown("## 📊 สรุปข้อมูลการลา")
        st.write(f"📋 ทั้งหมด {len(df_report)} รายการ รวมลา {df_report['จำนวนวันลา'].sum()} วัน")
        st.dataframe(df_report.astype(str), use_container_width=True)

# ===== โหมดผู้ดูแลระบบ (Admin) =====
elif main_menu == "🛠️ จัดการข้อมูล (Admin)":
    st.header("🔑 เข้าสู่ระบบผู้ดูแล")

    # ---- ระบบล็อกอิน ----
    password = st.text_input("กรอกรหัสผ่านเพื่อเข้าสู่โหมดผู้ดูแลระบบ", type="password")
    ADMIN_PASSWORD = "skc9admin@2025"

    if password == ADMIN_PASSWORD:
        st.success("✅ เข้าสู่โหมดผู้ดูแลเรียบร้อย")

        admin_menu = st.radio("เลือกข้อมูลที่ต้องการจัดการ", ["การไปราชการ", "การลา"])
        df_target = df_scan if admin_menu == "การไปราชการ" else df_report
        file_target = FILE_SCAN if admin_menu == "การไปราชการ" else FILE_REPORT

        if df_target.empty:
            st.warning("ยังไม่มีข้อมูลในหมวดนี้")
        else:
            st.dataframe(df_target.astype(str), use_container_width=True)

            row_id = st.number_input("เลือกแถวที่ต้องการแก้ไข/ลบ", 0, len(df_target)-1, 0)

            action = st.radio("เลือกการทำงาน", ["แก้ไข", "ลบ"], horizontal=True)

            if action == "แก้ไข":
                with st.form("edit_form"):
                    edited = {}
                    for col in df_target.columns:
                        value = df_target.loc[row_id, col]
                        edited[col] = st.text_input(f"{col}", value=str(value))
                    save = st.form_submit_button("💾 บันทึกการแก้ไข")

                if save:
                    for col in df_target.columns:
                        df_target.loc[row_id, col] = edited[col]
                    df_target.to_excel(file_target, index=False)
                    st.success("✅ แก้ไขข้อมูลสำเร็จแล้ว!")

            elif action == "ลบ":
                if st.button("🗑️ ยืนยันการลบข้อมูลนี้"):
                    df_target = df_target.drop(row_id).reset_index(drop=True)
                    df_target.to_excel(file_target, index=False)
                    st.success("🗑️ ลบข้อมูลเรียบร้อยแล้ว!")
    elif password:
        st.error("❌ รหัสผ่านไม่ถูกต้อง กรุณาลองใหม่")
