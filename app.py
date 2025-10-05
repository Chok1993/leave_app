import streamlit as st
import pandas as pd
import datetime as dt

# ===== Data Files =====
FILE_SCAN = "scan_report.xlsx"
FILE_REPORT = "leave_report.xlsx"

# ===== Load Files =====
# Try to load the travel report, create an empty DataFrame if it doesn't exist.
try:
    df_scan = pd.read_excel(FILE_SCAN)
except FileNotFoundError:
    df_scan = pd.DataFrame()

# Try to load the leave report, create an empty DataFrame if it doesn't exist.
try:
    df_report = pd.read_excel(FILE_REPORT)
except FileNotFoundError:
    df_report = pd.DataFrame()

# ===== List of Staff Groups =====
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

# ===== Main Menu =====
st.title("📋 โปรแกรมบันทึกข้อมูล (สคร.9)")
main_menu = st.sidebar.radio("เลือกหน้าเมนู", ["📊 Dashboard รวม", "🧭 การไปราชการ", "🕒 การลา", "🛠️ แอดมินจัดการข้อมูล"])

# ===== Main Dashboard =====
if main_menu == "📊 Dashboard รวม":
    st.header("📈 Dashboard สรุปข้อมูลทั้งหมด")

    # ========== Select Year for Summary ==========
    current_year = dt.date.today().year + 543  # Buddhist Era Year
    all_years = sorted(
        set(
            [current_year]
            + [pd.to_datetime(x).year + 543 for x in df_scan["วันที่เริ่ม"].dropna()] if not df_scan.empty and "วันที่เริ่ม" in df_scan.columns else [current_year]
            + [pd.to_datetime(x).year + 543 for x in df_report["วันที่เริ่ม"].dropna()] if not df_report.empty and "วันที่เริ่ม" in df_report.columns else [current_year]
        )
    )
    selected_year = st.selectbox("เลือกปี พ.ศ. สำหรับดูสรุป", all_years, index=len(all_years)-1)

    # ========== Summary Totals ==========
    total_travel = len(df_scan)
    total_leave = len(df_report)
    total_travel_days = df_scan["จำนวนวัน"].sum() if "จำนวนวัน" in df_scan.columns else 0
    total_leave_days = df_report["จำนวนวันลา"].sum() if "จำนวนวันลา" in df_report.columns else 0

    col1, col2 = st.columns(2)
    col1.metric("จำนวนผู้ไปราชการ", f"{total_travel:,} คน")
    col1.metric("จำนวนวันไปราชการรวม", f"{total_travel_days:,} วัน")
    col2.metric("จำนวนผู้ลา", f"{total_leave:,} คน")
    col2.metric("จำนวนวันลารวม", f"{total_leave_days:,} วัน")

    # ========== Monthly Comparison Chart ==========
    df_chart = pd.DataFrame({
        "เดือน": list(range(1, 13)),
        "วันไปราชการ": [0]*12,
        "วันลา": [0]*12
    })

    if not df_scan.empty and "วันที่เริ่ม" in df_scan.columns:
        for _, row in df_scan.iterrows():
            if pd.notna(row["วันที่เริ่ม"]):
                start_date = pd.to_datetime(row["วันที่เริ่ม"])
                if start_date.year + 543 == selected_year:
                    df_chart.loc[start_date.month-1, "วันไปราชการ"] += row.get("จำนวนวัน", 0)

    if not df_report.empty and "วันที่เริ่ม" in df_report.columns:
        for _, row in df_report.iterrows():
            if pd.notna(row["วันที่เริ่ม"]):
                start_date = pd.to_datetime(row["วันที่เริ่ม"])
                if start_date.year + 543 == selected_year:
                    df_chart.loc[start_date.month-1, "วันลา"] += row.get("จำนวนวันลา", 0)

    # Convert month numbers to Thai names
    month_names = [
        "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
        "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
    ]
    df_chart["เดือน"] = month_names

    # ========== Display Results ==========
    st.markdown("### 📊 กราฟเปรียบเทียบจำนวนวันไปราชการและวันลา (รายเดือน)")
    st.bar_chart(df_chart.set_index("เดือน"))

    # Summary Table
    st.markdown("### 📋 ตารางสรุปข้อมูลรายเดือน")
    df_chart["รวมทั้งหมด (วัน)"] = df_chart["วันไปราชการ"] + df_chart["วันลา"]
    st.dataframe(df_chart, use_container_width=True)

    # ========== Current Month Summary ==========
    this_month = dt.date.today().month
    st.markdown("---")
    st.subheader(f"📅 สรุปเฉพาะเดือน {month_names[this_month-1]} {selected_year}")
    month_travel = df_chart.loc[this_month-1, "วันไปราชการ"]
    month_leave = df_chart.loc[this_month-1, "วันลา"]
    col1, col2 = st.columns(2)
    col1.metric("วันไปราชการเดือนนี้", f"{month_travel} วัน")
    col2.metric("วันลาประจำเดือน", f"{month_leave} วัน")

# ===== Travel Record Form =====
elif main_menu == "🧭 การไปราชการ":
    st.header("🧾 แบบฟอร์มบันทึกการไปราชการ")
    with st.form("scan_form"):
        data = {}
        data["ชื่อ-สกุล"] = st.text_input("👤 ชื่อ-สกุล")
        data["กลุ่มงาน"] = st.selectbox("🏢 กลุ่มงาน", staff_groups)
        data["ปี พ.ศ."] = st.number_input("📅 ปี พ.ศ.", min_value=2560, max_value=2600, value=dt.date.today().year + 543)
        data["กิจกรรม"] = st.text_input("🏢 กิจกรรม (พิมพ์ชื่อกิจกรรม)")
        data["สถานที่"] = st.text_input("📍 สถานที่")
        data["ผู้จัด"] = st.text_input("👥 ผู้จัด")
        data["วันที่เริ่ม"] = st.date_input("📅 วันที่เริ่ม", dt.date.today())
        data["วันที่สิ้นสุด"] = st.date_input("📅 วันที่สิ้นสุด", dt.date.today())
        
        if data["วันที่สิ้นสุด"] < data["วันที่เริ่ม"]:
            st.error("วันที่สิ้นสุดต้องไม่น้อยกว่าวันที่เริ่ม")
            data["จำนวนวัน"] = 0
        else:
            data["จำนวนวัน"] = (data["วันที่สิ้นสุด"] - data["วันที่เริ่ม"]).days + 1
        
        st.write(f"📆 รวม {data['จำนวนวัน']} วัน")

        # Companions
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
        elif data["จำนวนวัน"] <= 0:
            st.error("⚠️ จำนวนวันต้องมากกว่า 0")
        else:
            df_scan = df_scan.dropna(how='all')
            df_scan = pd.concat([df_scan, pd.DataFrame([data])], ignore_index=True)
            df_scan.to_excel(FILE_SCAN, index=False)
            st.success("✅ บันทึกข้อมูลการไปราชการเรียบร้อย")

    # Sub-Dashboard
    if not df_scan.empty:
        st.markdown("## 📊 สรุปข้อมูลการไปราชการ")
        st.write(f"📋 ทั้งหมด {len(df_scan)} รายการ รวม {df_scan['จำนวนวัน'].sum()} วัน")
        st.dataframe(df_scan.astype(str), use_container_width=True)

# ===== Leave Record Form =====
elif main_menu == "🕒 การลา":
    st.header("📝 แบบฟอร์มบันทึกการลา")
    with st.form("leave_form"):
        data = {}
        data["ชื่อ-สกุล"] = st.text_input("👤 ชื่อ-สกุล")
        data["กลุ่มงาน"] = st.selectbox("🏢 กลุ่มงาน", staff_groups)
        data["ประเภทการลา"] = st.selectbox("📌 ประเภทการลา", ["ลาป่วย", "ลากิจ", "ลาพักผ่อน", "อื่นๆ"])
        data["วันที่เริ่ม"] = st.date_input("📅 วันที่เริ่ม", dt.date.today())
        data["วันที่สิ้นสุด"] = st.date_input("📅 วันที่สิ้นสุด", dt.date.today())
        
        if data["วันที่สิ้นสุด"] < data["วันที่เริ่ม"]:
            st.error("วันที่สิ้นสุดต้องไม่น้อยกว่าวันที่เริ่ม")
            data["จำนวนวันลา"] = 0
        else:
            data["จำนวนวันลา"] = (data["วันที่สิ้นสุด"] - data["วันที่เริ่ม"]).days + 1
        
        st.write(f"🗓️ รวมลา {data['จำนวนวันลา']} วัน")
        data["หมายเหตุ"] = st.text_area("📝 หมายเหตุ")

        submitted = st.form_submit_button("✅ บันทึกข้อมูล")

    if submitted:
        if not data["ชื่อ-สกุล"]:
            st.error("⚠️ กรุณากรอกชื่อ-นามสกุลของผู้ลา")
        elif data["จำนวนวันลา"] <= 0:
            st.error("⚠️ จำนวนวันลาต้องมากกว่า 0")
        else:
            df_report = df_report.dropna(how='all')
            df_report = pd.concat([df_report, pd.DataFrame([data])], ignore_index=True)
            df_report.to_excel(FILE_REPORT, index=False)
            st.success("✅ บันทึกข้อมูลการลาเรียบร้อย")

    # Sub-Dashboard with KPIs
    if not df_report.empty:
        st.markdown("---")
        st.markdown("## 📊 Dashboard และตัวชี้วัดข้อมูลการลา (KPIs)")

        # Key Metrics
        total_records = len(df_report)
        total_days = int(df_report['จำนวนวันลา'].sum())
        unique_staff = df_report['ชื่อ-สกุล'].nunique()
        avg_days_per_person = total_days / unique_staff if unique_staff > 0 else 0

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("จำนวนครั้งที่ลาทั้งหมด", f"{total_records:,}")
        col2.metric("จำนวนวันลารวม", f"{total_days:,}")
        col3.metric("จำนวนบุคลากรที่ลา", f"{unique_staff:,}")
        col4.metric("เฉลี่ยวันลาต่อคน", f"{avg_days_per_person:,.2f}")

        st.markdown("---")

        # Charts
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("### 📈 จำนวนวันลาตามประเภท")
            leave_by_type = df_report.groupby('ประเภทการลา')['จำนวนวันลา'].sum()
            st.bar_chart(leave_by_type)

        with col2:
            st.markdown("### 🏢 จำนวนวันลาตามกลุ่มงาน")
            leave_by_group = df_report.groupby('กลุ่มงาน')['จำนวนวันลา'].sum()
            st.bar_chart(leave_by_group)
        
        st.markdown("### 📋 ข้อมูลการลาทั้งหมด")
        st.dataframe(df_report.astype(str), use_container_width=True)

# ===== Admin Data Management =====
elif main_menu == "🛠️ แอดมินจัดการข้อมูล":
    st.header("🔐 ส่วนจัดการข้อมูล (Admin)")
    password = st.text_input("🔑 ใส่รหัสผ่านเพื่อเข้าใช้งาน", type="password")
    
    # Simple password check
    if password == "admin2568":
        st.success("✅ เข้าสู่โหมดแอดมินแล้ว")
        tab1, tab2 = st.tabs(["🧭 ข้อมูลไปราชการ", "🕒 ข้อมูลการลา"])
        
        with tab1:
            st.dataframe(df_scan.astype(str), use_container_width=True)
            # Use st.download_button for direct download
            csv = df_scan.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 ดาวน์โหลดข้อมูลไปราชการเป็น CSV",
                data=csv,
                file_name='scan_report.csv',
                mime='text/csv',
            )
            
        with tab2:
            st.dataframe(df_report.astype(str), use_container_width=True)
            # Use st.download_button for direct download
            csv = df_report.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 ดาวน์โหลดข้อมูลการลาเป็น CSV",
                data=csv,
                file_name='leave_report.csv',
                mime='text/csv',
            )

    elif password:
        st.error("❌ รหัสผ่านไม่ถูกต้อง")

