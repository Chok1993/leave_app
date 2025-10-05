import streamlit as st
import pandas as pd
import datetime as dt
import altair as alt
import matplotlib.pyplot as plt
from fpdf import FPDF
import tempfile
import base64
import io

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

# ===== รหัสผ่านแอดมิน =====
ADMIN_PASSWORD = "DDC9admin"

# ===== เมนูหลัก =====
st.title("📋 โปรแกรมบันทึกข้อมูล (สคร.9)")
main_menu = st.sidebar.radio("เลือกหน้าเมนู", ["📊 Dashboard รวม", "🧭 การไปราชการ", "🕒 การลา", "🛠️ ผู้ดูแลระบบ (Admin)"])

# ===============================================================
# ====================== DASHBOARD รวม ===========================
# ===============================================================
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
        avg_leave = round(df_report["จำนวนวันลา"].mean(), 2)
        c1, c2 = st.columns(2)
        c1.metric("จำนวนบุคลากรที่ลา", f"{unique_persons} คน")
        c2.metric("เฉลี่ยวันลาต่อคน", f"{avg_leave} วัน")

        st.markdown("### 📊 จำนวนวันลาตามประเภท")
        leave_by_type = df_report.groupby("ประเภทการลา")["จำนวนวันลา"].sum().reset_index()
        st.bar_chart(leave_by_type.set_index("ประเภทการลา"))

        st.markdown("### 🏢 จำนวนวันลาตามกลุ่มงาน")
        leave_by_group = df_report.groupby("กลุ่มงาน")["จำนวนวันลา"].sum().reset_index()
        st.bar_chart(leave_by_group.set_index("กลุ่มงาน"))

        st.markdown("### 📋 ตารางข้อมูลการลา (ดิบ)")
        st.dataframe(df_report.astype(str), use_container_width=True)

# ===============================================================
# =================== ฟอร์มบันทึกการไปราชการ ===================
# ===============================================================
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
        elif not data["กิจกรรม"]:
            st.error("⚠️ กรุณากรอกชื่อกิจกรรม")
        elif int(num_people) > 0 and len(companions) < int(num_people):
            st.error("⚠️ กรุณากรอกชื่อผู้ร่วมเดินทางให้ครบ")
        else:
            df_scan = df_scan.dropna(how='all')
            df_scan = pd.concat([df_scan, pd.DataFrame([data])], ignore_index=True)
            df_scan.to_excel(FILE_SCAN, index=False)
            st.success("✅ บันทึกข้อมูลเรียบร้อย")

    if not df_scan.empty:
        st.markdown("## 📊 สรุปข้อมูลการไปราชการ")
        st.write(f"📋 ทั้งหมด {len(df_scan)} รายการ รวม {df_scan['จำนวนวัน'].sum()} วัน")
        st.dataframe(df_scan.astype(str), use_container_width=True)

# ===============================================================
# =================== ฟอร์มบันทึกการลา ==========================
# ===============================================================
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

# ------------------ ADMIN ------------------
elif menu == "👩‍💼 Admin":
    import matplotlib.pyplot as plt
    from fpdf import FPDF
    import tempfile

    st.header("🔐 เข้าสู่ระบบผู้ดูแล")
    password = st.text_input("กรอกรหัสผ่าน", type="password")

    if password == ADMIN_PASSWORD:
        st.success("✅ เข้าสู่ระบบสำเร็จ")

        tab1, tab2, tab3 = st.tabs(["🧭 ข้อมูลไปราชการ", "🕒 ข้อมูลการลา", "📈 Dashboard กลุ่มงาน"])

        # ========== แท็บ 1: ไปราชการ ==========
        with tab1:
            st.markdown("### 🧭 ข้อมูลการไปราชการทั้งหมด")
            st.dataframe(df_scan.astype(str), use_container_width=True)

        # ========== แท็บ 2: การลา ==========
        with tab2:
            st.markdown("### 🕒 ข้อมูลการลาทั้งหมด")
            st.dataframe(df_report.astype(str), use_container_width=True)

        # ========== แท็บ 3: Dashboard กลุ่มงาน ==========
        with tab3:
            st.markdown("### 📈 Dashboard สรุปข้อมูลตามกลุ่มงาน")

            # ====== ตัวกรองปีและเดือน ======
            this_year = dt.date.today().year + 543
            year_choice = st.selectbox("เลือกปี พ.ศ.", list(range(this_year - 3, this_year + 1)), index=3)
            month_choice = st.selectbox("เลือกเดือน", list(range(1, 13)), format_func=lambda x: f"เดือน {x}")

            # ====== ฟังก์ชันกรองข้อมูล ======
            def filter_data_by_month(df, start_col, end_col):
                df = df.copy()
                df[start_col] = pd.to_datetime(df[start_col], errors="coerce")
                df[end_col] = pd.to_datetime(df[end_col], errors="coerce")
                df["ปี"] = df[start_col].dt.year + 543
                df["เดือน"] = df[start_col].dt.month
                return df[(df["ปี"] == year_choice) & (df["เดือน"] == month_choice)]

            df_scan_filtered = filter_data_by_month(df_scan, "วันที่เริ่ม", "วันที่สิ้นสุด") if not df_scan.empty else pd.DataFrame()
            df_report_filtered = filter_data_by_month(df_report, "วันที่เริ่ม", "วันที่สิ้นสุด") if not df_report.empty else pd.DataFrame()

            col1, col2 = st.columns(2)

            # --- กราฟไปราชการ ---
            fig1, fig2, fig3 = None, None, None
            if not df_scan_filtered.empty:
                travel_group = df_scan_filtered.groupby("กลุ่มงาน")["จำนวนวัน"].sum().sort_values(ascending=False).head(5)
                col1.subheader(f"🧭 Top 5 กลุ่มงานที่ไปราชการมากที่สุด ({month_choice}/{year_choice})")
                col1.bar_chart(travel_group)
                fig1 = travel_group.plot(kind="bar", color="skyblue", figsize=(5, 3), title="Top 5 กลุ่มงานไปราชการ").get_figure()
            else:
                col1.info("ไม่มีข้อมูลการไปราชการในเดือนที่เลือก")

            # --- กราฟการลา ---
            if not df_report_filtered.empty:
                leave_group = df_report_filtered.groupby("กลุ่มงาน")["จำนวนวันลา"].sum().sort_values(ascending=False).head(5)
                col2.subheader(f"🕒 Top 5 กลุ่มงานที่ลามากที่สุด ({month_choice}/{year_choice})")
                col2.bar_chart(leave_group)
                fig2 = leave_group.plot(kind="bar", color="salmon", figsize=(5, 3), title="Top 5 กลุ่มงานลา").get_figure()
            else:
                col2.info("ไม่มีข้อมูลการลาในเดือนที่เลือก")

            # ===== กราฟวงกลมประเภทการลา =====
            st.markdown("### 🥧 สัดส่วนประเภทการลา")
            if not df_report_filtered.empty and "ประเภทการลา" in df_report_filtered.columns:
                leave_type = df_report_filtered.groupby("ประเภทการลา")["จำนวนวันลา"].sum()
                if not leave_type.empty:
                    fig3, ax = plt.subplots(figsize=(5, 5))
                    ax.pie(leave_type, labels=leave_type.index, autopct="%1.1f%%", startangle=90)
                    ax.set_title("สัดส่วนประเภทการลา")
                    st.pyplot(fig3)
                else:
                    st.info("ไม่มีข้อมูลประเภทการลาในเดือนนี้")
            else:
                st.info("ไม่มีข้อมูลการลาในเดือนที่เลือก")

            # ===== ตารางสรุปรวม =====
            st.markdown("### 📋 ตารางสรุปผลรวมตามกลุ่มงาน")
            summary = None
            if not df_scan_filtered.empty or not df_report_filtered.empty:
                travel_sum = (
                    df_scan_filtered.groupby("กลุ่มงาน")["จำนวนวัน"]
                    .sum()
                    .reset_index()
                    .rename(columns={"จำนวนวัน": "รวมวันไปราชการ"})
                )
                leave_sum = (
                    df_report_filtered.groupby("กลุ่มงาน")["จำนวนวันลา"]
                    .sum()
                    .reset_index()
                    .rename(columns={"จำนวนวันลา": "รวมวันลา"})
                )
                summary = pd.merge(travel_sum, leave_sum, on="กลุ่มงาน", how="outer").fillna(0)
                summary["รวมทั้งหมด"] = summary["รวมวันไปราชการ"] + summary["รวมวันลา"]
                st.dataframe(summary.sort_values("รวมทั้งหมด", ascending=False), use_container_width=True)
            else:
                st.info("ไม่มีข้อมูลในเดือนที่เลือกสำหรับการสรุป")

            # ===== ปุ่มสร้าง PDF =====
            st.markdown("### 🖨️ พิมพ์รายงานสรุป (PDF)")
            if st.button("📄 สร้างรายงาน PDF"):
                pdf = FPDF()
                pdf.add_page()
                pdf.add_font('THSarabun', '', 'THSarabunNew.ttf', uni=True)
                pdf.set_font('THSarabun', '', 16)
                pdf.cell(0, 10, f"รายงานสรุปผลการลาและไปราชการ เดือน {month_choice} ปี {year_choice}", ln=True, align="C")
                pdf.ln(10)

                # เก็บไฟล์กราฟชั่วคราว
                temp_dir = tempfile.gettempdir()
                if fig1:
                    path1 = f"{temp_dir}/travel_chart.png"
                    fig1.savefig(path1)
                    pdf.image(path1, w=170)
                if fig2:
                    path2 = f"{temp_dir}/leave_chart.png"
                    fig2.savefig(path2)
                    pdf.image(path2, w=170)
                if fig3:
                    path3 = f"{temp_dir}/pie_chart.png"
                    fig3.savefig(path3)
                    pdf.image(path3, w=150)

                if summary is not None:
                    pdf.ln(10)
                    pdf.set_font('THSarabun', '', 14)
                    pdf.cell(0, 10, "ตารางสรุปผลรวม (วัน)", ln=True)
                    pdf.ln(5)
                    for _, row in summary.iterrows():
                        pdf.cell(0, 8, f"{row['กลุ่มงาน']} - ไปราชการ {int(row['รวมวันไปราชการ'])} / ลา {int(row['รวมวันลา'])} / รวม {int(row['รวมทั้งหมด'])}", ln=True)

                # ส่งออก PDF
                pdf_output = f"{temp_dir}/summary_{year_choice}_{month_choice}.pdf"
                pdf.output(pdf_output)
                with open(pdf_output, "rb") as f:
                    st.download_button(
                        label="📥 ดาวน์โหลดรายงาน PDF",
                        data=f,
                        file_name=f"รายงานสรุป_สคร9_{year_choice}_{month_choice}.pdf",
                        mime="application/pdf"
                    )

            # ===== ปุ่มดาวน์โหลดรายงาน Excel =====
            st.markdown("### 📥 ดาวน์โหลดรายงานรวมทั้งหมด (Excel)")
            def to_excel(download_scan, download_leave):
                from io import BytesIO
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    download_scan.to_excel(writer, sheet_name="การไปราชการ", index=False)
                    download_leave.to_excel(writer, sheet_name="การลา", index=False)
                return output.getvalue()

            excel_data = to_excel(df_scan, df_report)
            st.download_button(
                label="📥 ดาวน์โหลดรายงานสรุปทั้งหมด (Excel)",
                data=excel_data,
                file_name=f"สรุปรายงาน_สคร9_{dt.date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    elif password:
        st.error("❌ รหัสผ่านไม่ถูกต้อง กรุณาลองใหม่")
