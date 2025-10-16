import io, re, altair as alt, datetime as dt, pandas as pd, numpy as np, streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# =======================
# 🔐 Auth & App Config
# =======================
st.set_page_config(page_title="สคร.9 - ติดตามการลา/ราชการ/สแกน", layout="wide")
creds = service_account.Credentials.from_service_account_info(
    st.secrets["gcp_service_account"], scopes=["https://www.googleapis.com/auth/drive"]
)
ADMIN_PASSWORD = st.secrets.get("admin_password", "admin123")
FOLDER_ID = "1YFJZvs59ahRHmlRrKcQwepWJz6A-4B7d"
FILE_ATTEND = "attendance_report.xlsx"
FILE_LEAVE  = "leave_report.xlsx"
FILE_TRAVEL = "travel_report.xlsx"

service = build("drive", "v3", credentials=creds)

# =======================
# ❗ DataFrame Cleaners & Canonical Name Function
# =======================
def canonical_name(name: str) -> str:
    # Remove spaces, make lowercase, normalize for duplication detection
    return re.sub(r"\s+", "", str(name)).strip().lower()

def clean_df(df, schema):
    df_clean = df.copy()
    for col, t in schema.items():
        if col not in df_clean.columns:
            df_clean[col] = "" if t==str else np.nan
        if t == str:
            df_clean[col] = df_clean[col].astype(str).replace("nan", "").str.strip()
        elif t == dt.date:
            df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce').dt.date
        elif t == int:
            df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0).astype(int)
    df_clean = df_clean.fillna("")
    return df_clean

# =======================
# 🗂️ Drive Helpers (Handle Exception + Backup)
# =======================
def try_api_call(func, *args, **kwargs):
    try:
        return func(*args, **kwargs)
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาด Google API: {e}")
        return None

def get_file_id(filename: str, parent_id=FOLDER_ID):
    q = f"name='{filename}' and '{parent_id}' in parents and trashed=false"
    res = try_api_call(service.files().list, q=q, fields="files(id,name)", supportsAllDrives=True, includeItemsFromAllDrives=True)
    files = res.get("files", []) if res else []
    return files[0]["id"] if files else None

def read_excel_from_drive(filename: str) -> pd.DataFrame:
    file_id = get_file_id(filename)
    if not file_id: return pd.DataFrame()
    req = try_api_call(service.files().get_media, fileId=file_id, supportsAllDrives=True)
    if not req: return pd.DataFrame()
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, req)
    done = False
    while not done:
        try:
            _, done = downloader.next_chunk()
        except Exception as e:
            st.error(f"error reading chunk: {e}")
            return pd.DataFrame()
    fh.seek(0)
    try:
        return pd.read_excel(fh, engine="openpyxl")
    except Exception:
        fh.seek(0)
        return pd.read_excel(fh)

def write_excel_to_drive(filename: str, df: pd.DataFrame):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    media = MediaIoBaseUpload(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    file_id = get_file_id(filename)
    try:
        if file_id:
            service.files().update(fileId=file_id, media_body=media, supportsAllDrives=True).execute()
        else:
            service.files().create(body={"name": filename, "parents": [FOLDER_ID]}, media_body=media, fields="id", supportsAllDrives=True).execute()
    except Exception as e:
        st.error(f"Drive upload failed: {e}")

def backup_excel(original_filename: str, df: pd.DataFrame):
    if df.empty: return
    now = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    backup_filename = f"backup_{now}_{original_filename}"
    write_excel_to_drive(backup_filename, df)
    st.info(f"สำรองข้อมูลไว้แล้ว: {backup_filename}")

# ============= Safe Filename Utility =============
def safe_filename(text):
    # Remove Thai/English unsafe chars for filename
    s = re.sub(r"[^\w]", "_", str(text))
    return re.sub(r"_+", "_", s).strip("_")

# =======================
# 📥 Loading & Cleaning Data
# =======================
df_schema_att = {"ชื่อ-สกุล":str, "วันที่":dt.date, "เวลาเข้า":str, "เวลาออก":str, "หมายเหตุ":str}
df_schema_leave = {"ชื่อ-สกุล":str, "กลุ่มงาน":str, "ประเภทการลา":str, "วันที่เริ่ม":dt.date, "วันที่สิ้นสุด":dt.date, "จำนวนวันลา":int, "last_update":str}
df_schema_travel = {"ชื่อ-สกุล":str, "กลุ่มงาน":str, "กิจกรรม":str, "สถานที่":str, "วันที่เริ่ม":dt.date, "วันที่สิ้นสุด":dt.date, "จำนวนวัน":int, "ผู้ร่วมเดินทาง":str, "ลิงก์เอกสาร":str, "last_update":str}

df_att = clean_df(read_excel_from_drive(FILE_ATTEND), df_schema_att)
df_leave = clean_df(read_excel_from_drive(FILE_LEAVE), df_schema_leave)
df_travel = clean_df(read_excel_from_drive(FILE_TRAVEL), df_schema_travel)

# =======================
# 🎯 Canonical All-Staff List
# =======================
all_names_raw = pd.concat([
    df_leave["ชื่อ-สกุล"], df_travel["ชื่อ-สกุล"], df_att["ชื่อ-สกุล"]
])
all_names_canon = {}
for name in all_names_raw.dropna().unique():
    canon = canonical_name(name)
    if canon and canon not in all_names_canon:
        all_names_canon[canon] = name
all_names = sorted(all_names_canon.values())

staff_groups = sorted([ "กลุ่มโรคติดต่อ", "กลุ่มระบาดวิทยาและตอบโต้ภาวะฉุกเฉินทางสาธารณสุข", "กลุ่มพัฒนาองค์กร", "กลุ่มบริหารทั่วไป", "กลุ่มโรคไม่ติดต่อ", "กลุ่มห้องปฏิบัติการทางการแพทย์", "กลุ่มพัฒนานวัตกรรมและวิจัย", "กลุ่มโรคติดต่อเรื้อรัง", "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.1 จ.ชัยภูมิ", "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.2 จ.บุรีรัมย์", "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.3 จ.สุรินทร์", "ศูนย์ควบคุมโรคติดต่อนำโดยแมลงที่ 9.4 ปากช่อง", "ด่านควบคุมโรคช่องจอม จ.สุรินทร์", "ศูนย์บริการเวชศาสตร์ป้องกัน", "กลุ่มสื่อสารความเสี่ยง", "กลุ่มโรคจากการประกอบอาชีพและสิ่งแวดล้อม" ])
leave_types = ["ลาป่วย", "ลากิจ", "ลาพักผ่อน", "อื่นๆ"]

# =======================
#  Session State Per Form (Not Global)
# =======================
for key in ["submitted_travel","submitted_leave","submitted_attend"]:
    if key not in st.session_state: st.session_state[key]=False

# =======================
# 🧭 ส่วนฟอร์มไปราชการ (แก้ Input/Validate/Backup/Name)
# =======================
st.header("🧭 บันทึกการไปราชการ (หมู่คณะ)")
with st.form("travel_form"):
    group = st.selectbox("กลุ่มงาน", staff_groups, key="travel_group", disabled=st.session_state.submitted_travel)
    activity = st.text_input("กิจกรรม/โครงการ", key="travel_activity", disabled=st.session_state.submitted_travel)
    location = st.text_input("สถานที่", key="travel_location", disabled=st.session_state.submitted_travel)
    start_date = st.date_input("วันที่เริ่ม", dt.date.today(), key="travel_start", disabled=st.session_state.submitted_travel)
    end_date = st.date_input("วันที่สิ้นสุด", dt.date.today(), key="travel_end", disabled=st.session_state.submitted_travel)
    selected_names = st.multiselect("เลือกชื่อเจ้าหน้าที่", all_names, key="travel_names", disabled=st.session_state.submitted_travel)
    new_names_str = st.text_area("เพิ่มชื่อใหม่ (1 คนต่อ 1 บรรทัด)", key="travel_new_names", disabled=st.session_state.submitted_travel)
    submitted_travel = st.form_submit_button("💾 บันทึกข้อมูล", disabled=st.session_state.submitted_travel)

if submitted_travel:
    # Validate: field must not be blank, no duplication, date range valid
    new_names = [x.strip() for x in new_names_str.split("\n") if x.strip()]
    # Use canonical name to check duplication
    all_submit_names = {} # canonical => display name
    for name in selected_names + new_names:
        canon = canonical_name(name)
        if not canon: continue
        if canon not in all_submit_names: all_submit_names[canon]=name
    final_names = list(all_submit_names.values())
    if not final_names:
        st.warning("กรุณาเลือกหรือกรอกชื่อเจ้าหน้าที่อย่างน้อย 1 คน")
        st.session_state.submitted_travel=False
    elif start_date > end_date:
        st.error("'วันที่เริ่ม' ต้องมาก่อนหรือเท่ากับ 'วันที่สิ้นสุด'")
        st.session_state.submitted_travel=False
    elif not group or not activity or not location:
        st.error("กรอกข้อมูล 'กลุ่มงาน', 'กิจกรรม', 'สถานที่' ให้ครบทุกช่อง")
        st.session_state.submitted_travel=False
    else:
        try:
            backup_excel(FILE_TRAVEL, df_travel) # backup ก่อน
            fellow_str = lambda this, lst: ", ".join([n for n in lst if n != this]) if len(lst)>1 else "-"
            num_days = (end_date-start_date).days + 1
            timestamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
            new_records=[]
            for name in final_names:
                new_records.append({
                    "ชื่อ-สกุล": name,
                    "กลุ่มงาน": group,
                    "กิจกรรม": activity,
                    "สถานที่": location,
                    "วันที่เริ่ม": start_date,
                    "วันที่สิ้นสุด": end_date,
                    "จำนวนวัน": num_days,
                    "ผู้ร่วมเดินทาง": fellow_str(name, final_names),
                    "ลิงก์เอกสาร": "-",
                    "last_update": timestamp
                })
            df_travel_new = pd.concat([df_travel, pd.DataFrame(new_records)], ignore_index=True)
            write_excel_to_drive(FILE_TRAVEL, df_travel_new) # ยืนยันเขียน
            backup_excel(FILE_TRAVEL, df_travel_new) # backup หลังบันทึก
            st.success(f"✅ บันทึกเรียบร้อย (เจ้าหน้าที่ {', '.join(final_names)}) ({num_days} วัน)")
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดขณะบันทึกข้อมูล: {e}")
        st.session_state.submitted_travel=False
        st.rerun()

st.dataframe(df_travel.astype(str).sort_values('วันที่เริ่ม', ascending=False), use_container_width=True)

# ฟอร์มการลาและสแกน สามารถใช้ logic และ structure เดียวกันนี้ (validate, backup, session state per form, canonical name, cleaning ก่อน editor/save)
