import io
import re
import base64
import json
import urllib.request
import urllib.error
from pathlib import Path
from datetime import date

import pandas as pd
import streamlit as st

st.set_page_config(page_title="International Registration ETL", page_icon="🌏", layout="wide")

st.markdown("""
<style>
.main-title { font-size: 32px; font-weight: 700; margin-bottom: 5px; }
.sub-title { color: #666666; font-size: 17px; margin-bottom: 25px; }
div.stButton > button { width: 100%; height: 48px; font-size: 16px; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">🌏 INTERNATIONAL REGISTRATION ETL SYSTEM</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">ระบบ ETL ข้อมูลผู้สมัครนักศึกษาต่างชาติ</div>', unsafe_allow_html=True)
st.divider()

GITHUB_OWNER = "66011216410-gif"
GITHUB_REPO = "international-registration-etl"
GITHUB_BRANCH = "main"
MASTER_FILE = "master_etl.xlsx"


def get_github_token():
    try:
        return st.secrets.get("GITHUB_TOKEN", "")
    except Exception:
        return ""


def github_request(url, method="GET", data=None, token=None):
    headers = {
        "Accept": "application/vnd.github+json",
        "User-Agent": "international-registration-etl",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    if token:
        headers["Authorization"] = f"Bearer {token}"
    request = urllib.request.Request(url, data=data, headers=headers, method=method)
    with urllib.request.urlopen(request, timeout=30) as response:
        return json.loads(response.read().decode("utf-8"))


def load_master_file():
    token = get_github_token()
    if not token:
        raise RuntimeError("ยังไม่ได้ตั้งค่า GITHUB_TOKEN ใน Streamlit Secrets จึงไม่สามารถบันทึกข้อมูลสะสมลง GitHub ได้")
    url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/contents/{MASTER_FILE}?ref={GITHUB_BRANCH}"
    try:
        result = github_request(url, token=token)
        content = base64.b64decode(result["content"].replace("\n", ""))
        master_df = pd.read_excel(io.BytesIO(content), sheet_name="ETL_Data")
        return master_df, result["sha"]
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None, None
        raise RuntimeError(f"โหลดไฟล์เดิมจาก GitHub ไม่สำเร็จ: HTTP {e.code}")


def save_master_file(df, file_sha=None):
    token = get_github_token()
    if not token:
        raise RuntimeError("ยังไม่ได้ตั้งค่า GITHUB_TOKEN ใน Streamlit Secrets จึงไม่สามารถบันทึกข้อมูลลง GitHub ได้")
    quality = create_data_quality(df)
    excel_data = create_excel(df, quality)
    encoded = base64.b64encode(excel_data).decode("utf-8")
    url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/contents/{MASTER_FILE}"
    payload = {"message": "Update master ETL data", "content": encoded, "branch": GITHUB_BRANCH}
    if file_sha:
        payload["sha"] = file_sha
    try:
        result = github_request(url, method="PUT", data=json.dumps(payload).encode("utf-8"), token=token)
        return result.get("content", {}).get("sha")
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"บันทึกไฟล์ลง GitHub ไม่สำเร็จ: HTTP {e.code} {error_body[:300]}")


def clean_text(series):
    return series.astype("string").str.replace(r"\s+", " ", regex=True).str.strip().replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})


def calculate_age(birthdate):
    if pd.isna(birthdate):
        return pd.NA
    try:
        today = date.today()
        birth = pd.Timestamp(birthdate).date()
        age = today.year - birth.year - ((today.month, today.day) < (birth.month, birth.day))
        return pd.NA if age < 0 else age
    except Exception:
        return pd.NA


def get_age_group(age):
    if pd.isna(age): return "ไม่ทราบ"
    age = int(age)
    if age < 20: return "ต่ำกว่า 20"
    elif age <= 29: return "20-29"
    elif age <= 39: return "30-39"
    elif age <= 49: return "40-49"
    elif age <= 59: return "50-59"
    elif age <= 69: return "60-69"
    elif age <= 79: return "70-79"
    return "80 ปีขึ้นไป"


def get_gender(prefix):
    if pd.isna(prefix): return "Unknown"
    value = re.sub(r"[.\s]+", "", str(prefix).strip().lower())
    if value == "mr": return "Male"
    elif value in ["ms", "mrs", "miss"]: return "Female"
    return "Unknown"


def run_etl(uploaded_file):
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file)
    rows_before = len(df)
    columns_before = len(df.columns)

    df.columns = df.columns.astype(str).str.replace(r"\s+", " ", regex=True).str.strip()

    # ตัดข้อมูล Test No.273 ออกจากไฟล์ใหม่ทันที
    test_rows_removed = 0
    if "No." in df.columns:
        no_clean = pd.to_numeric(df["No."], errors="coerce")
        test_mask = no_clean.eq(273)
        test_rows_removed = int(test_mask.sum())
        df = df.loc[~test_mask].copy()

    text_columns = ["Prefix", "First Name", "Last Name", "Occupation", "Address", "Nationality", "University", "Level", "Faculty", "Major", "Major1"]
    for col in text_columns:
        if col in df.columns:
            df[col] = clean_text(df[col])

    for col in ["Birthdate", "Submitted Date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    if "Birthdate" in df.columns:
        df["Age"] = df["Birthdate"].apply(calculate_age)
        # แก้ปัญหา datetime64[us] cannot be converted to IntegerDtype
        df["Age"] = pd.to_numeric(df["Age"], errors="coerce").round().astype("Int64")
    if "Age" in df.columns:
        df["Age Group"] = df["Age"].apply(get_age_group)
    if "Prefix" in df.columns:
        df["Gender"] = df["Prefix"].apply(get_gender)
    if "Academic Year" in df.columns:
        academic_year = pd.to_numeric(df["Academic Year"], errors="coerce")
        df["Academic Year AD"] = academic_year - 543
    if "Submitted Date" in df.columns:
        df["Month Number"] = df["Submitted Date"].dt.month
        df["Month"] = df["Submitted Date"].dt.month_name()

    if "University" in df.columns:
        university_mapping = {"china weat normal university": "China West Normal University"}
        df["University"] = df["University"].replace(university_mapping).str.title()

    if "Level" in df.columns:
        level_mapping = {"Doctoral Degree": "Doctoral", "Master's Degree": "Master"}
        df["Level"] = df["Level"].replace(level_mapping)

    if "Major" in df.columns:
        df["Major1"] = (
            df["Major"].astype("string")
            .str.replace(r"\b(?:Doctor|Master)\s+of\b", "", regex=True, flags=re.IGNORECASE)
            .str.replace(r"\b(?:Type|Plan)\b\s*[^,;|/]*", "", regex=True, flags=re.IGNORECASE)
            .str.replace(r"\(\s*Revised\b[^)]*\)", "", regex=True, flags=re.IGNORECASE)
            .str.replace(r"\s*,\s*$", "", regex=True)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
            .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
        )

    duplicate_count = int(df.duplicated().sum())
    df = df.drop_duplicates().reset_index(drop=True)
    missing_total = int(df.isna().sum().sum())

    summary = {
        "rows_before": rows_before,
        "rows_after": len(df),
        "columns_before": columns_before,
        "columns_after": len(df.columns),
        "duplicates_removed": duplicate_count,
        "missing_values": missing_total,
        "test_rows_removed": test_rows_removed,
    }
    return df, summary


def create_data_quality(df):
    return pd.DataFrame({"Column": df.columns, "Missing Values": [df[col].isna().sum() for col in df.columns], "Data Type": [str(df[col].dtype) for col in df.columns]})


def create_excel(df, quality=None):
    if quality is None: quality = create_data_quality(df)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="ETL_Data")
        quality.to_excel(writer, index=False, sheet_name="Data_Quality")
    output.seek(0)
    return output.getvalue()


st.subheader("📂 1. อัปโหลดข้อมูลใหม่")
uploaded_file = st.file_uploader("เลือกไฟล์ Excel ที่ต้องการเพิ่มเข้าระบบ", type=["xlsx", "xls"], help="ระบบจะ ETL ข้อมูลใหม่ แล้วนำไปต่อกับ master_etl.xlsx ที่บันทึกไว้ใน GitHub")

if uploaded_file is not None:
    st.success(f"✅ เลือกไฟล์: {uploaded_file.name}")
    if st.button("🚀 เริ่ม ETL และบันทึกข้อมูล", type="primary"):
        try:
            with st.spinner("กำลัง ETL ข้อมูลใหม่..."):
                new_df, new_summary = run_etl(uploaded_file)
            with st.spinner("กำลังโหลดข้อมูลเดิมจาก GitHub..."):
                old_df, old_sha = load_master_file()

            if old_df is not None:
                # ป้องกัน No.273 ที่อาจมีอยู่ใน Master เดิมด้วย
                master_test_removed = 0
                if "No." in old_df.columns:
                    old_no = pd.to_numeric(old_df["No."], errors="coerce")
                    master_test_mask = old_no.eq(273)
                    master_test_removed = int(master_test_mask.sum())
                    old_df = old_df.loc[~master_test_mask].copy()
                combined_df = pd.concat([old_df, new_df], ignore_index=True, sort=False)
                before_dedup = len(combined_df)
                combined_df = combined_df.drop_duplicates().reset_index(drop=True)
                duplicates_removed = before_dedup - len(combined_df)
                old_rows = len(old_df)
            else:
                master_test_removed = 0
                combined_df = new_df.copy()
                duplicates_removed = int(new_df.duplicated().sum())
                old_rows = 0

            with st.spinner("กำลังบันทึกข้อมูลสะสมลง GitHub..."):
                save_master_file(combined_df, old_sha)

            quality = create_data_quality(combined_df)
            summary = {
                "new_rows": len(new_df),
                "old_rows": old_rows,
                "rows_after": len(combined_df),
                "duplicates_removed": duplicates_removed,
                "missing_values": int(combined_df.isna().sum().sum()),
                "test_rows_removed": new_summary.get("test_rows_removed", 0) + master_test_removed,
            }
            st.session_state["etl_df"] = combined_df
            st.session_state["quality"] = quality
            st.session_state["summary"] = summary
            st.session_state["source_name"] = uploaded_file.name
            st.success("🎉 ETL สำเร็จ และบันทึกข้อมูลสะสมลง GitHub แล้ว!")
            if summary["test_rows_removed"] > 0:
                st.info(f"🧪 ตัดข้อมูล Test No.273 ออกแล้ว {summary['test_rows_removed']:,} รายการ")
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาด: {e}")


if "etl_df" in st.session_state:
    df = st.session_state["etl_df"]
    quality = st.session_state["quality"]
    summary = st.session_state["summary"]
    st.divider()
    st.subheader("📊 2. ผลการประมวลผล")
    col1, col2, col3, col4 = st.columns(4)
    with col1: st.metric("ข้อมูลใหม่", f"{summary['new_rows']:,}")
    with col2: st.metric("ข้อมูลเดิม", f"{summary['old_rows']:,}")
    with col3: st.metric("ข้อมูลสะสมทั้งหมด", f"{summary['rows_after']:,}")
    with col4: st.metric("ข้อมูลซ้ำที่ไม่เพิ่ม", f"{summary['duplicates_removed']:,}")
    st.subheader("👁️ 3. ข้อมูลสะสมหลัง ETL")
    st.dataframe(df.head(20), use_container_width=True, height=400)
    st.subheader("🔍 4. ตรวจสอบคุณภาพข้อมูล")
    st.dataframe(quality, use_container_width=True)
    st.subheader("⬇️ 5. ดาวน์โหลดข้อมูลสะสม")
    excel_data = create_excel(df, quality)
    st.download_button(label="⬇️ ดาวน์โหลด master_etl.xlsx", data=excel_data, file_name="master_etl.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    st.success("ข้อมูลสะสมถูกบันทึกไว้ใน GitHub และพร้อมดาวน์โหลด")
