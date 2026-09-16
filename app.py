import io
import re
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


def clean_text(series):
    return (series.astype("string").str.replace(r"\s+", " ", regex=True).str.strip().replace({"": pd.NA, "nan": pd.NA, "None": pd.NA}))


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
    """จัดกลุ่มอายุเป็นช่วงละ 10 ปี"""
    if pd.isna(age):
        return "ไม่ทราบ"
    age = int(age)
    if age < 20:
        return "ต่ำกว่า 20"
    elif age <= 29:
        return "20-29"
    elif age <= 39:
        return "30-39"
    elif age <= 49:
        return "40-49"
    elif age <= 59:
        return "50-59"
    elif age <= 69:
        return "60-69"
    elif age <= 79:
        return "70-79"
    else:
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

    text_columns = ["Prefix", "First Name", "Last Name", "Occupation", "Address", "Nationality", "University", "Level", "Faculty", "Major", "Major1"]
    for col in text_columns:
        if col in df.columns:
            df[col] = clean_text(df[col])

    for col in ["Birthdate", "Submitted Date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    if "Birthdate" in df.columns:
        df["Age"] = df["Birthdate"].apply(calculate_age).astype("Int64")
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

    # Create Major1 from Major
    # Remove Doctor of / Master of, Type / Plan and any parenthetical Revised text.
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

    quality = pd.DataFrame({
        "Column": df.columns,
        "Missing Values": [df[col].isna().sum() for col in df.columns],
        "Data Type": [str(df[col].dtype) for col in df.columns]
    })

    summary = {
        "rows_before": rows_before,
        "rows_after": len(df),
        "columns_before": columns_before,
        "columns_after": len(df.columns),
        "duplicates_removed": duplicate_count,
        "missing_values": missing_total
    }
    return df, quality, summary


def create_excel(df, quality):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="ETL_Data")
        quality.to_excel(writer, index=False, sheet_name="Data_Quality")
    output.seek(0)
    return output.getvalue()


st.subheader("📂 1. อัปโหลดข้อมูล")
uploaded_file = st.file_uploader("เลือกไฟล์ Excel ที่ต้องการประมวลผล", type=["xlsx", "xls"], help="รองรับไฟล์ Excel .xlsx และ .xls")

if uploaded_file is not None:
    st.success(f"✅ เลือกไฟล์: {uploaded_file.name}")
    if st.button("🚀 เริ่ม ETL", type="primary"):
        try:
            with st.spinner("กำลัง Extract → Transform → Validate..."):
                df, quality, summary = run_etl(uploaded_file)
            st.session_state["etl_df"] = df
            st.session_state["quality"] = quality
            st.session_state["summary"] = summary
            st.session_state["source_name"] = uploaded_file.name
            st.success("🎉 ETL Process สำเร็จ!")
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาด: {e}")

if "etl_df" in st.session_state:
    df = st.session_state["etl_df"]
    quality = st.session_state["quality"]
    summary = st.session_state["summary"]

    st.divider()
    st.subheader("📊 2. ผลการประมวลผล")
    col1, col2, col3, col4 = st.columns(4)
    with col1: st.metric("ข้อมูลก่อน ETL", f"{summary['rows_before']:,}")
    with col2: st.metric("ข้อมูลหลัง ETL", f"{summary['rows_after']:,}")
    with col3: st.metric("ข้อมูลซ้ำที่ลบ", f"{summary['duplicates_removed']:,}")
    with col4: st.metric("Missing Values", f"{summary['missing_values']:,}")

    st.subheader("👁️ 3. ตัวอย่างข้อมูลหลัง ETL")
    st.dataframe(df.head(20), use_container_width=True, height=400)
    st.subheader("🔍 4. ตรวจสอบคุณภาพข้อมูล")
    st.dataframe(quality, use_container_width=True)
    st.subheader("⬇️ 5. ดาวน์โหลดข้อมูล")
    excel_data = create_excel(df, quality)
    source_name = st.session_state.get("source_name", "international_registration.xlsx")
    output_name = Path(source_name).stem + "_ETL.xlsx"
    st.download_button(label="⬇️ ดาวน์โหลดข้อมูล ETL (.xlsx)", data=excel_data, file_name=output_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    st.success("พร้อมดาวน์โหลดข้อมูลที่ผ่านกระบวนการ ETL แล้ว")
else:
    st.info("📌 กรุณาอัปโหลดไฟล์ Excel แล้วกด 🚀 เริ่ม ETL")
