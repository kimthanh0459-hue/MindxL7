import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="🏢 AI Lead Scoring - BĐS", layout="wide")

st.title("🏢 AI Lead Scoring — Bất Động Sản")
st.markdown("Hệ thống chấm điểm khách hàng tiềm năng ngành Bất Động Sản")

# ──────────────────────────────────────────────
# Đọc file Excel
# ──────────────────────────────────────────────
EXCEL_FILE = os.path.join(os.path.dirname(__file__), "leads_for_gsheet.xlsx")

@st.cache_data
def load_data():
    df = pd.read_excel(EXCEL_FILE, engine="openpyxl")
    return df

if not os.path.exists(EXCEL_FILE):
    st.error("❌ Không tìm thấy file `leads_for_gsheet.xlsx`")
    st.stop()

df = load_data()
st.success(f"✅ Đã tải {len(df)} khách hàng từ `leads_for_gsheet.xlsx`")

# ──────────────────────────────────────────────
# Hiển thị toàn bộ dữ liệu
# ──────────────────────────────────────────────
st.subheader("📋 Danh sách khách hàng")
st.dataframe(df, use_container_width=True)

# ──────────────────────────────────────────────
# Thống kê nhanh
# ──────────────────────────────────────────────
st.subheader("📊 Thống kê")
col1, col2, col3 = st.columns(3)
col1.metric("Tổng số leads", len(df))
col2.metric("Số cột dữ liệu", len(df.columns))
col3.metric("Tên các cột", ", ".join(df.columns.tolist()[:3]) + "...")

# ──────────────────────────────────────────────
# Xuất CSV
# ──────────────────────────────────────────────
st.subheader("📥 Tải xuống dữ liệu")
csv = df.to_csv(index=False).encode("utf-8")
st.download_button(
    label="⬇️ Tải về dạng CSV",
    data=csv,
    file_name="leads_export.csv",
    mime="text/csv"
)

