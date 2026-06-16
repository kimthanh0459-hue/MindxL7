import streamlit as st
import pandas as pd
import os
import re
import io

# ─── Page config ──────────────────────────────
st.set_page_config(
    page_title="🏢 AI Lead Scoring BĐS",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── Custom CSS ───────────────────────────────
st.markdown("""
<style>
    .main { background-color: #0f172a; color: #e2e8f0; }
    .block-container { padding-top: 1.5rem; }
    .metric-card {
        background: linear-gradient(135deg, #1e293b, #0f172a);
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 1rem 1.5rem;
        text-align: center;
    }
    .metric-card .num { font-size: 2rem; font-weight: 700; }
    .badge-vip    { background:#7f1d1d; color:#fca5a5; padding:3px 10px; border-radius:20px; font-size:.78rem; font-weight:600; }
    .badge-pot    { background:#14532d; color:#86efac; padding:3px 10px; border-radius:20px; font-size:.78rem; font-weight:600; }
    .badge-watch  { background:#78350f; color:#fcd34d; padding:3px 10px; border-radius:20px; font-size:.78rem; font-weight:600; }
    .badge-trash  { background:#1e293b; color:#94a3b8; padding:3px 10px; border-radius:20px; font-size:.78rem; font-weight:600; border:1px solid #475569; }
    [data-testid="stSidebar"] { background-color: #1e293b; }
    h1, h2, h3 { color: #f1f5f9 !important; }
</style>
""", unsafe_allow_html=True)

# ─── Scoring Engine ───────────────────────────
VIP_KEYWORDS = [
    r"\d+\s*tỷ", "tài chính mạnh", "không thành vấn đề",
    "biệt thự", "penthouse", "shophouse mặt đường",
    "quỹ đất công nghiệp", "sàn văn phòng",
    "quận 1", "ven sông", "vinhomes", "phú mỹ hưng",
    "chủ doanh nghiệp", "nhà đầu tư", "mua sỉ", "mua số lượng lớn",
    "pháp lý chuẩn", "sổ hồng riêng", "gặp trực tiếp chủ đầu tư"
]
TRASH_KEYWORDS = [
    "nhầm số", "không có nhu cầu", "dữ liệu cũ", "nhầm ngành",
    "hỏi giá cho vui", "chưa có ý định mua", "thái độ không hợp tác",
    "bảo hiểm", "vay vốn", "mời chào dịch vụ",
    "thuê bao", "không bắt máy", "không phản hồi zalo"
]

def score_lead(note: str) -> tuple[int, str]:
    if not note or str(note).strip() == "" or str(note).lower() == "nan":
        return 50, "Cần theo dõi"
    note_lower = str(note).lower()

    # Check budget >= 20 tỷ
    budget_match = re.findall(r"(\d+)\s*tỷ", note_lower)
    is_vip_budget = any(int(b) >= 20 for b in budget_match)

    vip_hits = sum(1 for kw in VIP_KEYWORDS if re.search(kw, note_lower))
    trash_hits = sum(1 for kw in TRASH_KEYWORDS if kw in note_lower)

    if is_vip_budget or vip_hits >= 2:
        return min(50 + vip_hits * 10, 100), "VIP"
    elif trash_hits >= 1:
        return max(50 - trash_hits * 15, 0), "Rác"
    elif vip_hits == 1:
        return 65, "Tiềm năng"
    else:
        return 50, "Cần theo dõi"

def classify(score: int) -> str:
    if score >= 80: return "VIP"
    if score >= 50: return "Tiềm năng"
    if score >= 20: return "Cần theo dõi"
    return "Rác"

def badge_html(cat: str) -> str:
    cls = {"VIP": "badge-vip", "Tiềm năng": "badge-pot",
           "Cần theo dõi": "badge-watch", "Rác": "badge-trash"}.get(cat, "badge-watch")
    icon = {"VIP": "🔥", "Tiềm năng": "✅", "Cần theo dõi": "⚠️", "Rác": "❌"}.get(cat, "")
    return f'<span class="{cls}">{icon} {cat}</span>'

# ─── Load data ────────────────────────────────
EXCEL_FILE = os.path.join(os.path.dirname(__file__), "leads_for_gsheet.xlsx")

@st.cache_data
def load_data():
    df = pd.read_excel(EXCEL_FILE, engine="openpyxl")
    return df

if not os.path.exists(EXCEL_FILE):
    st.error("❌ Không tìm thấy file `leads_for_gsheet.xlsx`")
    st.stop()

# ─── Session state ────────────────────────────
if "df_scored" not in st.session_state:
    raw = load_data()
    note_col = next((c for c in raw.columns if "ghi" in c.lower() or "note" in c.lower() or "nhu" in c.lower()), None)
    if note_col is None:
        note_col = raw.columns[-1]
    raw["_note_col"] = note_col
    scores, cats = zip(*[score_lead(str(r)) for r in raw[note_col]])
    raw["Điểm AI"] = list(scores)
    raw["Phân loại"] = list(cats)
    raw["Phân loại (Review)"] = list(cats)
    raw["Đã review"] = False
    st.session_state.df_scored = raw
    st.session_state.note_col = note_col

df = st.session_state.df_scored

# ─── Sidebar ──────────────────────────────────
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/real-estate.png", width=80)
    st.title("🏢 Lead Scoring")
    st.markdown("**AI Bất Động Sản**")
    st.divider()

    st.subheader("🔍 Bộ lọc")
    filter_cat = st.multiselect(
        "Phân loại",
        ["VIP", "Tiềm năng", "Cần theo dõi", "Rác"],
        default=["VIP", "Tiềm năng", "Cần theo dõi", "Rác"]
    )
    score_min, score_max = st.slider("Khoảng điểm", 0, 100, (0, 100))
    show_reviewed = st.checkbox("Chỉ hiện chưa review", value=False)

    st.divider()
    if st.button("🔄 Chấm điểm lại"):
        del st.session_state["df_scored"]
        st.rerun()

# ─── Filter ───────────────────────────────────
mask = (
    df["Phân loại"].isin(filter_cat) &
    (df["Điểm AI"] >= score_min) &
    (df["Điểm AI"] <= score_max)
)
if show_reviewed:
    mask &= ~df["Đã review"]
df_view = df[mask].copy()

# ─── Header ───────────────────────────────────
st.title("🏢 AI Lead Scoring — Bất Động Sản")
st.caption(f"Nguồn dữ liệu: `leads_for_gsheet.xlsx` | {len(df)} leads | {df['Đã review'].sum()} đã review")

# ─── KPI Cards ────────────────────────────────
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("📋 Tổng leads", len(df))
c2.metric("🔥 VIP",        len(df[df["Phân loại"] == "VIP"]))
c3.metric("✅ Tiềm năng",  len(df[df["Phân loại"] == "Tiềm năng"]))
c4.metric("⚠️ Theo dõi",  len(df[df["Phân loại"] == "Cần theo dõi"]))
c5.metric("❌ Rác",        len(df[df["Phân loại"] == "Rác"]))

st.divider()

# ─── Charts ───────────────────────────────────
col_chart1, col_chart2 = st.columns(2)
with col_chart1:
    st.subheader("📊 Phân bố phân loại")
    cat_counts = df["Phân loại"].value_counts().reset_index()
    cat_counts.columns = ["Phân loại", "Số lượng"]
    st.bar_chart(cat_counts.set_index("Phân loại"))

with col_chart2:
    st.subheader("📈 Phân bố điểm AI")
    hist_data = pd.cut(df["Điểm AI"], bins=[0,20,50,80,100], labels=["0-20","20-50","50-80","80-100"])
    st.bar_chart(hist_data.value_counts().sort_index())

st.divider()

# ─── Human-in-the-loop Table ──────────────────
st.subheader("👁️ Kiểm duyệt & Chỉnh sửa")
st.caption("Chỉnh sửa điểm và phân loại, sau đó nhấn **💾 Lưu thay đổi**.")

edited = st.data_editor(
    df_view.reset_index(drop=True),
    column_config={
        "Điểm AI": st.column_config.NumberColumn("Điểm AI", min_value=0, max_value=100, step=1),
        "Phân loại (Review)": st.column_config.SelectboxColumn(
            "Phân loại (Review)",
            options=["VIP", "Tiềm năng", "Cần theo dõi", "Rác"]
        ),
        "Đã review": st.column_config.CheckboxColumn("Đã review"),
    },
    use_container_width=True,
    num_rows="fixed",
    key="editor"
)

if st.button("💾 Lưu thay đổi"):
    for idx, row in edited.iterrows():
        orig_idx = df_view.index[idx]
        st.session_state.df_scored.at[orig_idx, "Điểm AI"] = row["Điểm AI"]
        st.session_state.df_scored.at[orig_idx, "Phân loại (Review)"] = row["Phân loại (Review)"]
        st.session_state.df_scored.at[orig_idx, "Đã review"] = row["Đã review"]
    st.success("✅ Đã lưu thay đổi!")
    st.rerun()

st.divider()

# ─── Export ───────────────────────────────────
st.subheader("📥 Xuất kết quả")
col_ex1, col_ex2 = st.columns(2)

with col_ex1:
    csv = st.session_state.df_scored.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Tải về CSV", csv, "lead_scoring_result.csv", "text/csv")

with col_ex2:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        st.session_state.df_scored.to_excel(writer, index=False, sheet_name="Lead Scoring")
    st.download_button("📊 Tải về Excel", buf.getvalue(),
                       "lead_scoring_result.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

