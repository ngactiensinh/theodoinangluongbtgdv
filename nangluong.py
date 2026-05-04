import streamlit as st
import pandas as pd
from supabase import create_client, Client
from datetime import datetime

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(page_title="Quản lý Lương - TGDV Tuyên Quang", page_icon="📈", layout="wide")

# Kết nối Supabase
SUPABASE_URL = "https://qqzsdxhqrdfvxnlurnyb.supabase.co"
SUPABASE_KEY = "SẾP_ĐIỀN_KEY_VÀO_ĐÂY"
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# --- 2. CSS GIAO DIỆN ---
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #004B87 0%, #17a2b8 100%);
        padding: 20px; border-radius: 10px; color: white; text-align: center; margin-bottom: 20px;
    }
    .stDataFrame { border: 1px solid #e6e9ef; border-radius: 10px; }
</style>
""", unsafe_allow_html=True)

# --- 3. HÀM DÒ TÌM CỘT THÔNG MINH ---
def get_col_name(df, keywords):
    for col in df.columns:
        for key in keywords:
            if key.lower() in col.lower():
                return col
    return None

# --- 4. TẢI DỮ LIỆU ---
@st.cache_data(ttl=60)
def load_data():
    try:
        res = supabase.table("theo_doi_luong").select("*").execute()
        return pd.DataFrame(res.data)
    except Exception:
        return pd.DataFrame()

# --- CHƯƠNG TRÌNH CHÍNH ---
df_raw = load_data()

if not df_raw.empty:
    # Xác định các cột quan trọng
    col_ten = get_col_name(df_raw, ['ho_ten', 'họ và tên', 'ten'])
    col_trang_thai = get_col_name(df_raw, ['trang_thai', 'trạng thái'])
    col_ngay = get_col_name(df_raw, ['ngay_du_kien', 'dự kiến'])
    col_chuc_vu = get_col_name(df_raw, ['chuc_vu', 'chức vụ'])
    col_loai = get_col_name(df_raw, ['loai_nang_luong', 'loại']) or 'loai_ao'
    
    if col_loai == 'loai_ao': 
        df_raw[col_loai] = 'Thường xuyên'

    # HEADER
    st.markdown('<div class="main-header"><h1>📈 HỆ THỐNG QUẢN LÝ LƯƠNG 4.0</h1><p>Ban Tuyên giáo và Dân vận Tỉnh ủy Tuyên Quang</p></div>', unsafe_allow_html=True)

    # --- 5. GIAO DIỆN TRỢ LÝ ---
    with st.chat_message("assistant"):
        st.write(f"**Chào Bạn Tuấn!** Hôm nay là ngày {datetime.now().strftime('%d/%m/%Y')}.")
        if col_trang_thai:
            sap_den_han = df_raw[df_raw[col_trang_thai].astype(str).str.contains("Sắp đến hạn", na=False)]
            if not sap_den_han.empty:
                st.error(f"🚨 Có **{len(sap_den_han)}** đồng chí sắp đến hạn nâng lương!")
            else:
                st.success("✅ Hiện tại không có hồ sơ nào sắp đến hạn.")

    # --- 6. BỘ LỌC VÀ HIỂN THỊ ---
    st.write("---")
    search = st.text_input("🔍 Tìm kiếm cán bộ:", placeholder="Nhập tên để tìm...")
    
    df_filtered = df_raw.copy()
    if search and col_ten:
        df_filtered = df_filtered[df_filtered[col_ten].astype(str).str.contains(search, case=False, na=False)]

    def style_dataframe(row):
        style = [''] * len(row)
        if col_loai in row and str(row[col_loai]) == 'Trước thời hạn':
            style = ['background-color: #fff9c4'] * len(row)
        return style

    # Lấy danh sách cột có thực để hiển thị
    display_cols = [c for c in [get_col_name(df_raw, ['stt']), col_ten, col_chuc_vu, col_ngay, col_trang_thai] if c]

    if not df_filtered.empty:
        st.dataframe(
            df_filtered[display_cols].style.apply(style_dataframe, axis=1),
            use_container_width=True,
            hide_index=True
        )
else:
    st.warning("Đang kết nối dữ liệu... Sếp hãy kiểm tra lại bảng 'theo_doi_luong' trên Supabase nhé.")
