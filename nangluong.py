import streamlit as st
import pandas as pd
from supabase import create_client, Client
from datetime import datetime

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(page_title="Quản lý Lương - TGDV Tuyên Quang", page_icon="📈", layout="wide")

# Kết nối Supabase
SUPABASE_URL = "https://qqzsdxhqrdfvxnlurnyb.supabase.co"
SUPABASE_KEY = "SẾP_ĐIỀN_KEY_CỦA_SẾP_VÀO_ĐÂY"
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# --- 2. CSS GIAO DIỆN (BANNER VIP) ---
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #004B87 0%, #17a2b8 100%);
        padding: 25px; border-radius: 15px; color: white; text-align: center; margin-bottom: 30px;
    }
    .stDataFrame { border: 1px solid #e6e9ef; border-radius: 10px; }
    .chat-bubble { background-color: #f0f2f6; padding: 15px; border-radius: 15px; margin-bottom: 10px; }
</style>
""", unsafe_allow_html=True)

# --- 3. HÀM LẤY DỮ LIỆU ---
@st.cache_data(ttl=600)
def load_data():
    res = supabase.table("theo_doi_luong").select("*").order("stt").execute()
    df = pd.DataFrame(res.data)
    # Đảm bảo cột loai_nang_luong luôn tồn tại
    if 'loai_nang_luong' not in df.columns:
        df['loai_nang_luong'] = 'Thường xuyên'
    return df

try:
    df = load_data()

    # HEADER
    st.markdown('<div class="main-header"><h1>📈 HỆ THỐNG QUẢN LÝ NÂNG LƯƠNG 4.0</h1><p>Ban Tuyên giáo và Dân vận Tỉnh ủy Tuyên Quang</p></div>', unsafe_allow_html=True)

    # --- 4. KHU VỰC TRỢ LÝ CHAT (NHƯ ẢNH SẾP GỬI) ---
    st.subheader("🤖 Trợ lý Thông báo")
    with st.container():
        col_ai, col_text = st.columns([1, 10])
        with col_ai:
            st.title("🤖")
        with col_text:
            st.markdown(f"**Chào Bạn Tuấn!** Hôm nay là ngày {datetime.now().strftime('%d/%m/%Y')}.")
            
            # Lọc danh sách sắp đến hạn
            upcoming = df[df['trang_thai'] == "Sắp đến hạn"]
            prioritized = df[df['loai_nang_luong'] == "Trước thời hạn"]
            
            if not upcoming.empty:
                st.warning(f"🚨 Có **{len(upcoming)}** đồng chí sắp đến hạn nâng lương. Sếp hãy kiểm tra tờ trình nhé!")
            
            if not prioritized.empty:
                st.info(f"🌟 Có **{len(prioritized)}** đồng chí trong danh sách Nâng lương trước thời hạn.")

    st.write("---")

    # --- 5. BỘ LỌC THÔNG MINH ---
    c1, c2, c3 = st.columns([2, 1, 1])
    with c1:
        search = st.text_input("🔍 Tìm cán bộ:", placeholder="Nhập tên cán bộ cần tra cứu...")
    with c2:
        status_filter = st.multiselect("Lọc trạng thái:", options=df['trang_thai'].unique(), default=df['trang_thai'].unique())
    with c3:
        type_filter = st.multiselect("Loại nâng lương:", options=["Thường xuyên", "Trước thời hạn"], default=["Thường xuyên", "Trước thời hạn"])

    # Xử lý lọc dữ liệu
    mask = df['trang_thai'].isin(status_filter) & df['loai_nang_luong'].isin(type_filter)
    if search:
        mask = mask & df['ho_ten'].str.contains(search, case=False)
    
    df_filtered = df[mask].copy()

    # --- 6. HIỂN THỊ DANH SÁCH ---
    st.subheader("📋 Danh sách chi tiết")

    # Hàm định dạng dòng
    def style_rows(row):
        styles = [''] * len(row)
        if row['loai_nang_luong'] == 'Trước thời hạn':
            styles = ['background-color: #fff9c4'] * len(row) # Nền vàng
        return styles

    def color_status(val):
        if val == "Sắp đến hạn": return 'color: #C8102E; font-weight: bold;'
        return 'color: #28a745;'

    # Gắn sao cho người được ưu tiên
    df_filtered['ho_ten_display'] = df_filtered.apply(
        lambda x: f"🌟 {x['ho_ten']}" if x['loai_nang_luong'] == 'Trước thời hạn' else x['ho_ten'], axis=1
    )

    cols_to_show = ['stt', 'ho_ten_display', 'chuc_vu', 'loai_nang_luong', 'bac_luong', 'ngay_du_kien', 'trang_thai']
    
    st.dataframe(
        df_filtered[cols_to_show].style
        .apply(style_rows, axis=1)
        .map(color_status, subset=['trang_thai']),
        use_container_width=True,
        hide_index=True,
        column_config={
            "stt": "STT",
            "ho_ten_display": "HỌ VÀ TÊN",
            "chuc_vu": "CHỨC VỤ",
            "loai_nang_luong": "LOẠI NÂNG LƯƠNG",
            "bac_luong": "BẬC",
            "ngay_du_kien": "NGÀY DỰ KIẾN",
            "trang_thai": "TRẠNG THÁI"
        }
    )

    # Nút bấm xuất dữ liệu
    st.download_button("📥 Tải danh sách (CSV)", df_filtered.to_csv(index=False).encode('utf-8-sig'), "danh_sach_nang_luong.csv", "text/csv")

except Exception as e:
    st.error(f"Lỗi hệ thống: {e}")
    st.info("💡 Sếp hãy kiểm tra lại bảng 'theo_doi_luong' trên Supabase đã đầy đủ dữ liệu chưa nhé.")
