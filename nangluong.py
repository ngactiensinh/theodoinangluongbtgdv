import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

st.set_page_config(page_title="Quản lý Lương Ban TG&DV", layout="wide")

st.title("📊 Hệ thống Theo dõi Nâng lương Tuyên Quang")

# 1. Hàm xử lý ngày tháng từ CSV (Định dạng DD/MM/YYYY)
def parse_date(date_str):
    try:
        return pd.to_datetime(date_str, format='%d/%m/%Y')
    except:
        return None

# 2. Tải dữ liệu (Bạn có thể chuyển sang kết nối Supabase sau khi upload xong)
@st.cache_data
def load_data():
    # Thay 'data.csv' bằng tên tệp của bạn hoặc dùng link Supabase
    df = pd.read_csv('BẢNG THEO DÕI NÂNG BẬC LƯƠNG CHO CÁN BỘ, CÔNG CHỨC, NGƯỜI LAO ĐỘNG BAN TG&DV TỈNH UỶ TUYÊN QUANG - Theo dõi nâng lương (1).csv')
    df['NGÀY NÂNG LƯƠNG GẦN NHẤT'] = df['NGÀY NÂNG LƯƠNG GẦN NHẤT'].apply(parse_date)
    df['NGÀY DỰ KIẾN NÂNG'] = df['NGÀY DỰ KIẾN NÂNG LƯƠNG'].apply(parse_date)
    return df

df = load_data()

# 3. Sidebar: Chức năng nâng lương trước thời hạn
st.sidebar.header("⚙️ Cấu hình Nâng lương sớm")
months_early = st.sidebar.selectbox("Số tháng được rút ngắn (Thành tích xuất sắc):", [0, 6, 9, 12])

# 4. Tính toán ngày nâng lương mới dựa trên ưu tiên
def calculate_early_date(original_date, early_months):
    if pd.isnull(original_date): return None
    # Trừ đi số tháng ưu tiên từ ngày dự kiến gốc
    return original_date - pd.DateOffset(months=early_months)

df['NGÀY DỰ KIẾN MỚI'] = df['NGÀY DỰ KIẾN NÂNG'].apply(calculate_early_date, args=(months_early,))

# 5. Hiển thị Dashboard
col1, col2, col3 = st.columns(3)
col1.metric("Tổng số cán bộ", len(df))
col2.metric("Sắp đến hạn (6 tháng)", len(df[df['TRẠNG THÁI NÂNG LƯƠNG'] == 'Sắp đến hạn']))
col3.metric("Ưu tiên rút ngắn", f"{months_early} tháng")

# 6. Bảng dữ liệu chính
st.subheader("📋 Danh sách chi tiết")
st.dataframe(df[['HỌ VÀ TÊN', 'CHỨC VỤ', 'NGÀY NÂNG LƯƠNG GẦN NHẤT', 'NGÀY DỰ KIẾN NÂNG', 'NGÀY DỰ KIẾN MỚI', 'TRẠNG THÁI NÂNG LƯƠNG']])

# 7. Xuất danh sách đề nghị nâng lương sớm
if months_early > 0:
    st.success(f"💡 Danh sách dự kiến nâng lương khi được rút ngắn {months_early} tháng:")
    # Lọc những người có ngày dự kiến mới nằm trong năm hiện tại
    current_year = datetime.now().year
    df_early = df[df['NGÀY DỰ KIẾN MỚI'].dt.year <= current_year]
    st.table(df_early[['HỌ VÀ TÊN', 'CHỨC VỤ', 'NGÀY DỰ KIẾN MỚI']])
