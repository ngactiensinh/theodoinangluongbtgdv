import streamlit as st
import pandas as pd
from supabase import create_client

# 1. CẤU HÌNH KẾT NỐI (Sếp điền thông tin của sếp vào 2 dòng này)
URL = "https://qqzsdxhqrdfvxnlurnyb.supabase.co"
KEY = "SẾP_DÁN_CÁI_KEY_DÀI_DÀI_VÀO_ĐÂY"

# 2. GIAO DIỆN CHÍNH
st.set_page_config(page_title="Quản lý Lương Tuyên Quang", layout="wide")

st.markdown("""
    <div style="background: linear-gradient(135deg, #004B87, #17a2b8); padding: 20px; border-radius: 10px; color: white; text-align: center;">
        <h1>📈 HỆ THỐNG QUẢN LÝ LƯƠNG 4.0</h1>
        <p>Ban Tuyên giáo và Dân vận Tỉnh ủy Tuyên Quang</p>
    </div>
""", unsafe_allow_html=True)

# 3. XỬ LÝ DỮ LIỆU
try:
    # Khởi tạo kết nối ngay tại đây
    client = create_client(URL, KEY)
    
    # Lấy dữ liệu
    res = client.table("theo_doi_luong").select("*").execute()
    df = pd.DataFrame(res.data)

    if not df.empty:
        # Tự tìm tên cột Trạng thái (vì lúc sếp import có thể tên nó khác một chút)
        col_status = next((c for c in df.columns if 'trang' in c.lower() or 'status' in c.lower()), None)
        
        # Giao diện Chatbot thông báo
        st.write("")
        with st.chat_message("assistant"):
            if col_status:
                sap_den_han = df[df[col_status].astype(str).str.contains("Sắp đến hạn", na=False)]
                if not sap_den_han.empty:
                    st.error(f"Chào Bạn Tuấn! Có **{len(sap_den_han)}** đồng chí sắp đến hạn nâng lương. Sếp kiểm tra danh sách nhé!")
                else:
                    st.success("Chào Bạn Tuấn! Hiện tại danh sách nâng lương đều ổn thỏa, chưa có ai đến hạn.")
            else:
                st.info("Chào Bạn Tuấn! Chúc sếp một ngày làm việc hiệu quả.")

        # Thanh tìm kiếm
        st.write("---")
        search = st.text_input("🔍 Tìm tên cán bộ:", placeholder="Gõ tên để tìm nhanh...")
        
        if search:
            # Tự tìm cột Tên để lọc
            col_name = next((c for c in df.columns if 'ho' in c.lower() or 'ten' in c.lower()), df.columns[1])
            df = df[df[col_name].astype(str).str.contains(search, case=False, na=False)]

        # Hiển thị bảng dữ liệu
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.warning("Kết nối thành công nhưng bảng dữ liệu đang trống sếp ơi!")

except Exception as e:
    st.error(f"Lỗi rồi sếp ơi: {e}")
    st.info("Sếp kiểm tra lại cái KEY xem dán đúng chưa nhé!")
