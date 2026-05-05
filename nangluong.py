import streamlit as st
import pandas as pd
from supabase import create_client
import sys

# Ép hệ thống dùng UTF-8 để không bị lỗi 'ascii' codec
import importlib
importlib.reload(sys)

# 1. THÔNG TIN KẾT NỐI
URL = "https://qqzsdxhqrdfvxnlurnyb.supabase.co"
KEY = "DÁN_KEY_CỦA_SẾP_VÀO_ĐÂY"

def main():
    st.set_page_config(page_title="Quản lý Lương Tuyên Quang", layout="wide")

    # Banner tiêu đề
    st.markdown("""
        <div style="background: linear-gradient(135deg, #004B87, #17a2b8); padding: 20px; border-radius: 10px; color: white; text-align: center;">
            <h1>📈 HỆ THỐNG QUẢN LÝ LƯƠNG 4.0</h1>
            <p>Ban Tuyên giáo và Dân vận Tỉnh ủy Tuyên Quang</p>
        </div>
    """, unsafe_allow_html=True)

    try:
        # Khởi tạo kết nối an toàn
        supabase = create_client(URL, KEY)
        
        # Lấy dữ liệu và ép kiểu sang string ngay từ đầu để tránh lỗi encode
        res = supabase.table("theo_doi_luong").select("*").execute()
        
        if res.data:
            df = pd.DataFrame(res.data)
            # Chuyển toàn bộ dữ liệu sang dạng string an toàn
            df = df.astype(str)

            # Tìm cột trạng thái và tên thông minh
            col_status = next((c for c in df.columns if 'trang' in c.lower()), None)
            col_name = next((c for c in df.columns if 'ho' in c.lower() or 'ten' in c.lower()), df.columns[0])
            
            st.write("")
            with st.chat_message("assistant"):
                st.write(f"**Chào Bạn Tuấn!** Chúc sếp một ngày làm việc hiệu quả.")
                if col_status:
                    # Lọc an toàn
                    sap_den_han = df[df[col_status].str.contains("Sắp đến hạn", na=False)]
                    if not sap_den_han.empty:
                        st.error(f"🚨 Có **{len(sap_den_han)}** đồng chí sắp đến hạn nâng lương!")
                    else:
                        st.success("✅ Hiện tại chưa có ai đến hạn nâng lương sếp ạ.")

            st.write("---")
            search = st.text_input("🔍 Tìm tên cán bộ:", placeholder="Gõ tên để tìm nhanh...")
            
            if search:
                df = df[df[col_name].str.contains(search, case=False, na=False)]

            st.dataframe(df, use_container_width=True, hide_index=True)
        else:
            st.warning("Kết nối thành công nhưng chưa có dữ liệu trong bảng.")

    except Exception as e:
        # Hiển thị lỗi chi tiết nhưng không làm văng app
        st.error(f"Đã có lỗi xảy ra: {str(e)}")

if __name__ == "__main__":
    main()
