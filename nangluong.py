import streamlit as st
import pandas as pd
from supabase import create_client

# 1. THÔNG TIN KẾT NỐI
# Sếp nhớ thay URL và KEY chuẩn của sếp vào đây nhé
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
        # Khởi tạo kết nối
        supabase = create_client(URL, KEY)
        
        # Lấy dữ liệu từ bảng
        res = supabase.table("theo_doi_luong").select("*").execute()
        df = pd.DataFrame(res.data)

        if not df.empty:
            # Tìm cột trạng thái thông minh
            col_status = next((c for c in df.columns if 'trang' in c.lower()), None)
            
            st.write("")
            with st.chat_message("assistant"):
                if col_status:
                    sap_den_han = df[df[col_status].astype(str).str.contains("Sắp đến hạn", na=False)]
                    if not sap_den_han.empty:
                        st.error(f"Chào Bạn Tuấn! Có **{len(sap_den_han)}** đồng chí sắp đến hạn nâng lương. Sếp kiểm tra nhé!")
                    else:
                        st.success("Chào Bạn Tuấn! Hiện tại danh sách nâng lương đều ổn, chưa có ai đến hạn.")
                else:
                    st.info("Chào Bạn Tuấn! Chúc sếp một ngày làm việc hiệu quả.")

            st.write("---")
            # Tìm cột tên thông minh để lọc
            col_name = next((c for c in df.columns if 'ho' in c.lower() or 'ten' in c.lower()), df.columns[0])
            search = st.text_input("🔍 Tìm tên cán bộ:", placeholder="Gõ tên để tìm nhanh...")
            
            if search:
                df = df[df[col_name].astype(str).str.contains(search, case=False, na=False)]

            # Hiển thị bảng
            st.dataframe(df, use_container_width=True, hide_index=True)
        else:
            st.warning("Kết nối OK nhưng bảng 'theo_doi_luong' chưa có dữ liệu sếp ơi!")

    except Exception as e:
        st.error(f"Lỗi rồi sếp ơi: {e}")

if __name__ == "__main__":
    main()
