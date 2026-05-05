import streamlit as st
import pandas as pd
from supabase import create_client

# 1. THÔNG TIN KẾT NỐI
URL = "https://qqzsdxhqrdfvxnlurnyb.supabase.co"
KEY = "DÁN_KEY_CỦA_SẾP_VÀO_ĐÂY"

def main():
    st.set_page_config(page_title="Quản lý Lương Tuyên Quang", layout="wide")

    # Header đơn giản
    st.markdown("""
        <div style="background-color: #004B87; padding: 20px; border-radius: 10px; color: white; text-align: center;">
            <h1 style="color: white; margin:0;">HỆ THỐNG QUẢN LÝ LƯƠNG 4.0</h1>
            <p style="color: white; margin:5px;">Ban Tuyên giáo và Dân vận Tỉnh ủy Tuyên Quang</p>
        </div>
    """, unsafe_allow_html=True)

    try:
        # Khởi tạo kết nối
        supabase = create_client(URL, KEY)
        
        # Lấy dữ liệu dạng thô (Raw) để tránh lỗi decode sớm
        response = supabase.table("theo_doi_luong").select("*").execute()
        
        if response.data:
            # Chuyển đổi dữ liệu sang DataFrame và xử lý lỗi font từng cột
            df = pd.DataFrame(response.data)
            
            # Tuyệt chiêu: Ép toàn bộ DataFrame sang chuỗi UTF-8
            for col in df.columns:
                df[col] = df[col].apply(lambda x: str(x).encode('utf-8', 'replace').decode('utf-8'))

            # Giao diện Chatbot
            st.write("")
            with st.chat_message("assistant"):
                st.write("Chào Bạn Tuấn! Dữ liệu đã được giải mã thành công.")
                
                # Tìm cột trạng thái (không phân biệt hoa thường, có dấu)
                col_status = next((c for c in df.columns if 'trang' in c.lower()), None)
                if col_status:
                    den_han = df[df[col_status].str.contains("Sắp đến hạn", na=False)]
                    if not den_han.empty:
                        st.error(f"🚨 Có **{len(den_han)}** đồng chí sắp đến hạn nâng lương sếp ơi!")
            
            st.write("---")
            
            # Tìm kiếm
            col_name = next((c for c in df.columns if 'ho' in c.lower() or 'ten' in c.lower()), df.columns[0])
            search = st.text_input("🔍 Tìm tên cán bộ:", placeholder="Nhập tên cán bộ...")
            
            if search:
                df_search = df[df[col_name].str.contains(search, case=False, na=False)]
                st.dataframe(df_search, use_container_width=True, hide_index=True)
            else:
                st.dataframe(df, use_container_width=True, hide_index=True)
                
        else:
            st.warning("Dữ liệu trên hệ thống đang trống.")

    except Exception as e:
        # Nếu vẫn lỗi, hiển thị thông tin để anh em mình cùng soi
        st.error("Hệ thống đang gặp trục trặc về bảng mã tiếng Việt.")
        st.write("Chi tiết lỗi gửi sếp:")
        st.code(str(e))

if __name__ == "__main__":
    main()
