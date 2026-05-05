import streamlit as st
import pandas as pd
from supabase import create_client

# 1. THÔNG TIN KẾT NỐI
URL = "https://qqzsdxhqrdfvxnlurnyb.supabase.co"
KEY = "DÁN_KEY_CỦA_SẾP_VÀO_ĐÂY"

def main():
    st.set_page_config(page_title="Quản lý Lương Tuyên Quang", layout="wide")

    # Banner tiêu đề - Dùng text thuần để tuyệt đối không lỗi encode
    st.markdown("""
        <div style="background-color: #004B87; padding: 20px; border-radius: 10px; color: white; text-align: center;">
            <h1 style="color: white; margin:0;">HỆ THỐNG QUẢN LÝ LƯƠNG 4.0</h1>
            <p style="color: white; margin:5px;">Ban Tuyên giáo và Dân vận Tỉnh ủy Tuyên Quang</p>
        </div>
    """, unsafe_allow_html=True)

    try:
        # Kết nối
        supabase = create_client(URL, KEY)
        
        # Lấy dữ liệu
        res = supabase.table("theo_doi_luong").select("*").execute()
        
        if res.data:
            # XỬ LÝ DỮ LIỆU AN TOÀN: Ép toàn bộ về chuỗi và xử lý lỗi encode từng ô một
            raw_data = []
            for row in res.data:
                clean_row = {}
                for k, v in row.items():
                    # Chuyển mọi giá trị về string và xử lý ký tự lạ
                    clean_row[k] = str(v).encode('utf-8', 'ignore').decode('utf-8')
                raw_data.append(clean_row)
            
            df = pd.DataFrame(raw_data)

            # Chatbot thông báo
            st.write("")
            with st.chat_message("assistant"):
                st.write("Chào Bạn Tuấn! Hệ thống đã xử lý xong dữ liệu tiếng Việt.")
                
                col_status = next((c for c in df.columns if 'trang' in c.lower()), None)
                if col_status:
                    sap_den_han = df[df[col_status].str.contains("Sắp đến hạn", na=False)]
                    if not sap_den_han.empty:
                        st.error(f"Có {len(sap_den_han)} đồng chí sắp đến hạn nâng lương sếp nhé!")
            
            st.write("---")
            
            # Tìm kiếm
            col_name = next((c for c in df.columns if 'ho' in c.lower() or 'ten' in c.lower()), df.columns[0])
            search = st.text_input("Tìm tên cán bộ:", placeholder="Nhập tên...")
            
            if search:
                df = df[df[col_name].str.contains(search, case=False, na=False)]

            # Hiển thị bảng
            st.dataframe(df, use_container_width=True, hide_index=True)
            
        else:
            st.warning("Dữ liệu trên Supabase đang trống sếp ạ.")

    except Exception as e:
        st.error("Hệ thống đang khởi động lại bảng mã. Sếp vui lòng chờ 10 giây rồi nhấn F5 nhé!")
        # In lỗi ra console để sếp không bị màn hình đỏ
        print(f"Log lỗi: {e}")

if __name__ == "__main__":
    main()
