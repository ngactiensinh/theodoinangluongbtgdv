import streamlit as st
import pandas as pd
from supabase import create_client
import os

# --- PHẦN FIX LỖI TIẾNG VIỆT (ASCII) ---
import sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

# 1. THÔNG TIN KẾT NỐI
URL = "https://qqzsdxhqrdfvxnlurnyb.supabase.co"
KEY = "DÁN_KEY_CỦA_SẾP_VÀO_ĐÂY"

def main():
    st.set_page_config(page_title="Quản lý Lương Tuyên Quang", layout="wide")

    # Banner tiêu đề (Dùng Markdown đơn giản để tránh lỗi hiển thị)
    st.markdown(f"""
        <div style="background-color: #004B87; padding: 20px; border-radius: 10px; color: white; text-align: center;">
            <h1 style="color: white;">📈 HỆ THỐNG QUẢN LÝ LƯƠNG 4.0</h1>
            <p style="color: white;">Ban Tuyên giáo và Dân vận Tỉnh ủy Tuyên Quang</p>
        </div>
    """, unsafe_allow_html=True)

    try:
        # Khởi tạo kết nối
        supabase = create_client(URL, KEY)
        
        # Lấy dữ liệu
        res = supabase.table("theo_doi_luong").select("*").execute()
        
        if res.data:
            # Ép kiểu dữ liệu về chuỗi ngay lập tức với mã hóa UTF-8
            df = pd.DataFrame(res.data)
            
            # Giao diện thông báo
            st.write("")
            with st.chat_message("assistant"):
                st.write("Chào Bạn Tuấn! Hệ thống đã sẵn sàng phục vụ sếp.")
                
                # Tìm cột trạng thái thông minh
                col_status = next((c for c in df.columns if 'trang' in c.lower()), None)
                if col_status:
                    df[col_status] = df[col_status].astype(str)
                    sap_den_han = df[df[col_status].str.contains("Sắp đến hạn", na=False)]
                    if not sap_den_han.empty:
                        st.error(f"🚨 Có **{len(sap_den_han)}** đồng chí sắp đến hạn nâng lương!")
            
            st.write("---")
            
            # Tìm cột tên để tìm kiếm
            col_name = next((c for c in df.columns if 'ho' in c.lower() or 'ten' in c.lower()), df.columns[0])
            search = st.text_input("🔍 Tìm tên cán bộ:", placeholder="Gõ tên để tìm nhanh...")
            
            if search:
                df = df[df[col_name].astype(str).str.contains(search, case=False, na=False)]

            # Hiển thị bảng an toàn
            st.dataframe(df, use_container_width=True, hide_index=True)
            
        else:
            st.warning("Kết nối thành công nhưng bảng dữ liệu đang trống.")

    except Exception as e:
        # Hiển thị lỗi dưới dạng chuỗi để tránh lỗi Encode khi in lỗi ra màn hình
        st.error(f"Hệ thống đang gặp vấn đề về hiển thị. Sếp hãy F5 lại nhé!")
        st.info(f"Chi tiết kỹ thuật: {str(e).encode('utf-8')}")

if __name__ == "__main__":
    main()
