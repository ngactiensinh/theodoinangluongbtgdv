import streamlit as st
import pandas as pd
from supabase import create_client

def main():
    # 1. CẤU HÌNH TRANG CHUYÊN NGHIỆP
    st.set_page_config(page_title="Quản lý Lương Tuyên Quang", page_icon="📊", layout="wide")

    # 2. CSS ĐỂ LÀM ĐẸP (Dùng bảng mã an toàn)
    st.markdown("""
        <style>
        .main-header {
            background: linear-gradient(135deg, #004B87 0%, #17a2b8 100%);
            padding: 30px; border-radius: 15px; color: white; text-align: center; margin-bottom: 25px;
        }
        .stDataFrame { border: 1px solid #e6e9ef; border-radius: 10px; }
        </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="main-header"><h1>📈 HỆ THỐNG QUẢN LÝ LƯƠNG 4.0</h1><p>Ban Tuyên giáo và Dân vận Tỉnh ủy Tuyên Quang</p></div>', unsafe_allow_html=True)

    try:
        # 3. KẾT NỐI (Lấy từ Secrets)
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
        supabase = create_client(url, key)
        
        # 4. LẤY DỮ LIỆU
        res = supabase.table("theo_doi_luong").select("*").execute()
        
        if res.data:
            df = pd.DataFrame(res.data)
            
            # --- XỬ LÝ TRỢ LÝ THÔNG BÁO ---
            col_status = next((c for c in df.columns if 'trang' in c.lower()), None)
            
            st.subheader("🤖 Trợ lý Nhân sự")
            with st.chat_message("assistant"):
                st.write("Chào Bạn Tuấn! Hệ thống đã sẵn sàng.")
                if col_status:
                    df[col_status] = df[col_status].astype(str)
                    sap_den_han = df[df[col_status].str.contains("Sắp đến hạn", na=False)]
                    if not sap_den_han.empty:
                        st.error(f"🚨 Có **{len(sap_den_han)}** đồng chí sắp đến hạn nâng lương. Sếp lưu ý nhé!")
                    else:
                        st.success("✅ Hiện tại danh sách đều ổn thỏa, chưa có ai đến hạn.")

            st.write("---")

            # --- BỘ LỌC THÔNG MINH ---
            col_name = next((c for c in df.columns if 'ho' in c.lower() or 'ten' in c.lower()), df.columns[0])
            
            c1, c2 = st.columns([2, 1])
            with c1:
                search = st.text_input("🔍 Tra cứu tên cán bộ:", placeholder="Nhập tên để tìm nhanh...")
            with c2:
                # Thêm lọc theo trạng thái nếu có cột
                if col_status:
                    status_list = ["Tất cả"] + list(df[col_status].unique())
                    filter_val = st.selectbox("Lọc theo trạng thái:", status_list)

            # Thực hiện lọc
            df_display = df.copy()
            if search:
                df_display = df_display[df_display[col_name].astype(str).str.contains(search, case=False, na=False)]
            if col_status and filter_val != "Tất cả":
                df_display = df_display[df_display[col_status] == filter_val]

            # --- HIỂN THỊ BẢNG ĐẸP ---
            st.subheader("📋 Danh sách chi tiết")
            
            # Định dạng màu sắc cho cột trạng thái
            def color_status(val):
                if "Sắp đến hạn" in str(val): return 'color: red; font-weight: bold'
                return 'color: green'

            st.dataframe(
                df_display.style.map(color_status, subset=[col_status] if col_status else []),
                use_container_width=True,
                hide_index=True
            )
            
            # Nút tải dữ liệu
            st.download_button(
                "📥 Tải file báo cáo (CSV)",
                df_display.to_csv(index=False).encode('utf-8-sig'),
                "bao_cao_luong.csv",
                "text/csv"
            )

        else:
            st.info("Chưa có dữ liệu trong bảng theo_doi_luong.")

    except Exception as e:
        st.error("Hệ thống đang bảo trì phần hiển thị. Sếp hãy kiểm tra lại Secrets hoặc F5 nhé!")

if __name__ == "__main__":
    main()
