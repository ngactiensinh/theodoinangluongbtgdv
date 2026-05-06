import streamlit as st
import pandas as pd
from supabase import create_client
from datetime import datetime
from dateutil.relativedelta import relativedelta

# 1. CẤU HÌNH TRANG CHUYÊN NGHIỆP CỦA SẾP
st.set_page_config(page_title="Quản lý Lương Tuyên Quang", page_icon="📊", layout="wide")

# 2. CSS ĐỂ LÀM ĐẸP (Dùng bảng mã an toàn của sếp)
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

# 3. KẾT NỐI (Lấy từ Secrets)
try:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    supabase = create_client(url, key)
except Exception as e:
    st.error("Lỗi cấu hình Secrets. Sếp kiểm tra lại nhé!")
    st.stop()

# 4. HÀM TỰ ĐỘNG TÍNH TOÁN THEO LUẬT
def tinh_toan_nang_luong(df):
    res = df.copy()
    today = datetime.now().date()
    
    for idx, row in res.iterrows():
        ngach = str(row.get('ngach_luong', '')).strip().upper()
        chuc_vu = str(row.get('chuc_vu', '')).strip().upper()
        bac_ht = str(row.get('bac_luong', '')).strip() 
        
        hs_str = str(row.get('he_so_hien_tai', '0')).replace(',', '.')
        try:
            hs_ht = float(hs_str)
        except:
            hs_ht = 0.0
            
        vk_ht = str(row.get('vuot_khung_hien_tai', 'None')).strip()
        ngay_ht_str = str(row.get('ngay_gan_nhat', ''))
        
        try:
            ngay_ht = datetime.strptime(ngay_ht_str, '%d/%m/%Y').date()
        except:
            continue
            
        is_vk = False
        vk_val = 0
        if vk_ht.lower() != 'none' and '%' in vk_ht:
            is_vk = True
            vk_val = int(vk_ht.replace('%', '').strip())
            
        bac_moi = bac_ht
        hs_moi = hs_ht
        vk_moi = vk_ht
        ngay_dk = ngay_ht
        
        if is_vk:
            ngay_dk = ngay_ht + relativedelta(years=1)
            vk_moi = f"{vk_val + 1}%"
        else:
            try:
                if '/' in bac_ht:
                    x_str, y_str = bac_ht.split('/')
                    x, y = int(x_str), int(y_str)
                else:
                    x, y = int(bac_ht), 99
                
                if x >= y:
                    ngay_dk = ngay_ht + relativedelta(years=3) 
                    vk_moi = "5%"
                else:
                    bac_moi = f"{x+1}/{y}"
                    if 'KẾ TOÁN VIÊN TRUNG CẤP' in ngach or 'KẾ TOÁN TRUNG CẤP' in ngach:
                        interval, delta = 2, 0.20
                    elif 'LÁI XE' in chuc_vu or 'PHỤC VỤ' in chuc_vu or 'VĂN THƯ' in chuc_vu or 'VĂN THƯ VIÊN TRUNG CẤP' in ngach:
                        interval, delta = 2, 0.18
                    elif 'KẾ TOÁN' in ngach or ngach == 'CV':
                        interval, delta = 3, 0.33
                    elif ngach == 'CVC':
                        interval, delta = 3, 0.34
                    elif ngach == 'CVCC':
                        interval, delta = 3, 0.62 
                    else:
                        interval, delta = 3, 0.33 
                        
                    ngay_dk = ngay_ht + relativedelta(years=interval)
                    hs_moi = hs_ht + delta
            except:
                pass
                
        res.at[idx, 'bac_luong_moi'] = bac_moi
        res.at[idx, 'he_so_moi'] = f"{hs_moi:.2f}".replace('.', ',')
        res.at[idx, 'vuot_khung_moi'] = vk_moi
        res.at[idx, 'ngay_du_kien'] = ngay_dk.strftime('%d/%m/%Y')
        
        # Đã tinh chỉnh lại Trạng thái để ghép chuẩn với code của sếp
        days_left = (ngay_dk - today).days
        if days_left < 0:
            res.at[idx, 'trang_thai'] = "Đã quá hạn"
        elif days_left <= 30:
            res.at[idx, 'trang_thai'] = "Sắp đến hạn (Tháng này)"
        elif days_left <= 90:
            res.at[idx, 'trang_thai'] = "Sắp đến hạn (Quý này)"
        else:
            res.at[idx, 'trang_thai'] = "Chưa đến hạn"
            
    return res

def main():
    try:
        res = supabase.table("theo_doi_luong").select("*").execute()
        
        if res.data:
            df = pd.DataFrame(res.data)
            
            # Tính toán lại trước khi đưa vào Trợ lý báo cáo
            df_calculated = tinh_toan_nang_luong(df)
            
            # --- XỬ LÝ TRỢ LÝ THÔNG BÁO (Code siêu chuẩn của sếp) ---
            col_status = next((c for c in df_calculated.columns if 'trang' in c.lower()), None)
            
            st.subheader("🤖 Trợ lý Nhân sự")
            with st.chat_message("assistant"):
                st.write("Chào Bạn Tuấn! Hệ thống đã sẵn sàng và đã cập nhật lại các mốc thời gian.")
                if col_status:
                    df_calculated[col_status] = df_calculated[col_status].astype(str)
                    # Bổ sung thêm tìm người "Đã quá hạn" cho chắc cốp
                    sap_den_han = df_calculated[df_calculated[col_status].str.contains("Sắp đến hạn|Đã quá hạn", na=False)]
                    if not sap_den_han.empty:
                        st.error(f"🚨 Có **{len(sap_den_han)}** đồng chí sắp (hoặc đã) đến hạn nâng lương. Sếp lưu ý nhé!")
                    else:
                        st.success("✅ Hiện tại danh sách đều ổn thỏa, chưa có ai đến hạn.")

            st.write("---")

            # --- BỘ LỌC THÔNG MINH ---
            col_name = next((c for c in df_calculated.columns if 'ho' in c.lower() or 'ten' in c.lower()), df_calculated.columns[0])
            
            c1, c2 = st.columns([2, 1])
            with c1:
                search = st.text_input("🔍 Tra cứu tên cán bộ:", placeholder="Nhập tên để tìm nhanh...")
            with c2:
                if col_status:
                    status_list = ["Tất cả"] + list(df_calculated[col_status].unique())
                    filter_val = st.selectbox("Lọc theo trạng thái:", status_list)

            # Thực hiện lọc
            df_display = df_calculated.copy()
            if search:
                df_display = df_display[df_display[col_name].astype(str).str.contains(search, case=False, na=False)]
            if col_status and filter_val != "Tất cả":
                df_display = df_display[df_display[col_status] == filter_val]

            # --- HIỂN THỊ BẢNG ĐẸP ---
            st.subheader("📋 Danh sách chi tiết (Có thể chỉnh sửa)")
            st.caption("✍️ Bấm vào các ô thông tin HIỆN TẠI để sửa hoặc bấm dòng dưới cùng để THÊM MỚI. Cột TƯƠNG LAI máy tự tính.")
            
            # Định dạng màu sắc cho cột trạng thái
            def color_status(val):
                if pd.isna(val): return ''
                val_str = str(val)
                if "Sắp đến hạn" in val_str or "Đã quá hạn" in val_str: return 'color: red; font-weight: bold'
                return 'color: green'

            edited_df = st.data_editor(
                df_display.style.map(color_status, subset=[col_status] if col_status else []),
                num_rows="dynamic",
                disabled=["bac_luong_moi", "he_so_moi", "vuot_khung_moi", "ngay_du_kien", "trang_thai"],
                use_container_width=True,
                hide_index=True
            )
            
           # --- NÚT LƯU & TẢI ---
            col_luu, col_tai = st.columns(2)
            with col_luu:
                if st.button("💾 Lưu các chỉnh sửa lên cơ sở dữ liệu"):
                    try:
                        # Rút ruột DataFrame từ Styler Object
                        raw_df = edited_df.data if hasattr(edited_df, 'data') else edited_df
                        
                        supabase.table("theo_doi_luong").delete().neq("ho_ten", "Xóa_Tất_Cả").execute()
                        luu_df = raw_df[raw_df['ho_ten'].str.strip().astype(bool)].copy()
                        
                        cols_to_drop = ["bac_luong_moi", "he_so_moi", "vuot_khung_moi", "ngay_du_kien", "trang_thai"]
                        luu_df = luu_df.drop(columns=[c for c in cols_to_drop if c in luu_df.columns])
                        
                        # 🌟 THÊM DÒNG NÀY ĐỂ TRỊ LỖI "NaN" 🌟
                        # Chuyển đổi tất cả các giá trị rỗng/lỗi thành None để Supabase (JSON) hiểu được
                        luu_df = luu_df.where(pd.notnull(luu_df), None)
                        
                        records = luu_df.to_dict(orient="records")
                        if records:
                            supabase.table("theo_doi_luong").insert(records).execute()
                        st.success("Đã lưu thành công! Cập nhật xong các mốc nâng lương mới.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Lỗi khi lưu: {e}")

            with col_tai:
                raw_df = df_display.data if hasattr(df_display, 'data') else df_display
                st.download_button(
                    "📥 Tải file báo cáo (CSV)",
                    raw_df.to_csv(index=False).encode('utf-8-sig'),
                    f"bao_cao_luong_{datetime.now().strftime('%d%m%Y')}.csv",
                    "text/csv"
                )

        else:
            st.info("Chưa có dữ liệu trong bảng theo_doi_luong.")

    except Exception as e:
        st.error(f"Hệ thống đang bảo trì phần hiển thị. Sếp hãy kiểm tra lại Secrets hoặc F5 nhé! Chi tiết: {e}")

if __name__ == "__main__":
    main()
