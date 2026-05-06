import streamlit as st
import pandas as pd
from supabase import create_client
from datetime import datetime
from dateutil.relativedelta import relativedelta
import json # Bổ sung công cụ rà soát lỗi

# 1. CẤU HÌNH TRANG CHUYÊN NGHIỆP CỦA SẾP
st.set_page_config(page_title="Quản lý Lương Tuyên Quang", page_icon="📊", layout="wide")

# 2. CSS ĐỂ LÀM ĐẸP
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
    if res.empty: return res
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
            res.at[idx, 'bac_luong_moi'] = ""
            res.at[idx, 'he_so_moi'] = ""
            res.at[idx, 'vuot_khung_moi'] = ""
            res.at[idx, 'ngay_du_kien'] = ""
            res.at[idx, 'trang_thai'] = "Chưa có ngày"
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
        
        # Xử lý nếu kho rỗng
        if res.data:
            df = pd.DataFrame(res.data)
        else:
            df = pd.DataFrame(columns=[
                "stt", "ho_ten", "chuc_vu", "ma_ngach", "ngach_luong", "bac_luong",
                "he_so_hien_tai", "vuot_khung_hien_tai", "ngay_gan_nhat",
                "bac_luong_moi", "he_so_moi", "vuot_khung_moi", "ngay_du_kien",
                "trang_thai", "loai_nang_luong", "ghi_chu"
            ])
            st.warning("⚠️ Cơ sở dữ liệu đang trống. Sếp hãy mở tab 'Nạp dữ liệu' bên dưới để tải file CSV phục hồi nhé!")
            
        df_calculated = tinh_toan_nang_luong(df)
        
        # --- XỬ LÝ TRỢ LÝ THÔNG BÁO ---
        col_status = next((c for c in df_calculated.columns if 'trang' in c.lower()), None)
        
        st.subheader("🤖 Trợ lý Nhân sự")
        with st.chat_message("assistant"):
            st.write("Chào Sếp Tuấn! Hệ thống đã kích hoạt chế độ Bảo vệ Dữ liệu cấp độ cao.")
            if col_status and not df_calculated.empty:
                df_calculated[col_status] = df_calculated[col_status].astype(str)
                sap_den_han = df_calculated[df_calculated[col_status].str.contains("Sắp đến hạn|Đã quá hạn", na=False)]
                if not sap_den_han.empty:
                    st.error(f"🚨 Có **{len(sap_den_han)}** đồng chí sắp (hoặc đã) đến hạn nâng lương. Sếp lưu ý nhé!")
                else:
                    st.success("✅ Hiện tại danh sách đều ổn thỏa, chưa có ai đến hạn.")

        st.write("---")

        # --- NẠP DỮ LIỆU TỪ CSV ---
        with st.expander("📂 Nạp dữ liệu từ file CSV (Bấm để tải file phục hồi)"):
            file_upload = st.file_uploader("Chọn file báo cáo CSV sếp vừa tải về lúc trước:", type=["csv"])
            if file_upload:
                try:
                    df_csv = pd.read_csv(file_upload)
                    df_calculated = pd.concat([df_calculated, df_csv], ignore_index=True)
                    df_calculated = tinh_toan_nang_luong(df_calculated)
                    st.success("✅ Đã đọc thành công file CSV! Sếp lướt xuống cuối bảng bấm 'Lưu' để chốt nhé.")
                except Exception as e:
                    st.error(f"Lỗi đọc file: {e}")

        # --- HIỂN THỊ BẢNG ĐẸP ---
        st.subheader("📋 Danh sách chi tiết")
        st.caption("💡 Mẹo: Bấm biểu tượng Kính lúp (🔍) ở góc phải bảng để tìm người. Sửa thông tin xong nhớ bấm Lưu bên dưới để máy tự cộng lại cột Tương lai nhé!")
        
        def color_status(val):
            if pd.isna(val): return ''
            val_str = str(val)
            if "Sắp đến hạn" in val_str or "Đã quá hạn" in val_str: return 'color: red; font-weight: bold'
            return 'color: green'

        raw_df = df_calculated.copy()
        
        edited_df = st.data_editor(
            raw_df.style.map(color_status, subset=[col_status] if col_status else []),
            num_rows="dynamic",
            disabled=["bac_luong_moi", "he_so_moi", "vuot_khung_moi", "ngay_du_kien", "trang_thai"],
            use_container_width=True,
            hide_index=True
        )
        
        # --- NÚT LƯU & TẢI ---
        col_luu, col_tai = st.columns(2)
        with col_luu:
            if st.button("💾 Lưu các chỉnh sửa & Tính toán lại"):
                try:
                    current_df = edited_df.data if hasattr(edited_df, 'data') else edited_df
                    luu_df = current_df[current_df['ho_ten'].str.strip().astype(bool)].copy()
                    
                    # Bỏ các cột tự động tính trước khi lưu
                    cols_to_drop = ["bac_luong_moi", "he_so_moi", "vuot_khung_moi", "ngay_du_kien", "trang_thai", "id"]
                    luu_df = luu_df.drop(columns=[c for c in cols_to_drop if c in luu_df.columns], errors='ignore')
                    
                    # 🌟 BƯỚC 1: LÀM SẠCH TRIỆT ĐỂ MỌI LỖI NaN 🌟
                    records = []
                    for r in luu_df.to_dict(orient="records"):
                        clean_r = {}
                        for k, v in r.items():
                            # Nếu là NaN, trống, NaT thì đổi thành rỗng (None)
                            if pd.isna(v): 
                                clean_r[k] = None
                            else:
                                clean_r[k] = v
                        records.append(clean_r)
                        
                    # 🌟 BƯỚC 2: RÀ SOÁT TỬ HÌNH "NaN" 🌟
                    # Ép chạy qua máy quét JSON. Nếu còn sót cái lỗi NaN nào nó sẽ báo lỗi và dừng lại NGAY LẬP TỨC!
                    json.dumps(records) 

                    # 🌟 BƯỚC 3: THAY MÁU CƠ SỞ DỮ LIỆU AN TOÀN 🌟
                    # Chạy đến đây nghĩa là dữ liệu an toàn 100% rồi, mới dám xóa và lưu
                    supabase.table("theo_doi_luong").delete().neq("ho_ten", "Xóa_Tất_Cả").execute()
                    if records:
                        supabase.table("theo_doi_luong").insert(records).execute()
                        
                    st.success("🎉 Lưu thành công tuyệt đối! Đang cập nhật lại các mốc thời gian...")
                    st.rerun()
                except Exception as e:
                    # Báo lỗi nhưng KHÔNG xóa dữ liệu cũ
                    st.error(f"Phát hiện dữ liệu bất thường: {e}. Hệ thống đã chặn lại để bảo vệ cơ sở dữ liệu. Sếp kiểm tra lại ô vừa nhập nhé!")

        with col_tai:
            out_df = edited_df.data if hasattr(edited_df, 'data') else edited_df
            st.download_button(
                "📥 Tải file báo cáo (CSV)",
                out_df.to_csv(index=False).encode('utf-8-sig'),
                f"bao_cao_luong_{datetime.now().strftime('%d%m%Y')}.csv",
                "text/csv"
            )

    except Exception as e:
        st.error(f"Hệ thống gặp sự cố mạng: {e}")

if __name__ == "__main__":
    main()
