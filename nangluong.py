import streamlit as st
import pandas as pd
from supabase import create_client
from datetime import datetime
from dateutil.relativedelta import relativedelta
import io

# 1. KẾT NỐI SUPABASE
URL = st.secrets["SUPABASE_URL"]
KEY = st.secrets["SUPABASE_KEY"]
supabase = create_client(URL, KEY)

# 2. HÀM TỰ ĐỘNG TÍNH TOÁN (ĐÚNG LUẬT CỦA SẾP & KHỚP FORM ẢNH)
def tinh_toan_nang_luong(df):
    res = df.copy()
    today = datetime.now().date()
    
    for idx, row in res.iterrows():
        # Lấy dữ liệu hiện tại
        ngach = str(row.get('ngach_luong', '')).strip().upper()
        chuc_vu = str(row.get('chuc_vu', '')).strip().upper()
        bac_ht = str(row.get('bac_luong', '')).strip() # Ví dụ: '5/8'
        
        # Xử lý hệ số có dấu phẩy (ví dụ: '6,78' -> 6.78)
        hs_str = str(row.get('he_so_hien_tai', '0')).replace(',', '.')
        try:
            hs_ht = float(hs_str)
        except:
            hs_ht = 0.0
            
        vk_ht = str(row.get('vuot_khung_hien_tai', 'None')).strip()
        ngay_ht_str = str(row.get('ngay_gan_nhat', ''))
        
        # Chuyển đổi ngày
        try:
            ngay_ht = datetime.strptime(ngay_ht_str, '%d/%m/%Y').date()
        except:
            continue # Bỏ qua nếu nhập sai ngày
            
        # Kiểm tra đang ở chế độ Vượt khung không?
        is_vk = False
        vk_val = 0
        if vk_ht.lower() != 'none' and '%' in vk_ht:
            is_vk = True
            vk_val = int(vk_ht.replace('%', '').strip())
            
        # Khởi tạo giá trị MỚI mặc định
        bac_moi = bac_ht
        hs_moi = hs_ht
        vk_moi = vk_ht
        ngay_dk = ngay_ht
        
        if is_vk:
            # LUẬT VƯỢT KHUNG: Giữ nguyên hệ số, +1 năm, +1%
            ngay_dk = ngay_ht + relativedelta(years=1)
            vk_moi = f"{vk_val + 1}%"
        else:
            # LUẬT BÌNH THƯỜNG / CHUYỂN TIẾP VƯỢT KHUNG
            try:
                # Tách '5/8' thành x=5, y=8
                if '/' in bac_ht:
                    x_str, y_str = bac_ht.split('/')
                    x, y = int(x_str), int(y_str)
                else:
                    x, y = int(bac_ht), 99 # Xử lý lỗi nếu ai nhập thiếu '/'
                
                if x >= y:
                    # Kịch trần bậc -> Chuyển sang Vượt khung 5% (Thường là sau 3 năm)
                    ngay_dk = ngay_ht + relativedelta(years=3) 
                    vk_moi = "5%"
                else:
                    # Tăng bậc bình thường
                    bac_moi = f"{x+1}/{y}"
                    
                    # Xác định thời gian và hệ số cộng thêm theo luật sếp đưa
                    # Kế toán trung cấp, Lái xe, Phục vụ, Văn thư trung cấp -> 2 năm
                    if 'KẾ TOÁN VIÊN TRUNG CẤP' in ngach or 'KẾ TOÁN TRUNG CẤP' in ngach:
                        interval, delta = 2, 0.20
                    elif 'LÁI XE' in chuc_vu or 'PHỤC VỤ' in chuc_vu or 'VĂN THƯ' in chuc_vu or 'VĂN THƯ VIÊN TRUNG CẤP' in ngach:
                        interval, delta = 2, 0.18
                    # CV, CVC, CVCC, Kế toán viên -> 3 năm
                    elif 'KẾ TOÁN' in ngach or ngach == 'CV':
                        interval, delta = 3, 0.33
                    elif ngach == 'CVC':
                        interval, delta = 3, 0.34
                    elif ngach == 'CVCC':
                        interval, delta = 3, 0.62 # Hệ số trần CVCC
                    else:
                        interval, delta = 3, 0.33 # Mặc định
                        
                    ngay_dk = ngay_ht + relativedelta(years=interval)
                    hs_moi = hs_ht + delta
            except:
                pass
                
        # Cập nhật kết quả vào cột (định dạng lại hệ số dùng dấu phẩy cho đẹp)
        res.at[idx, 'bac_luong_moi'] = bac_moi
        res.at[idx, 'he_so_moi'] = f"{hs_moi:.2f}".replace('.', ',')
        res.at[idx, 'vuot_khung_moi'] = vk_moi
        res.at[idx, 'ngay_du_kien'] = ngay_dk.strftime('%d/%m/%Y')
        
        # Cập nhật TRẠNG THÁI thông minh
        days_left = (ngay_dk - today).days
        if days_left < 0:
            res.at[idx, 'trang_thai'] = "Đã quá hạn"
        elif days_left <= 30:
            res.at[idx, 'trang_thai'] = "Đến hạn tháng này"
        elif days_left <= 90:
            res.at[idx, 'trang_thai'] = "Sắp đến hạn (Quý này)"
        else:
            res.at[idx, 'trang_thai'] = "Chưa đến hạn"
            
    return res

def main():
    st.set_page_config(page_title="Theo dõi nâng lương", layout="wide")
    
    st.title("📋 Ứng dụng Theo dõi Nâng lương Tự động")
    
    # 3. LẤY DỮ LIỆU TỪ SUPABASE
    try:
        res = supabase.table("danh_sach_luong_tuyenquang").select("*").execute()
        df = pd.DataFrame(res.data) if res.data else pd.DataFrame()
    except Exception:
        st.error("Chưa kết nối được cơ sở dữ liệu. Sếp kiểm tra lại bảng 'danh_sach_luong_tuyenquang' nhé.")
        return

    # Nếu bảng rỗng, tạo khung chuẩn theo ảnh của sếp
    if df.empty:
        df = pd.DataFrame(columns=[
            "stt", "ho_ten", "chuc_vu", "ma_ngach", "ngach_luong", "bac_luong", 
            "he_so_hien_tai", "vuot_khung_hien_tai", "ngay_gan_nhat", 
            "bac_luong_moi", "he_so_moi", "vuot_khung_moi", "ngay_du_kien", 
            "trang_thai", "loai_nang_luong", "ghi_chu"
        ])

    # 4. GIAO DIỆN LỌC (Phía trên cùng)
    col_search, col_filter = st.columns([2, 1])
    with col_search:
        search_query = st.text_input("🔍 Tra cứu tên cán bộ:", placeholder="Nhập tên để tìm nhanh...")
    with col_filter:
        loc_thoi_gian = st.selectbox("Lọc theo thời gian đến hạn:", 
                                     ["Tất cả", "Trong tháng này", "Trong Quý này", "Trong 6 tháng tới", "Trong 9 tháng tới", "Trong năm nay"])

    # 5. CHẠY BỘ NÃO TÍNH TOÁN
    df_calculated = tinh_toan_nang_luong(df)

    # 6. XỬ LÝ LỌC
    # Lọc theo tên
    if search_query:
        df_calculated = df_calculated[df_calculated['ho_ten'].str.contains(search_query, case=False, na=False)]
    
    # Lọc theo thời gian
    if loc_thoi_gian != "Tất cả":
        today = datetime.now()
        df_calculated['ngay_dk_datetime'] = pd.to_datetime(df_calculated['ngay_du_kien'], format='%d/%m/%Y', errors='coerce')
        df_calculated['so_ngay'] = (df_calculated['ngay_dk_datetime'] - today).dt.days
        
        if loc_thoi_gian == "Trong tháng này":
            df_calculated = df_calculated[(df_calculated['so_ngay'] >= 0) & (df_calculated['so_ngay'] <= 30)]
        elif loc_thoi_gian == "Trong Quý này":
            df_calculated = df_calculated[(df_calculated['so_ngay'] >= 0) & (df_calculated['so_ngay'] <= 90)]
        elif loc_thoi_gian == "Trong 6 tháng tới":
            df_calculated = df_calculated[(df_calculated['so_ngay'] >= 0) & (df_calculated['so_ngay'] <= 180)]
        elif loc_thoi_gian == "Trong 9 tháng tới":
            df_calculated = df_calculated[(df_calculated['so_ngay'] >= 0) & (df_calculated['so_ngay'] <= 270)]
        elif loc_thoi_gian == "Trong năm nay":
            df_calculated = df_calculated[df_calculated['ngay_dk_datetime'].dt.year == today.year]
            
        # Dọn dẹp cột tạm
        df_calculated = df_calculated.drop(columns=['ngay_dk_datetime', 'so_ngay'])

    # 7. HIỂN THỊ BẢNG CHO PHÉP CHỈNH SỬA (Sửa, Thêm, Xóa)
    st.markdown("### 📋 Danh sách chi tiết")
    st.caption("✍️ Bấm vào các cột thông tin HIỆN TẠI để chỉnh sửa. Các cột TƯƠNG LAI sẽ được máy tự động tính và điền giúp sếp!")
    
    edited_df = st.data_editor(
        df_calculated,
        num_rows="dynamic",
        disabled=["bac_luong_moi", "he_so_moi", "vuot_khung_moi", "ngay_du_kien", "trang_thai"], # Khóa các cột tự động tính
        use_container_width=True,
        hide_index=True
    )

    # 8. NÚT LƯU & XUẤT FILE
    col_luu, col_xuat = st.columns(2)
    with col_luu:
        if st.button("💾 Lưu thay đổi & Tính toán lại"):
            try:
                # Xóa sạch data cũ để chèn lại cho đồng bộ (cách nhanh gọn nhất)
                supabase.table("danh_sach_luong_tuyenquang").delete().neq("ho_ten", "Xóa_Tất_Cả").execute()
                
                # Loại bỏ những dòng trắng do lỡ tay bấm thêm
                luu_df = edited_df[edited_df['ho_ten'].str.strip().astype(bool)].copy()
                
                records = luu_df.to_dict(orient="records")
                if records:
                    supabase.table("danh_sach_luong_tuyenquang").insert(records).execute()
                
                st.success("Đã lưu và cập nhật các mốc thời gian nâng lương mới!")
                st.rerun() # Tự động làm mới trang
            except Exception as e:
                st.error(f"Có lỗi khi lưu: {e}")

    with col_xuat:
        # Xuất file CSV (như trong ảnh cũ của sếp)
        csv = df_calculated.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 Tải file báo cáo (CSV)",
            data=csv,
            file_name=f"Bao_cao_luong_{datetime.now().strftime('%d%m%Y')}.csv",
            mime="text/csv",
        )

if __name__ == "__main__":
    main()
