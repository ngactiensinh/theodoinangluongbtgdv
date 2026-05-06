import streamlit as st
import pandas as pd
from supabase import create_client
from datetime import datetime
from dateutil.relativedelta import relativedelta
import json
import plotly.express as px  # THƯ VIỆN VẼ BIỂU ĐỒ MỚI

# 1. CẤU HÌNH TRANG
st.set_page_config(page_title="Quản lý Lương Tuyên Quang", page_icon="📊", layout="wide")

# 2. CSS LÀM ĐẸP
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

# 3. KẾT NỐI SUPABASE
try:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    supabase = create_client(url, key)
except Exception as e:
    st.error("Lỗi cấu hình Secrets. Sếp kiểm tra lại nhé!")
    st.stop()

# 4. HÀM TỰ ĐỘNG TÍNH TOÁN
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
            
    # Dọn dẹp lỗi hiển thị chữ 'nan' thừa mứa
    res = res.fillna("")
    return res

def main():
    try:
        res = supabase.table("theo_doi_luong").select("*").execute()
        
        if res.data:
            df = pd.DataFrame(res.data)
        else:
            df = pd.DataFrame(columns=[
                "stt", "ho_ten", "chuc_vu", "ma_ngach", "ngach_luong", "bac_luong",
                "he_so_hien_tai", "vuot_khung_hien_tai", "ngay_gan_nhat",
                "bac_luong_moi", "he_so_moi", "vuot_khung_moi", "ngay_du_kien",
                "trang_thai", "loai_nang_luong", "ghi_chu"
            ])
            
        df_calculated = tinh_toan_nang_luong(df)
        
        # --- TRỢ LÝ THÔNG BÁO ---
        col_status = next((c for c in df_calculated.columns if 'trang' in c.lower()), None)
        st.subheader("🤖 Trợ lý Nhân sự")
        with st.chat_message("assistant"):
            st.write("Chào Sếp Tuấn! Hệ thống đã phục hồi Bộ lọc thông minh và bổ sung Bảng điều khiển.")
            if col_status and not df_calculated.empty:
                df_calculated[col_status] = df_calculated[col_status].astype(str)
                sap_den_han = df_calculated[df_calculated[col_status].str.contains("Sắp đến hạn|Đã quá hạn", na=False)]
                if not sap_den_han.empty:
                    st.error(f"🚨 Đang có **{len(sap_den_han)}** đồng chí sắp (hoặc đã) đến hạn nâng lương. Sếp chuyển sang Tab Dữ liệu lọc xem ai nhé!")
                else:
                    st.success("✅ Danh sách đều ổn thỏa, chưa có ai đến hạn.")

        st.write("---")

        # --- CHIA 2 TAB: QUẢN LÝ & THỐNG KÊ ---
        tab1, tab2 = st.tabs(["📋 Quản lý & Lọc Dữ liệu", "📊 Bảng Thống kê (Dashboard)"])

        # ==========================================
        # TAB 1: BẢNG DỮ LIỆU & BỘ LỌC (Đã mang trả lại sếp)
        # ==========================================
        with tab1:
            # 🌟 BỘ LỌC ĐÃ TRỞ LẠI 🌟
            c1, c2 = st.columns([2, 1])
            with c1:
                search = st.text_input("🔍 Tra cứu tên cán bộ / chức vụ:", placeholder="Nhập tên hoặc chức vụ để tìm nhanh...")
            with c2:
                loc_thoi_gian = st.selectbox("⏳ Lọc theo thời gian đến hạn:", 
                                            ["Tất cả", "Trong tháng này", "Trong Quý này", "Trong 6 tháng tới", "Trong năm nay", "Đã quá hạn"])

            # Xử lý Lọc
            df_display = df_calculated.copy()
            
            if search:
                # Tìm cả trong cột Tên và Chức vụ
                mask = df_display.astype(str).apply(lambda x: x.str.contains(search, case=False, na=False)).any(axis=1)
                df_display = df_display[mask]
                
            if loc_thoi_gian != "Tất cả":
                today = datetime.now()
                df_display['ngay_dk_dt'] = pd.to_datetime(df_display['ngay_du_kien'], format='%d/%m/%Y', errors='coerce')
                df_display['so_ngay'] = (df_display['ngay_dk_dt'] - today).dt.days
                
                if loc_thoi_gian == "Trong tháng này":
                    df_display = df_display[(df_display['so_ngay'] >= 0) & (df_display['so_ngay'] <= 30)]
                elif loc_thoi_gian == "Trong Quý này":
                    df_display = df_display[(df_display['so_ngay'] >= 0) & (df_display['so_ngay'] <= 90)]
                elif loc_thoi_gian == "Trong 6 tháng tới":
                    df_display = df_display[(df_display['so_ngay'] >= 0) & (df_display['so_ngay'] <= 180)]
                elif loc_thoi_gian == "Trong năm nay":
                    df_display = df_display[df_display['ngay_dk_dt'].dt.year == today.year]
                elif loc_thoi_gian == "Đã quá hạn":
                    df_display = df_display[df_display['so_ngay'] < 0]
                    
                df_display = df_display.drop(columns=['ngay_dk_dt', 'so_ngay'], errors='ignore')

            # Hiển thị Data Editor
            def color_status(val):
                val_str = str(val)
                if "Sắp đến hạn" in val_str or "Đã quá hạn" in val_str: return 'color: red; font-weight: bold'
                return 'color: green'

            # --- XỬ LÝ LỖI MÃ NGẠCH BỊ BIẾN THÀNH SỐ ---
            def format_ma_ngach(val):
                if pd.isna(val) or val == "" or str(val).lower() == "nan":
                    return ""
                val_str = str(val).strip()
                # Nếu hệ thống đang lỡ lưu kiểu 1001.0 thì tự động gọt đuôi .0 đi cho đẹp
                if val_str.endswith(".0"):
                    val_str = val_str[:-2]
                return val_str

            df_display['ma_ngach'] = df_display['ma_ngach'].apply(format_ma_ngach)

            # Hiển thị Data Editor
            st.caption("✍️ Sửa trực tiếp trên bảng. Sửa xong bấm LƯU để máy tự cộng cột Tương lai!")
            edited_df = st.data_editor(
                df_display.style.map(color_status, subset=['trang_thai']),
                num_rows="dynamic",
                column_config={
                    # 🌟 ÉP CHUẨN CỘT MÃ NGẠCH THÀNH VĂN BẢN (TEXT) 🌟
                    "ma_ngach": st.column_config.TextColumn("Mã ngạch")
                },
                disabled=["bac_luong_moi", "he_so_moi", "vuot_khung_moi", "ngay_du_kien", "trang_thai"],
                use_container_width=True,
                hide_index=True
            )
            
            # --- NÚT LƯU & TẢI ---
            col_luu, col_tai = st.columns(2)
            with col_luu:
                if st.button("💾 Lưu thay đổi & Tính toán lại"):
                    try:
                        current_df = edited_df.data if hasattr(edited_df, 'data') else edited_df
                        luu_df = current_df[current_df['ho_ten'].str.strip().astype(bool)].copy()
                        
                        cols_to_drop = ["bac_luong_moi", "he_so_moi", "vuot_khung_moi", "ngay_du_kien", "trang_thai", "id"]
                        luu_df = luu_df.drop(columns=[c for c in cols_to_drop if c in luu_df.columns], errors='ignore')
                        
                        # Rà soát tử hình NaN siêu chặt chẽ
                        records = []
                        for r in luu_df.to_dict(orient="records"):
                            clean_r = {}
                            for k, v in r.items():
                                if pd.isna(v) or v == "" or str(v).lower() == 'nan':
                                    clean_r[k] = None
                                else:
                                    clean_r[k] = v
                            records.append(clean_r)
                            
                        json.dumps(records) # Chạy máy quét JSON
                        
                        supabase.table("theo_doi_luong").delete().neq("ho_ten", "Xóa_Tất_Cả").execute()
                        if records:
                            supabase.table("theo_doi_luong").insert(records).execute()
                            
                        st.success("🎉 Lưu thành công tuyệt đối!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Lỗi nhập liệu: {e}")

            with col_tai:
                out_df = edited_df.data if hasattr(edited_df, 'data') else edited_df
                st.download_button(
                    "📥 Xuất file báo cáo (CSV)",
                    out_df.to_csv(index=False).encode('utf-8-sig'),
                    f"bao_cao_luong_{datetime.now().strftime('%d%m%Y')}.csv",
                    "text/csv"
                )

        # ==========================================
        # TAB 2: DASHBOARD BIỂU ĐỒ (Order của sếp)
        # ==========================================
        with tab2:
            st.subheader("Báo cáo phân tích tổng quan")
            if not df_calculated.empty:
                # Trích xuất dữ liệu sạch để vẽ
                df_chart = df_calculated.copy()
                df_chart['vuot_khung_hien_tai'] = df_chart['vuot_khung_hien_tai'].replace(['None', '', None, 'nan'], 'Không Vượt khung')
                
                # Chia 2 cột cho hàng đầu tiên
                c_bieu_1, c_bieu_2 = st.columns(2)
                
                with c_bieu_1:
                    # 1. Biểu đồ Ngạch lương (Pie)
                    df_ngach = df_chart['ngach_luong'].value_counts().reset_index()
                    df_ngach.columns = ['Ngạch', 'Số lượng']
                    fig_ngach = px.pie(df_ngach, values='Số lượng', names='Ngạch', title='1. Cơ cấu Cán bộ theo Ngạch lương', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
                    st.plotly_chart(fig_ngach, use_container_width=True)
                
                with c_bieu_2:
                    # 2. Biểu đồ Vượt khung (Pie)
                    df_vk = df_chart['vuot_khung_hien_tai'].value_counts().reset_index()
                    df_vk.columns = ['Tình trạng', 'Số lượng']
                    # Gộp tất cả các % vượt khung lại thành 1 nhóm "Đang hưởng Vượt khung"
                    df_vk['Nhóm'] = df_vk['Tình trạng'].apply(lambda x: 'Không Vượt khung' if x == 'Không Vượt khung' else 'Đang hưởng Vượt khung')
                    df_vk_group = df_vk.groupby('Nhóm')['Số lượng'].sum().reset_index()
                    fig_vk = px.pie(df_vk_group, values='Số lượng', names='Nhóm', title='2. Tỷ lệ hưởng phụ cấp Vượt khung', color_discrete_sequence=['#17a2b8', '#ffc107'])
                    st.plotly_chart(fig_vk, use_container_width=True)

                st.divider()
                
                # Chia 2 cột cho hàng thứ hai
                c_bieu_3, c_bieu_4 = st.columns(2)
                
                with c_bieu_3:
                    # 3. Biểu đồ Bậc lương (Bar)
                    df_bac = df_chart['bac_luong'].value_counts().reset_index()
                    df_bac.columns = ['Bậc lương', 'Số lượng cán bộ']
                    df_bac = df_bac.sort_values(by='Bậc lương')
                    fig_bac = px.bar(df_bac, x='Bậc lương', y='Số lượng cán bộ', title='3. Phân bổ Cán bộ theo Bậc lương hiện tại', text='Số lượng cán bộ')
                    fig_bac.update_traces(textposition='outside')
                    st.plotly_chart(fig_bac, use_container_width=True)
                
                with c_bieu_4:
                    # 4. Biểu đồ Mã ngạch (Bar ngang)
                    # Ép kiểu toàn bộ Mã ngạch sang chuỗi (String) trước khi đếm
                    df_chart['ma_ngach'] = df_chart['ma_ngach'].fillna("Chưa có").astype(str)
                    
                    df_ma = df_chart['ma_ngach'].value_counts().reset_index()
                    df_ma.columns = ['Mã ngạch', 'Số lượng']
                    
                    fig_ma = px.bar(df_ma, x='Số lượng', y='Mã ngạch', orientation='h', 
                                    title='4. Phân bổ theo Mã ngạch', text='Số lượng', color='Mã ngạch')
                    
                    # Lệnh "thần thánh" ép Plotly hiện thị trục Y dạng Danh mục (Category)
                    fig_ma.update_yaxes(type='category', categoryorder='total ascending')
                    fig_ma.update_layout(showlegend=False)
                    
                    st.plotly_chart(fig_ma, use_container_width=True)
                    
            else:
                st.info("Chưa có dữ liệu để vẽ biểu đồ.")

    except Exception as e:
        st.error(f"Hệ thống gặp sự cố: {e}")

if __name__ == "__main__":
    main()
