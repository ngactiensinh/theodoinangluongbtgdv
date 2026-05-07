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
import io
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime

def tao_file_word_dien_bien(df):
    doc = Document()
    
    # 1. Cài đặt khổ giấy ngang A4
    section = doc.sections[0]
    new_w, new_h = section.page_height, section.page_width
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = new_w
    section.page_height = new_h
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)

    # 2. Header: Quốc hiệu, Tiêu ngữ / Đảng hiệu
    table_header = doc.add_table(rows=1, cols=2)
    table_header.autofit = False
    table_header.columns[0].width = Cm(10)
    table_header.columns[1].width = Cm(16)

    # Cột trái (Đã đổi Hà Giang thành Tuyên Quang)
    cell_left = table_header.cell(0, 0)
    p_left = cell_left.paragraphs[0]
    p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run1 = p_left.add_run("TỈNH UỶ TUYÊN QUANG\n")
    run1.bold = True
    run2 = p_left.add_run("BAN TUYÊN GIÁO VÀ DÂN VẬN\n")
    run2.bold = True
    p_left.add_run("*")

    # Cột phải
    cell_right = table_header.cell(0, 1)
    p_right = cell_right.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run3 = p_right.add_run("ĐẢNG CỘNG SẢN VIỆT NAM\n")
    run3.bold = True
    
    # 3. Tiêu đề chính
    doc.add_paragraph() 
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title1 = p_title.add_run("BIỂU TỔNG HỢP DIỄN BIẾN LƯƠNG\n")
    run_title1.bold = True
    run_title1.font.size = Pt(14)
    
    thang_nam = datetime.now().strftime('tháng %m năm %Y')
    run_title2 = p_title.add_run(f"Ban Tuyên giáo và Dân vận Tỉnh uỷ {thang_nam}")
    run_title2.italic = True
    run_title2.font.size = Pt(12)
    
    doc.add_paragraph()

    # 4. Kẻ Bảng dữ liệu chuẩn 11 cột
    table = doc.add_table(rows=1, cols=11)
    table.style = 'Table Grid'
    
    headers = ['TT', 'Họ và tên', 'Chức vụ', 'Mã ngạch lương hiện hưởng', 'Bậc lương', 'Hệ số lương hiện tại', 'Ngày hưởng', 'Nâng bậc lương tiếp theo', 'Hệ số', 'Hưởng từ ngày', 'Ghi chú']
    # Chia tỷ lệ độ rộng cho vừa khít khổ A4 ngang
    widths = [Cm(1.0), Cm(3.5), Cm(3.0), Cm(2.2), Cm(1.5), Cm(2.5), Cm(2.3), Cm(2.3), Cm(2.5), Cm(2.3), Cm(2.5)]
    
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        hdr_cells[i].paragraphs[0].runs[0].bold = True
        hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        hdr_cells[i].width = widths[i]
        
    # 5. Đổ dữ liệu từ bảng lọc vào Word
    stt = 1
    for idx, row in df.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(stt)
        row_cells[1].text = str(row.get('ho_ten', ''))
        row_cells[2].text = str(row.get('chuc_vu', ''))
        row_cells[3].text = str(row.get('ma_ngach', ''))
        row_cells[4].text = str(row.get('bac_luong', ''))
        
        # Xử lý hệ số hiện tại (Nối thêm % vượt khung nếu có)
        hs_ht = str(row.get('he_so_hien_tai', '')).strip()
        vk_ht = str(row.get('vuot_khung_hien_tai', '')).strip()
        if vk_ht and str(vk_ht).lower() not in ['none', 'nan', '']:
            hs_ht = f"{hs_ht} (Vượt khung {vk_ht})"
        row_cells[5].text = hs_ht
        
        row_cells[6].text = str(row.get('ngay_gan_nhat', ''))
        row_cells[7].text = str(row.get('bac_luong_moi', ''))
        
        # Xử lý hệ số mới (Nối thêm % vượt khung tương lai)
        hs_moi = str(row.get('he_so_moi', '')).strip()
        vk_moi = str(row.get('vuot_khung_moi', '')).strip()
        if vk_moi and str(vk_moi).lower() not in ['none', 'nan', '']:
            hs_moi = f"{hs_moi} (Vượt khung {vk_moi})"
        row_cells[8].text = hs_moi
        
        row_cells[9].text = str(row.get('ngay_du_kien', ''))
        row_cells[10].text = "Nâng lương thường xuyên" # Default ghi chú

        # Căn chỉnh lề chữ trong ô
        for i in range(11):
            row_cells[i].width = widths[i]
            row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            # Căn trái cho cột Tên, Chức vụ và Ghi chú để dễ đọc
            if i in [1, 2, 10]:
                row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        stt += 1

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()
    
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
        # TAB 1: BẢNG DỮ LIỆU & BỘ LỌC (BẢN TỐI THƯỢNG)
        # ==========================================
        with tab1:
            # 🌟 BỘ LỌC ĐA CHIỀU - THÊM CHỌN THÁNG 🌟
            c1, c2, c3, c4, c5 = st.columns([1.5, 1.2, 1.2, 0.8, 0.8])
            with c1:
                search = st.text_input("🔍 Tra cứu tên / chức vụ:", placeholder="Gõ tên tìm nhanh...")
            with c2:
                loc_thoi_gian = st.selectbox("⏳ Trạng thái đến hạn:", 
                                            ["Tất cả", "Trong tháng này", "Trong Quý này", "Trong 6 tháng tới", "Trong năm nay", "Đã quá hạn"])
            with c3:
                loai_ngay_loc = st.selectbox("📅 Báo cáo theo ngày nào?", 
                                            ["Ngày dự kiến (Tương lai)", "Ngày gần nhất (Đã nâng)"])
            with c4:
                danh_sach_nam = ["Tất cả"] + [str(year) for year in range(2020, 2036)]
                loc_nam = st.selectbox("🎯 Chọn Năm:", danh_sach_nam, index=0)
            with c5:
                # Thêm danh sách 12 tháng
                danh_sach_thang = ["Tất cả"] + [str(m) for m in range(1, 13)]
                loc_thang = st.selectbox("🌙 Chọn Tháng:", danh_sach_thang, index=0)

            # Xử lý Lọc Dữ Liệu
            df_display = df_calculated.copy()
            
            # 1. Lọc theo chữ (Tên, chức vụ)
            if search:
                mask = df_display.astype(str).apply(lambda x: x.str.contains(search, case=False, na=False)).any(axis=1)
                df_display = df_display[mask]
                
            # 2. Lọc theo Trạng thái (Chỉ áp dụng cho tương lai)
            if loc_thoi_gian != "Tất cả":
                today = datetime.now()
                df_display['ngay_dk_dt_temp'] = pd.to_datetime(df_display['ngay_du_kien'], format='%d/%m/%Y', errors='coerce')
                df_display['so_ngay'] = (df_display['ngay_dk_dt_temp'] - today).dt.days
                
                quy_hien_tai = (today.month - 1) // 3 + 1
                
                if loc_thoi_gian == "Trong tháng này":
                    df_display = df_display[(df_display['ngay_dk_dt_temp'].dt.month == today.month) & (df_display['ngay_dk_dt_temp'].dt.year == today.year)]
                elif loc_thoi_gian == "Trong Quý này":
                    df_display = df_display[(df_display['ngay_dk_dt_temp'].dt.quarter == quy_hien_tai) & (df_display['ngay_dk_dt_temp'].dt.year == today.year)]
                elif loc_thoi_gian == "Trong 6 tháng tới":
                    df_display = df_display[(df_display['so_ngay'] >= 0) & (df_display['so_ngay'] <= 180)]
                elif loc_thoi_gian == "Trong năm nay":
                    df_display = df_display[df_display['ngay_dk_dt_temp'].dt.year == today.year]
                elif loc_thoi_gian == "Đã quá hạn":
                    df_display = df_display[df_display['so_ngay'] < 0]
                    
                df_display = df_display.drop(columns=['ngay_dk_dt_temp', 'so_ngay'], errors='ignore')

            # 3. Lọc theo NĂM VÀ THÁNG CỤ THỂ 
            if loc_nam != "Tất cả" or loc_thang != "Tất cả":
                cot_ngay = 'ngay_du_kien' if loai_ngay_loc == "Ngày dự kiến (Tương lai)" else 'ngay_gan_nhat'
                df_display['temp_date_for_filter'] = pd.to_datetime(df_display[cot_ngay], format='%d/%m/%Y', errors='coerce')
                
                # Nếu có chọn Năm thì lọc theo Năm
                if loc_nam != "Tất cả":
                    df_display = df_display[df_display['temp_date_for_filter'].dt.year == int(loc_nam)]
                
                # Nếu có chọn Tháng thì lọc tiếp theo Tháng
                if loc_thang != "Tất cả":
                    df_display = df_display[df_display['temp_date_for_filter'].dt.month == int(loc_thang)]
                
                # Dọn rác
                df_display = df_display.drop(columns=['temp_date_for_filter'], errors='ignore')

            # --- KẾT THÚC ĐOẠN LỌC - HIỂN THỊ RA BẢNG DƯỚI ĐÂY ---
            def format_ma_ngach(val):
                if pd.isna(val) or val == "" or str(val).lower() == "nan":
                    return ""
                val_str = str(val).strip()
                if val_str.endswith(".0"):
                    val_str = val_str[:-2]
                return val_str

            df_display['ma_ngach'] = df_display['ma_ngach'].apply(format_ma_ngach)

            def color_status(val):
                val_str = str(val)
                if "Sắp đến hạn" in val_str or "Đã quá hạn" in val_str: return 'color: red; font-weight: bold'
                return 'color: green'

            st.caption("✍️ Sửa trực tiếp trên bảng. Sửa xong bấm LƯU để máy tự cộng cột Tương lai!")
            edited_df = st.data_editor(
                df_display.style.map(color_status, subset=['trang_thai']),
                num_rows="dynamic",
                column_config={
                    "ma_ngach": st.column_config.TextColumn("Mã ngạch")
                },
                disabled=["bac_luong_moi", "he_so_moi", "vuot_khung_moi", "ngay_du_kien", "trang_thai"],
                use_container_width=True,
                hide_index=True
            )
            
            # --- NÚT LƯU & TẢI ---
            col_luu, col_tai = st.columns(2)
            with col_luu:
                # ... (Giữ nguyên đoạn code của nút Lưu thay đổi ở đây) ...
                pass # Bạn không cần copy dòng pass này, chỉ cần giữ nguyên code nút Lưu của bạn

            with col_tai:
                out_df = edited_df.data if hasattr(edited_df, 'data') else edited_df
                
                # --- XUẤT EXCEL BẢN ĐẸP TỰ ĐỘNG ---
                import io
                from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
                from openpyxl.utils import get_column_letter
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    # Viết dữ liệu ra sheet
                    out_df.to_excel(writer, index=False, sheet_name='Theo_Doi_Luong')
                    worksheet = writer.sheets['Theo_Doi_Luong']
                    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                                         top=Side(style='thin'), bottom=Side(style='thin'))
                    
                    # 1. Trang điểm cho dòng Tiêu đề (Header)
                    header_fill = PatternFill(start_color="004B87", end_color="004B87", fill_type="solid")
                    for col_num, col_name in enumerate(out_df.columns, 1):
                        cell = worksheet.cell(row=1, column=col_num)
                        cell.fill = header_fill
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        cell.border = thin_border
                        
                    # 2. Trang điểm cho Data & Tự động căn chỉnh độ rộng cột
                    for i, col in enumerate(worksheet.columns, 1):
                        max_length = 0
                        column_letter = get_column_letter(i)
                        for cell in col:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass

                with col_tai_word:
                # Dữ liệu truyền vào chính là cái bảng sếp vừa lọc ở trên
                out_df_word = edited_df.data if hasattr(edited_df, 'data') else edited_df
                
                # Chạy máy in Word
                word_data = tao_file_word_dien_bien(out_df_word)
                
                st.download_button(
                    label="📝 Xuất Word Diễn biến lương",
                    data=word_data,
                    file_name=f"Dien_Bien_Luong_Thang_{datetime.now().strftime('%m_%Y')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="secondary" # Để nút màu xám nhạt cho đỡ rối với nút Excel
                )
                            # Kẻ viền cho mọi ô dữ liệu
                            if cell.row > 1:
                                cell.border = thin_border
                                cell.alignment = Alignment(vertical='center')
                                
                        # Giới hạn độ rộng cột không quá to, không quá nhỏ
                        worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 10), 35)
                        
                    worksheet.row_dimensions[1].height = 25 # Cột tiêu đề cao lên tí cho đẹp
                
                # Nút tải xuống giao diện Streamlit
                st.download_button(
                    label="📥 Xuất báo cáo EXCEL (Bản đẹp)",
                    data=buffer.getvalue(),
                    file_name=f"Danh_Sach_Nang_Luong_{datetime.now().strftime('%d%m%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary" # Nổi bật nút tải lên
                )

        # ==========================================
        # TAB 2: DASHBOARD BIỂU ĐỒ (Clone chuẩn Google Sheet của sếp)
        # ==========================================
        with tab2:
            st.subheader("📊 Bảng điều khiển (Dashboard) Tổng quan")
            if not df_calculated.empty:
                df_chart = df_calculated.copy()
                
                # --- HÀNG 1: 3 BIỂU ĐỒ (Chia 3 cột đều nhau) ---
                c1, c2, c3 = st.columns(3)
                
                # 1. Biểu đồ Mã ngạch lương hiện hưởng
                with c1:
                    df_ma = df_chart['ma_ngach'].fillna("Chưa có").astype(str).value_counts().reset_index()
                    df_ma.columns = ['Mã ngạch', 'Số lượng']
                    fig_ma = px.bar(df_ma, x='Mã ngạch', y='Số lượng', text='Số lượng')
                    fig_ma.update_traces(marker_color='#4A8af4', textposition='outside')
                    fig_ma.update_layout(
                        title={'text': "MÃ NGẠCH LƯƠNG HIỆN HƯỞNG", 'x': 0.5, 'font': {'color': 'blue', 'size': 14}},
                        paper_bgcolor='#f8e4b7', plot_bgcolor='#f8e4b7', # Màu nền vàng nhạt
                        xaxis={'type': 'category'}, xaxis_title=None, yaxis_title=None,
                        margin=dict(l=10, r=10, t=50, b=10),
                        height=350
                    )
                    st.plotly_chart(fig_ma, use_container_width=True)
                    
                # 2. Biểu đồ Ngạch lương
                with c2:
                    df_ngach = df_chart['ngach_luong'].fillna("Chưa có").astype(str).value_counts().reset_index()
                    df_ngach.columns = ['Ngạch', 'Số lượng']
                    fig_ngach = px.bar(df_ngach, x='Số lượng', y='Ngạch', orientation='h', text='Số lượng')
                    fig_ngach.update_traces(marker_color='#ba2812', textposition='outside')
                    fig_ngach.update_layout(
                        title={'text': "NGẠCH LƯƠNG", 'x': 0.5, 'font': {'color': 'blue', 'size': 14}},
                        paper_bgcolor='#f1f1f1', plot_bgcolor='#f1f1f1', # Màu nền xám nhạt
                        yaxis={'categoryorder': 'total ascending'}, xaxis_title=None, yaxis_title=None,
                        margin=dict(l=10, r=10, t=50, b=10),
                        height=350
                    )
                    st.plotly_chart(fig_ngach, use_container_width=True)
                    
                # 3. Biểu đồ Bậc lương hiện hưởng
                with c3:
                    df_bac = df_chart['bac_luong'].fillna("Chưa có").astype(str).value_counts().reset_index()
                    df_bac.columns = ['Bậc lương', 'Số lượng']
                    fig_bac = px.pie(df_bac, names='Bậc lương', values='Số lượng')
                    fig_bac.update_traces(textposition='inside', textinfo='percent+label')
                    fig_bac.update_layout(
                        title={'text': "BẬC LƯƠNG HIỆN HƯỞNG", 'x': 0.5, 'font': {'color': 'blue', 'size': 14}},
                        paper_bgcolor='#e2ccd9', plot_bgcolor='#e2ccd9', # Màu nền hồng nhạt
                        showlegend=False,
                        margin=dict(l=10, r=10, t=50, b=10),
                        height=350
                    )
                    st.plotly_chart(fig_bac, use_container_width=True)
                    
                st.write("---") # Đường kẻ phân cách
                
                # --- HÀNG 2: 1 BIỂU ĐỒ TRẢI DÀI (Full width) ---
                # 4. Biểu đồ Hệ số lương hiện tại
                df_heso = df_chart['he_so_hien_tai'].fillna("0").astype(str).value_counts().reset_index()
                df_heso.columns = ['Hệ số', 'Số lượng']
                fig_heso = px.bar(df_heso, x='Hệ số', y='Số lượng', text='Số lượng')
                fig_heso.update_traces(marker_color='#4A8af4', textposition='outside')
                fig_heso.update_layout(
                    title={'text': "HỆ SỐ LƯƠNG HIỆN TẠI", 'x': 0.5, 'font': {'color': 'blue', 'size': 16}},
                    paper_bgcolor='white', plot_bgcolor='white',
                    xaxis={'type': 'category'}, xaxis_title=None, yaxis_title=None,
                    margin=dict(l=20, r=20, t=50, b=20),
                    height=450
                )
                st.plotly_chart(fig_heso, use_container_width=True)
                
            else:
                st.info("Chưa có dữ liệu để vẽ biểu đồ.")

    except Exception as e:
        st.error(f"Hệ thống gặp sự cố: {e}")

if __name__ == "__main__":
    main()
