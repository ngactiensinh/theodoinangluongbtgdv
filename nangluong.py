import streamlit as st
import pandas as pd
from supabase import create_client
from datetime import datetime
from dateutil.relativedelta import relativedelta
import json
import io
import plotly.express as px
import plotly.graph_objects as go

# --- KHAI BÁO THƯ VIỆN WORD & EXCEL ---
try:
    from docx import Document
    from docx.shared import Cm, Pt
    from docx.enum.section import WD_ORIENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except:
    pass

try:
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    try:
        from openpyxl.utils import get_column_letter
    except ImportError:
        from openpyxl.utils.cell import get_column_letter
except:
    pass

# 1. CẤU HÌNH TRANG & CSS
st.set_page_config(page_title="Quản lý Lương Tuyên Quang", page_icon="📊", layout="wide")

st.markdown("""
    <style>
    .main-header {
        background: linear-gradient(135deg, #004B87 0%, #17a2b8 100%);
        padding: 30px; border-radius: 15px; color: white; text-align: center; margin-bottom: 25px;
    }
    .stDataFrame { border: 1px solid #e6e9ef; border-radius: 10px; }
    .metric-container {background-color: #ffffff; padding: 20px; border-radius: 15px; border: 2px solid #e6e9ef; text-align: center; height: 160px; display: flex; flex-direction: column; justify-content: center; margin-bottom: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);}
    .metric-label {font-size: 16px; color: #004B87; font-weight: bold; text-transform: uppercase; margin-bottom: 5px;}
    .metric-value {font-size: 50px; color: #C8102E; font-weight: 900; margin: 0;}
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header"><h1>📈 HỆ THỐNG QUẢN LÝ LƯƠNG 4.0</h1><p>Ban Tuyên giáo và Dân vận Tỉnh ủy Tuyên Quang</p></div>', unsafe_allow_html=True)

# 2. KẾT NỐI SUPABASE
try:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    supabase = create_client(url, key)
except Exception as e:
    st.error("Sếp Tuấn ơi, kiểm tra lại file Secrets trên Streamlit Cloud nhé!")
    st.stop()

# 3. CÁC HÀM XỬ LÝ DỮ LIỆU
def format_ma_ngach(val):
    if pd.isna(val) or val == "" or str(val).lower() == "nan": return ""
    val_str = str(val).strip()
    return val_str[:-2] if val_str.endswith(".0") else val_str

def tinh_toan_nang_luong(df):
    res = df.copy()
    if res.empty: return res
    today = datetime.now().date()
    for idx, row in res.iterrows():
        ngach = str(row.get('ngach_luong', '')).strip().upper()
        chuc_vu = str(row.get('chuc_vu', '')).strip().upper()
        bac_ht = str(row.get('bac_luong', '')).strip() 
        hs_str = str(row.get('he_so_hien_tai', '0')).replace(',', '.')
        try: hs_ht = float(hs_str)
        except: hs_ht = 0.0
        vk_ht = str(row.get('vuot_khung_hien_tai', 'None')).strip()
        ngay_ht_str = str(row.get('ngay_gan_nhat', ''))
        try:
            ngay_ht = datetime.strptime(ngay_ht_str, '%d/%m/%Y').date()
        except:
            for col in ['bac_luong_moi', 'he_so_moi', 'vuot_khung_moi', 'ngay_du_kien']: res.at[idx, col] = ""
            res.at[idx, 'trang_thai'] = "Chưa có ngày"
            continue
        is_vk = (vk_ht.lower() != 'none' and '%' in vk_ht)
        bac_moi, hs_moi, vk_moi, ngay_dk = bac_ht, hs_ht, vk_ht, ngay_ht
        if is_vk:
            vk_val = int(vk_ht.replace('%', '').strip())
            ngay_dk = ngay_ht + relativedelta(years=1); vk_moi = f"{vk_val + 1}%"
        else:
            try:
                if '/' in bac_ht: x, y = map(int, bac_ht.split('/'))
                else: x, y = int(bac_ht), 99
                if x >= y:
                    ngay_dk = ngay_ht + relativedelta(years=3); vk_moi = "5%"
                else:
                    bac_moi = f"{x+1}/{y}"
                    interval = 2 if any(k in ngach or k in chuc_vu for k in ['KẾ TOÁN VIÊN TRUNG CẤP', 'LÁI XE', 'PHỤC VỤ', 'VĂN THƯ']) else 3
                    delta = 0.34 if 'CVC' in ngach else (0.62 if 'CVCC' in ngach else 0.33)
                    ngay_dk = ngay_ht + relativedelta(years=interval); hs_moi = hs_ht + delta
            except: pass
        res.at[idx, 'bac_luong_moi'] = bac_moi
        res.at[idx, 'he_so_moi'] = f"{hs_moi:.2f}".replace('.', ',')
        res.at[idx, 'vuot_khung_moi'] = vk_moi
        res.at[idx, 'ngay_du_kien'] = ngay_dk.strftime('%d/%m/%Y')
        days_left = (ngay_dk - today).days
        if days_left < 0: res.at[idx, 'trang_thai'] = "Đã quá hạn"
        elif days_left <= 30: res.at[idx, 'trang_thai'] = "Sắp đến hạn (Tháng này)"
        elif days_left <= 90: res.at[idx, 'trang_thai'] = "Sắp đến hạn (Quý này)"
        else: res.at[idx, 'trang_thai'] = "Chưa đến hạn"
    return res.fillna("")

def tao_file_word_dien_bien(df, thang_chon="Tất cả", nam_chon="Tất cả"):
    from docx import Document
    from docx.shared import Cm, Pt
    from docx.enum.section import WD_ORIENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    for m in ['left', 'right', 'top', 'bottom']: setattr(section, f'{m}_margin', Cm(1.5))
    table_h = doc.add_table(rows=1, cols=2)
    table_h.columns[0].width, table_h.columns[1].width = Cm(10), Cm(16)
    cl = table_h.cell(0, 0).paragraphs[0]; cl.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cl.add_run("TỈNH UỶ TUYÊN QUANG\n").bold = True; cl.add_run("BAN TUYÊN GIÁO VÀ DÂN VẬN\n*").bold = True
    cr = table_h.cell(0, 1).paragraphs[0]; cr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cr.add_run("ĐẢNG CỘNG SẢN VIỆT NAM\n").bold = True
    p_t = doc.add_paragraph(); p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_t = p_t.add_run("\nBIỂU TỔNG HỢP DIỄN BIẾN LƯƠNG\n"); run_t.bold = True; run_t.font.size = Pt(14)
    txt_thang = thang_chon if thang_chon != "Tất cả" else datetime.now().strftime('%m')
    txt_nam = nam_chon if nam_chon != "Tất cả" else datetime.now().strftime('%Y')
    run_s = p_t.add_run(f"Ban Tuyên giáo và Dân vận Tỉnh uỷ tháng {txt_thang} năm {txt_nam}"); run_s.italic = True
    table = doc.add_table(rows=1, cols=11); table.style = 'Table Grid'
    headers = ['TT', 'Họ và tên', 'Chức vụ', 'Mã ngạch', 'Bậc', 'Hệ số HT', 'Ngày hưởng', 'Nâng bậc', 'Hệ số mới', 'Hưởng từ', 'Ghi chú']
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]; cell.text = h; cell.paragraphs[0].runs[0].bold = True; cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for idx, (index, r) in enumerate(df.iterrows(), 1):
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx); row_cells[1].text = str(r.get('ho_ten', '')); row_cells[2].text = str(r.get('chuc_vu', ''))
        row_cells[3].text = str(r.get('ma_ngach', '')); row_cells[4].text = str(r.get('bac_luong', ''))
        hs_ht = f"{r.get('he_so_hien_tai', '')}"
        if r.get('vuot_khung_hien_tai', '') not in ['', 'None', 'nan']: hs_ht += f" (VK {r.get('vuot_khung_hien_tai')})"
        row_cells[5].text = hs_ht; row_cells[6].text = str(r.get('ngay_gan_nhat', '')); row_cells[7].text = str(r.get('bac_luong_moi', ''))
        hs_m = f"{r.get('he_so_moi', '')}"
        if r.get('vuot_khung_moi', '') not in ['', 'None', 'nan']: hs_m += f" (VK {r.get('vuot_khung_moi')})"
        row_cells[8].text = hs_m; row_cells[9].text = str(r.get('ngay_du_kien', '')); row_cells[10].text = "Nâng lương TX"
        for i in range(11): row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    bio = io.BytesIO(); doc.save(bio); return bio.getvalue()

# 5. GIAO DIỆN CHÍNH
def main():
    try:
        res = supabase.table("theo_doi_luong").select("*").execute()
        df_base = pd.DataFrame(res.data) if res.data else pd.DataFrame(columns=["ho_ten", "ngay_gan_nhat", "ma_ngach"])
        df_calculated = tinh_toan_nang_luong(df_base)
        
        tab1, tab2 = st.tabs(["📋 Quản lý & Lọc Dữ liệu", "📊 Dashboard Thống kê"])
        
        with tab1:
            c1, c2, c3, c4, c5 = st.columns([1.5, 1.2, 1.2, 0.8, 0.8])
            search = c1.text_input("🔍 Tra cứu tên / chức vụ:", placeholder="Gõ tên...")
            loc_tg = c2.selectbox("⏳ Trạng thái:", ["Tất cả", "Trong tháng này", "Trong Quý này", "Trong 6 tháng tới", "Trong năm nay", "Đã quá hạn"])
            loai_ngay = c3.selectbox("📅 Loại ngày lọc:", ["Ngày dự kiến (Tương lai)", "Ngày gần nhất (Đã nâng)"])
            loc_nam = c4.selectbox("🎯 Năm:", ["Tất cả"] + [str(y) for y in range(2020, 2036)])
            loc_thang = c5.selectbox("🌙 Tháng:", ["Tất cả"] + [str(m) for m in range(1, 13)])
            
            df_display = df_calculated.copy()
            if search: df_display = df_display[df_display.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)]
            df_display['ngay_temp'] = pd.to_datetime(df_display['ngay_du_kien' if "dự kiến" in loai_ngay else 'ngay_gan_nhat'], format='%d/%m/%Y', errors='coerce')
            today = datetime.now()
            
            if loc_tg != "Tất cả":
                if loc_tg == "Trong tháng này": df_display = df_display[(df_display['ngay_temp'].dt.month == today.month) & (df_display['ngay_temp'].dt.year == today.year)]
                elif loc_tg == "Trong Quý này": df_display = df_display[(df_display['ngay_temp'].dt.quarter == (today.month-1)//3+1) & (df_display['ngay_temp'].dt.year == today.year)]
                elif loc_tg == "Trong năm nay": df_display = df_display[df_display['ngay_temp'].dt.year == today.year]
                elif loc_tg == "Đã quá hạn": df_display = df_display[(df_display['ngay_temp'].dt.date < today.date())]
            
            if loc_nam != "Tất cả": df_display = df_display[df_display['ngay_temp'].dt.year == int(loc_nam)]
            if loc_thang != "Tất cả": df_display = df_display[df_display['ngay_temp'].dt.month == int(loc_thang)]
            df_display['ma_ngach'] = df_display['ma_ngach'].apply(format_ma_ngach)
            
            edited_df = st.data_editor(
                df_display.style.map(lambda x: 'color:red; font-weight:bold' if any(s in str(x) for s in ["Sắp đến", "Quá hạn"]) else 'color:green', subset=['trang_thai']),
                num_rows="dynamic", use_container_width=True, hide_index=True,
                column_config={"ma_ngach": st.column_config.TextColumn("Mã ngạch")},
                disabled=["bac_luong_moi", "he_so_moi", "vuot_khung_moi", "ngay_du_kien", "trang_thai"]
            )
            
            st.write("---")
            col_l, col_e, col_w = st.columns(3)
            with col_l:
                if st.button("💾 Lưu thay đổi", use_container_width=True):
                    curr = edited_df.data if hasattr(edited_df, 'data') else edited_df
                    recs = []
                    for r in curr[curr['ho_ten'].str.strip().astype(bool)].to_dict(orient="records"):
                        recs.append({k: (None if pd.isna(v) or v == "" else v) for k, v in r.items() if k not in ["bac_luong_moi", "he_so_moi", "vuot_khung_moi", "ngay_du_kien", "trang_thai", "id", "ngay_temp"]})
                    supabase.table("theo_doi_luong").delete().neq("ho_ten", "Xóa_Hết").execute()
                    if recs: supabase.table("theo_doi_luong").insert(recs).execute()
                    st.success("Đã lưu!"); st.rerun()
            
            with col_e:
                try:
                    from openpyxl.styles import PatternFill, Font
                    try: from openpyxl.utils import get_column_letter
                    except: from openpyxl.utils.cell import get_column_letter
                    buf_e = io.BytesIO()
                    with pd.ExcelWriter(buf_e, engine='openpyxl') as wr:
                        edited_df.to_excel(wr, index=False, sheet_name='Luong')
                        ws = wr.sheets['Luong']
                        for col_num in range(1, len(edited_df.columns) + 1):
                            cell = ws.cell(row=1, column=col_num)
                            cell.fill = PatternFill(start_color="004B87", end_color="004B87", fill_type="solid")
                            cell.font = Font(bold=True, color="FFFFFF")
                            ws.column_dimensions[get_column_letter(col_num)].width = 20
                    st.download_button("📥 Xuất Excel", buf_e.getvalue(), "Bao_Cao.xlsx", use_container_width=True)
                except Exception as ex: st.warning(f"Lỗi vẽ Excel: {ex}")
            
            with col_w:
                st.download_button("📝 Xuất Word", tao_file_word_dien_bien(edited_df, loc_thang, loc_nam), "Dien_Bien.docx", use_container_width=True)

        with tab2:
            st.markdown("<h3 style='color:#004B87; text-align:center; margin-top: 10px; margin-bottom: 20px;'>📊 THỐNG KÊ TỔNG QUAN CHẤT LƯỢNG ĐỘI NGŨ</h3>", unsafe_allow_html=True)
            
            # --- TÍNH TOÁN SỐ LIỆU TỔNG ---
            tong_nv = len(df_calculated[df_calculated['ho_ten'].str.strip() != ""])
            sap_den_han = len(df_calculated[df_calculated['trang_thai'].str.contains("Sắp|quá", na=False, case=False)])
            
            # --- HÀNG 1: 2 Ô SỐ LIỆU CỰC TO ---
            r1c1, r1c2 = st.columns(2)
            with r1c1:
                st.markdown(f"<div class='metric-container'><div class='metric-label'>TỔNG SỐ CÁN BỘ</div><div class='metric-value' style='color:#004B87;'>{tong_nv}</div></div>", unsafe_allow_html=True)
            with r1c2:
                st.markdown(f"<div class='metric-container'><div class='metric-label'>SẮP ĐẾN HẠN / QUÁ HẠN LƯƠNG</div><div class='metric-value'>{sap_den_han}</div></div>", unsafe_allow_html=True)
            
            # --- HÀNG 2: 2 BIỂU ĐỒ (TRÒN & CỘT) ---
            cc1, cc2 = st.columns([1, 1.2]) 
            
            # Biểu đồ Tròn
            df_p = df_calculated['bac_luong'].value_counts().reset_index()
            df_p = df_p[df_p['bac_luong'].str.strip() != ""]
            
            fig_p = px.pie(df_p, names='bac_luong', values='count', hole=0.55, color_discrete_sequence=px.colors.sequential.Blues_r)
            fig_p.update_traces(textposition='inside', textinfo='percent+label', insidetextorientation='radial', marker=dict(line=dict(color='#FFFFFF', width=2)))
            fig_p.update_layout(
                title=dict(text="CƠ CẤU BẬC LƯƠNG", x=0.5, font=dict(size=16, color='#004B87', family='Arial')),
                showlegend=False, height=350, margin=dict(t=40, b=20, l=20, r=20),
                annotations=[dict(text=f"<b>{tong_nv}</b><br>Cán bộ", x=0.5, y=0.5, font_size=20, showarrow=False, font=dict(color='#C8102E'))]
            )
            with cc1: st.plotly_chart(fig_p, use_container_width=True)
                
            # Biểu đồ Cột
            df_b = df_calculated['ma_ngach'].value_counts().reset_index()
            df_b = df_b[df_b['ma_ngach'].str.strip() != ""]
            
            fig_b = px.bar(df_b, x='ma_ngach', y='count', text='count', color='count', color_continuous_scale='Blues')
            fig_b.update_traces(textposition='outside', textfont_size=14, marker_line_color='rgb(8,48,107)', marker_line_width=1.5, opacity=0.9)
            fig_b.update_layout(
                title=dict(text="PHÂN BỔ THEO MÃ NGẠCH", x=0.5, font=dict(size=16, color='#004B87', family='Arial')),
                xaxis_title="", yaxis_title="", xaxis_tickangle=-30, coloraxis_showscale=False, plot_bgcolor='rgba(0,0,0,0)',
                height=350, margin=dict(t=40, b=20, l=20, r=20)
            )
            fig_b.update_yaxes(showgrid=True, gridwidth=1, gridcolor='#e6e6e6', showticklabels=False)
            with cc2: st.plotly_chart(fig_b, use_container_width=True)

    except Exception as e:
        st.error(f"Lỗi hệ thống: {e}")

if __name__ == "__main__":
    main()
