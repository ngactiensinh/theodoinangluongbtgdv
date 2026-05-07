import streamlit as st
import pandas as pd
import json
import os
import plotly.express as px
import plotly.graph_objects as go
import io
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Hệ thống Báo cáo TGDV", page_icon="📊", layout="wide")

# ==========================================
# 1. CẤU HÌNH DỮ LIỆU & DANH SÁCH ĐƠN VỊ
# ==========================================
DATA_FILE = "dulieu_baocao.json"
CONFIG_FILE = "config_donvi.json"

DEFAULT_UNITS = [
    "Đảng ủy Công an tỉnh", "Đảng ủy Quân sự tỉnh", "Đảng ủy các cơ quan Đảng tỉnh", "Đảng ủy Ủy ban nhân dân tỉnh",
    "Phường Mỹ Lâm", "Phường Minh Xuân", "Phường Nông Tiến", "Phường An Tường", "Phường Bình Thuận", "Phường Hà Giang 1", "Phường Hà Giang 2",
    "Xã Thượng Lâm", "Xã Lâm Bình", "Xã Minh Quang", "Xã Bình An", "Xã Côn Lôn", "Xã Yên Hoa", "Xã Thượng Nông", "Xã Hồng Thái", "Xã Nà Hang", "Xã Tân Mỹ", "Xã Yên Lập", "Xã Tân An", "Xã Chiêm Hóa", "Xã Hòa An", "Xã Kiên Đài", "Xã Tri Phú", "Xã Kim Bình", "Xã Yên Nguyên", "Xã Yên Phú", "Xã Bạch Xa", "Xã Phù Lưu", "Xã Hàm Yên", "Xã Bình Xa", "Xã Thái Sơn", "Xã Thái Hòa", "Xã Hùng Lợi", "Xã Trung Sơn", "Xã Thái Bình", "Xã Tân Long", "Xã Xuân Vân", "Xã Lực Hành", "Xã Yên Sơn", "Xã Nhữ Khê", "Xã Tân Trào", "Xã Minh Thanh", "Xã Sơn Dương", "Xã Bình Ca", "Xã Tân Thanh", "Xã Sơn Thủy", "Xã Phú Lương", "Xã Trường Sinh", "Xã Hồng Sơn", "Xã Đông Thọ",
    "Xã Lũng Cú", "Xã Đồng Văn", "Xã Sà Phìn", "Xã Phố Bảng", "Xã Lũng Phìn", "Xã Sủng Máng", "Xã Sơn Vĩ", "Xã Mèo Vạc", "Xã Khâu Vai", "Xã Niêm Sơn", "Xã Tát Ngà", "Xã Thắng Mố", "Xã Bạch Đích", "Xã Yên Minh", "Xã Mậu Duệ", "Xã Du Già", "Xã Đường Thượng", "Xã Lùng Tám", "Xã Cán Tỷ", "Xã Nghĩa Thuận", "Xã Quản Bạ", "Xã Tùng Vài", "Xã Yên Cường", "Xã Đường Hồng", "Xã Bắc Mê", "Xã Minh Ngọc", "Xã Ngọc Đường", "Xã Lao Chải", "Xã Thanh Thủy", "Xã Phú Linh", "Xã Linh Hồ", "Xã Bạch Ngọc", "Xã Vị Xuyên", "Xã Việt Lâm", "Xã Tân Quang", "Xã Đồng Tâm", "Xã Liên Hiệp", "Xã Bằng Hành", "Xã Bắc Quang", "Xã Hùng An", "Xã Vĩnh Tuy", "Xã Đồng Yên", "Xã Tiên Yên", "Xã Xuân Giang", "Xã Bằng Lang", "Xã Yên Thành", "Xã Quang Bình", "Xã Tân Trịnh", "Xã Thông Nguyên", "Xã Hồ Thầu", "Xã Nậm Dịch", "Xã Tân Tiến", "Xã Hoàng Su Phì", "Xã Thàng Tín", "Xã Bản Máy", "Xã Pờ Ly Ngài", "Xã Xín Mần", "Xã Pà Vầy Sủ", "Xã Nấm Dẩn", "Xã Trung Thịnh", "Xã Khuôn Lùng", "Xã Trung Hà", "Xã Kiến Thiết", "Xã Hùng Đức", "Xã Minh Sơn", "Xã Minh Tân", "Xã Thuận Hòa", "Xã Tùng Bá", "Xã Thượng Sơn", "Xã Cao Bồ", "Xã Ngọc Long", "Xã Giáp Trung", "Xã Tiên Nguyên", "Xã Quảng Nguyên"
]

DANH_SACH_THANG = [f"Tháng {i}" for i in range(1, 13)]

DICT_DICH_THUAT = {
    "don_vi": "Đơn vị báo cáo", "nguoi_bao_cao": "Người BC / SĐT", "ky_bao_cao": "Kỳ báo cáo",
    "ld_vanban": "Số VB cấp ủy ban hành", "ld_thammuu": "Số VB tham mưu cấp trên", "ld_cuochop": "Số cuộc họp, hội nghị",
    "nq_hoinghi": "Số hội nghị NQ", "nq_nguoi": "Số người tham gia NQ", "nq_vanban": "Số VB đã triển khai", "nq_tyle": "Tỷ lệ ĐV tham gia (%)",
    "tt_tinbai": "Số tin, bài, pano", "tt_loa": "Số lượt loa truyền thanh", "tt_buoi": "Số buổi TT miệng", "tt_nguoi": "Số người nghe TT", "tt_mxh_bai": "Số bài trên MXH/Cổng TT", "tt_mxh_tuongtac": "Lượt tương tác MXH",
    "dl_baocao": "Số BC DLXH gửi đi", "dl_vande": "Số vấn đề nổi cộm", "dl_xuly": "Số vụ việc đã xử lý",
    "kg_hoatdong": "Số HĐ Văn hóa-Văn nghệ", "kg_chuongtrinh": "Số CT tuyên truyền GD", "kg_lop": "Số buổi Y tế/Môi trường",
    "kg_bd_chuyennghiep": "Số buổi BDNT chuyên nghiệp", "kg_bd_quanchung": "Số buổi BDNT quần chúng", 
    "kg_clb_thanhlap": "Số CLB VH-NT thành lập", "kg_clb_thanhvien": "Số thành viên CLB", 
    "kg_hd_vhtt": "Số HĐ Lễ hội, Thể thao", "kg_khokhan": "Khó khăn Khoa giáo, VH-VN",
    "dv_mh_dangky": "Mô hình DVK đăng ký", "dv_mh_hieuqua": "Mô hình DVK hiệu quả", "dv_mh_moi": "Mô hình mới trong kỳ", "dv_cuocvandong": "Số cuộc vận động, TT", "dv_nguoithamgia": "Số lượt người tham gia", "dv_tiepxuc": "Số buổi đối thoại Nhân dân",
    "nv_duocgiao": "Nhiệm vụ trọng tâm giao", "nv_hoanthanh": "Nhiệm vụ TT hoàn thành", "nv_dangtrienkhai": "Nhiệm vụ đang triển khai", "nv_ketqua": "Kết quả thí điểm nổi bật",
    "bd_tinbai": "Số tin bài CĐS", "bd_cuocthi": "Số cuộc thi CĐS", "kq_tocongnghe": "Số Tổ công nghệ số",
    "ts_chibo": "Tổng số Chi bộ", "kq_chibo_cd": "Số CB SH chuyên đề", "kq_chibo_sotay": "Số CB dùng Sổ tay ĐV",
    "ts_cbccvc": "Tổng số CBCCVC", "kq_cb_ai": "Số CB biết dùng AI", "kq_cb_khoahoc": "Số CB học xong CĐS",
    "ts_nd_truongthanh": "Tổng ND trưởng thành", "kq_nd_kynang": "Số ND có Kỹ năng số", "kq_nd_vneid": "Số ND phổ cập VNeID",
    "kq_nd_smartphone": "Số ND dùng Smartphone", "kq_lop_nd": "Số buổi học cộng đồng",
    "tl_mohinh": "Mô hình hay, sáng tạo", "tl_khokhan": "Khó khăn, vướng mắc chung"
}

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f: return json.load(f)
    return []

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=4)

def load_units():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f: return json.load(f)
    return DEFAULT_UNITS

def save_units(units):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f: json.dump(units, f, ensure_ascii=False, indent=4)

def get_months_for_filter(filter_type):
    if filter_type == "Quý I": return ["Tháng 1", "Tháng 2", "Tháng 3"]
    if filter_type == "Quý II": return ["Tháng 4", "Tháng 5", "Tháng 6"]
    if filter_type == "Quý III": return ["Tháng 7", "Tháng 8", "Tháng 9"]
    if filter_type == "Quý IV": return ["Tháng 10", "Tháng 11", "Tháng 12"]
    if filter_type == "6 Tháng Đầu Năm": return [f"Tháng {i}" for i in range(1, 7)]
    if filter_type == "6 Tháng Cuối Năm": return [f"Tháng {i}" for i in range(7, 13)]
    if filter_type == "9 Tháng": return [f"Tháng {i}" for i in range(1, 10)]
    return DANH_SACH_THANG

# ==========================================
# 2. GIAO DIỆN & ĐĂNG NHẬP
# ==========================================
st.markdown("""
<style>
    .stApp { background-color: #f4f6f9; }
    .main-header {color: #004B87; font-weight: 900; text-align: center; text-transform: uppercase; margin-bottom: 25px;}
    .stButton>button {background-color: #004B87; color: white; font-weight: bold; border-radius: 6px;}
    [data-testid="stForm"] {background-color: #ffffff; padding: 25px; border-radius: 12px; border-top: 5px solid #004B87;}
    .metric-container {background-color: #ffffff; padding: 30px; border-radius: 15px; border: 1px solid #e0e0e0; text-align: center; height: 400px; display: flex; flex-direction: column; justify-content: center;}
    .metric-label {font-size: 16px; color: #004B87; font-weight: bold; text-transform: uppercase; margin-bottom: 20px;}
    .metric-value {font-size: 80px; color: #C8102E; font-weight: 900; margin: 0;}
</style>
""", unsafe_allow_html=True)

if "role" not in st.session_state: st.session_state.role = None
if st.session_state.role is None:
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        with st.form("login"):
            st.markdown("<h2 style='text-align: center; color: #004B87;'>ĐĂNG NHẬP BÁO CÁO TGDV</h2>", unsafe_allow_html=True)
            pwd = st.text_input("Mật khẩu:", type="password")
            if st.form_submit_button("Vào hệ thống", use_container_width=True):
                if pwd == "TGDV@2026": st.session_state.role = "user"; st.rerun()
                elif pwd == "admin123": st.session_state.role = "admin"; st.rerun()
                else: st.error("Sai mật khẩu!")
    st.stop()

# ==========================================
# 3. NỘI DUNG CHÍNH
# ==========================================
st.markdown("<h1 class='main-header'>HỆ THỐNG THU THẬP BÁO CÁO CƠ SỞ</h1>", unsafe_allow_html=True)

if st.session_state.role == "admin":
    tabs = st.tabs(["📝 NHẬP BÁO CÁO", "📊 THỐNG KÊ & BIỂU ĐỒ", "⚙️ QUẢN TRỊ ADMIN"])
    tab_nhap, tab_bieudo, tab_admin = tabs[0], tabs[1], tabs[2]
else:
    tabs = st.tabs(["📝 NHẬP BÁO CÁO"])
    tab_nhap = tabs[0]

# --- TAB NHẬP BÁO CÁO ---
with tab_nhap:
    with st.form("form_nhap"):
        c1, c2, c3 = st.columns([2, 1.5, 1.5])
        don_vi = c1.selectbox("🏢 Đơn vị:", load_units(), index=None, placeholder="Chọn đơn vị...")
        nguoi_bc = c2.text_input("👤 Người báo cáo / SĐT:")
        ky_bc = c3.selectbox("🗓️ Tháng báo cáo:", DANH_SACH_THANG, index=None)
        
        with st.expander("8. CHUYÊN ĐỀ: BÌNH DÂN HỌC VỤ SỐ", expanded=True):
            st.markdown("#### - Đối với Chi bộ & Cán bộ:")
            ts_chibo = st.number_input("Tổng số Chi bộ", min_value=0, value=0)
            kq_chibo_cd = st.number_input("Số Chi bộ đã sinh hoạt chuyên đề Kỹ năng số", min_value=0, value=0)
            ts_cbccvc = st.number_input("Tổng số CB, CC, VC", min_value=0, value=0)
            kq_cb_ai = st.number_input("Số Cán bộ biết và ứng dụng AI", min_value=0, value=0)
        
        st.markdown("*(Sếp nhập thêm các mục khác nếu cần, ở đây bây bề tập trung vào 4 nội dung sếp yêu cầu)*")
        
        if st.form_submit_button("🚀 GỬI BÁO CÁO", type="primary", use_container_width=True):
            if not don_vi or not ky_bc: st.error("Thiếu thông tin!")
            else:
                data = load_data()
                rec = {"don_vi": don_vi, "ky_bao_cao": ky_bc, "ts_chibo": ts_chibo, "kq_chibo_cd": kq_chibo_cd, "ts_cbccvc": ts_cbccvc, "kq_cb_ai": kq_cb_ai}
                data = [d for d in data if not (d['don_vi'] == don_vi and d['ky_bao_cao'] == ky_bc)]
                data.append(rec); save_data(data); st.success("Đã gửi thành công!")

# --- TAB THỐNG KÊ & BIỂU ĐỒ (DASHBOARD 2x2 CỦA TUẤN ĐẸP ZAI) ---
if st.session_state.role == "admin":
    with tab_bieudo:
        data = load_data()
        if not data: st.warning("Chưa có số liệu.")
        else:
            df_raw = pd.DataFrame(data)
            c_f1, c_f2 = st.columns(2)
            loai_bc = c_f1.selectbox("Kỳ tổng hợp:", ["Tháng", "Quý I", "Quý II", "Quý III", "Quý IV", "Cả Năm"])
            if loai_bc == "Tháng":
                th_bc = c_f2.selectbox("Chọn tháng:", DANH_SACH_THANG)
                df = df_raw[df_raw['ky_bao_cao'] == th_bc]
            else:
                df = df_raw[df_raw['ky_bao_cao'].isin(get_months_for_filter(loai_bc))]

            if df.empty: st.warning("Không có dữ liệu.")
            else:
                # Tính toán số liệu tổng
                df_sum = df.groupby('don_vi').sum(numeric_only=True).reset_index()
                
                sum_chibo = int(df_sum['ts_chibo'].sum())
                sum_chibo_cd = int(df_sum['kq_chibo_cd'].sum())
                sum_cb = int(df_sum['ts_cbccvc'].sum())
                sum_cb_ai = int(df_sum['kq_cb_ai'].sum())

                st.markdown("<br>", unsafe_allow_html=True)
                
                # --- HÀNG 1: CHI BỘ ---
                r1c1, r1c2 = st.columns([1, 1.2])
                with r1c1:
                    st.markdown(f"""
                    <div class='metric-container'>
                        <div class='metric-label'>TỔNG SỐ CHI BỘ</div>
                        <div class='metric-value'>{sum_chibo:,}</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with r1c2:
                    # Biểu đồ Donut Chi bộ
                    df_p1 = pd.DataFrame({
                        "Trạng thái": ["Đã sinh hoạt CĐ", "Chưa sinh hoạt"],
                        "Số lượng": [sum_chibo_cd, max(0, sum_chibo - sum_chibo_cd)]
                    })
                    fig1 = px.pie(df_p1, names='Trạng thái', values='Số lượng', hole=0.6,
                                  color_discrete_sequence=['#004B87', '#E6E6E6'])
                    fig1.update_traces(textposition='inside', textinfo='percent+label')
                    fig1.update_layout(title=dict(text="CƠ CẤU CHI BỘ SINH HOẠT CHUYÊN ĐỀ KỸ NĂNG SỐ", x=0.5),
                                      showlegend=False, height=400, margin=dict(t=50, b=20, l=20, r=20))
                    # Thêm số % vào giữa
                    percent_cb = (sum_chibo_cd / sum_chibo * 100) if sum_chibo > 0 else 0
                    fig1.add_annotation(text=f"<b>{percent_cb:.1f}%</b>", x=0.5, y=0.5, font_size=24, showarrow=False)
                    st.plotly_chart(fig1, use_container_width=True)

                st.markdown("<br>", unsafe_allow_html=True)

                # --- HÀNG 2: CÁN BỘ ---
                r2c1, r2c2 = st.columns([1, 1.2])
                with r2c1:
                    st.markdown(f"""
                    <div class='metric-container'>
                        <div class='metric-label'>TỔNG SỐ CÁN BỘ, CÔNG CHỨC, VIÊN CHỨC</div>
                        <div class='metric-value'>{sum_cb:,}</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with r2c2:
                    # Biểu đồ Donut Cán bộ
                    df_p2 = pd.DataFrame({
                        "Trạng thái": ["Biết & Ứng dụng AI", "Chưa ứng dụng"],
                        "Số lượng": [sum_cb_ai, max(0, sum_cb - sum_cb_ai)]
                    })
                    fig2 = px.pie(df_p2, names='Trạng thái', values='Số lượng', hole=0.6,
                                  color_discrete_sequence=['#C8102E', '#E6E6E6'])
                    fig2.update_traces(textposition='inside', textinfo='percent+label')
                    fig2.update_layout(title=dict(text="TỈ LỆ CÁN BỘ BIẾT VÀ ỨNG DỤNG AI", x=0.5),
                                      showlegend=False, height=400, margin=dict(t=50, b=20, l=20, r=20))
                    # Thêm số % vào giữa
                    percent_ai = (sum_cb_ai / sum_cb * 100) if sum_cb > 0 else 0
                    fig2.add_annotation(text=f"<b>{percent_ai:.1f}%</b>", x=0.5, y=0.5, font_size=24, showarrow=False, font_color="#C8102E")
                    st.plotly_chart(fig2, use_container_width=True)

# --- TAB ADMIN ---
    with tab_admin:
        st.write("#### ⚙️ QUẢN TRỊ DỮ LIỆU")
        if st.button("🔥 XÓA TOÀN BỘ DỮ LIỆU TEST"):
            save_data([]); st.success("Sạch bong kin kít!"); st.rerun()

if __name__ == "__main__":
    pass
