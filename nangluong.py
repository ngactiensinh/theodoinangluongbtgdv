import streamlit as st
import pandas as pd
from supabase import create_client

def main():
    st.set_page_config(page_title="Quan ly Luong", layout="wide")

    # Header dung Markdown co ban de tranh loi encode
    st.title("HE THONG QUAN LY LUONG 4.0")
    st.write("Ban Tuyen giao va Dan van Tinh uy Tuyen Quang")

    try:
        # Lay thong tin tu Secrets (Cach nay giup tranh loi ASCII truc tiep trong code)
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
        
        supabase = create_client(url, key)
        
        # Lay du lieu
        res = supabase.table("theo_doi_luong").select("*").execute()
        
        if res.data:
            df = pd.DataFrame(res.data)
            
            # Chatbot thong bao
            st.divider()
            st.info("Chao Ban Tuan! He thong da ket noi thanh cong.")
            
            # Tim cot thong minh
            col_status = next((c for c in df.columns if 'trang' in c.lower()), None)
            if col_status:
                df[col_status] = df[col_status].astype(str)
                sap_den_han = df[df[col_status].str.contains("Sap den han", na=False)]
                if not sap_den_han.empty:
                    st.warning(f"Co {len(sap_den_han)} dong chi sap den han nang luong.")
            
            # Tim kiem va hien thi
            search = st.text_input("Tim ten can bo:")
            col_name = next((c for c in df.columns if 'ho' in c.lower() or 'ten' in c.lower()), df.columns[0])
            
            if search:
                df = df[df[col_name].astype(str).str.contains(search, case=False, na=False)]
            
            st.dataframe(df, use_container_width=True)
        else:
            st.write("Chua co du lieu.")

    except Exception as e:
        st.error("Loi ket noi. Sep hay kiem tra lai phan Secrets nhe!")
        # Khong in e ra de tranh loi encode tiep tuc
        print(f"Log: {str(e)}")

if __name__ == "__main__":
    main()
