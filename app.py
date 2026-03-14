import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import os
from datetime import datetime

# Cấu hình trang
st.set_page_config(page_title="Bamboo Gendec System", layout="wide")

# --- XỬ LÝ ĐƯỜNG DẪN FILE ---
base_dir = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_FILE = os.path.join(base_dir, "template.docx")

# CSS fix Dark Mode và làm đẹp UI
st.markdown("""
    <style>
    .main { background-color: transparent; }
    [data-testid="stMetric"] {
        background-color: #f0f2f6;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    [data-testid="stMetricLabel"] { color: #31333F !important; }
    [data-testid="stMetricValue"] { color: #1a73e8 !important; font-weight: bold; }
    .crew-box {
        background-color: #ffffff;
        color: #1a1c21 !important;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #1a73e8;
        line-height: 1.6;
    }
    .crew-box ul li {
        color: #1a1c21 !important;
        list-style-type: none;
        border-bottom: 1px solid #eee;
        padding: 5px 0;
    }
    h3 { color: #4A90E2 !important; }
    </style>
    """, unsafe_allow_html=True)

st.title("✈️ Bamboo Airways Gendec Generator")

if not os.path.exists(TEMPLATE_FILE):
    st.error(f"⚠️ Không tìm thấy file template tại: {TEMPLATE_FILE}")
else:
    with st.sidebar:
        st.header("Cài đặt")
        uploaded_excel = st.file_uploader("1. Upload Excel Data", type=["xlsx"])
        st.info("File template đã được tích hợp sẵn.")

    if uploaded_excel:
        # Đọc file (header ở dòng 3 -> index 2)
        df = pd.read_excel(uploaded_excel, header=2)
        
        # Làm sạch tên cột
        df.columns = [str(c).strip() for c in df.columns]

        # --- DYNAMIC COLUMN DETECTION ---
        # Tìm cột chứa từ khóa "JumpSeaters" và "Crew"
        js_col = None
        crew_col = None
        
        for col in df.columns:
            if "JumpSeaters" in col:
                js_col = col
            if "Crew" in col and "Crew #" not in col: # Tránh nhầm với cột Crew #
                crew_col = col

        # Ô nhập số hiệu chuyến bay
        search_flt = st.text_input("🔍 Nhập số hiệu chuyến bay (ví dụ: 101, 102, 208...)", "")

        if search_flt:
            # Tìm dòng chứa chuyến bay
            target_row = df[df['FLT'].astype(str).str.contains(search_flt)]

            if not target_row.empty:
                start_idx = target_row.index[0]

                # --- TRÍCH XUẤT DỮ LIỆU CƠ BẢN ---
                flt_val = str(df.loc[start_idx, 'FLT'])
                reg_val = str(df.loc[start_idx, 'REG'])
                dep_val = str(df.loc[start_idx, 'DEP'])
                arr_val = str(df.loc[start_idx, 'ARR'])

                # Format DATE (DDMMMYY)
                raw_date = df.loc[start_idx, 'DATE']
                try:
                    if isinstance(raw_date, datetime):
                        date_val = raw_date.strftime('%d%b%y').upper()
                    else:
                        date_obj = pd.to_datetime(raw_date, dayfirst=True)
                        date_val = date_obj.strftime('%d%b%y').upper()
                except:
                    date_val = str(raw_date)

                # --- QUÉT CREW VÀ JUMPSEATERS ---
                crew_list = []
                jump_seaters_list = []
                
                for i in range(start_idx, len(df)):
                    # Dừng lại nếu gặp chuyến bay tiếp theo
                    if i > start_idx and pd.notna(df.loc[i, 'FLT']):
                        break
                    
                    # Lấy Crew (Dò theo tên cột đã detect)
                    if crew_col and pd.notna(df.loc[i, crew_col]):
                        val = str(df.loc[i, crew_col]).strip()
                        if val.lower() != "crew": # Tránh lấy trúng header nếu có
                            crew_list.append(val)
                    
                    # Lấy JumpSeaters (Dò theo tên cột đã detect)
                    if js_col and pd.notna(df.loc[i, js_col]):
                        val = str(df.loc[i, js_col]).strip()
                        # Loại bỏ chữ "name" nếu nó nằm ở hàng ngay dưới tiêu đề
                        if val.lower() not in ["name", "jumpseaters"]:
                            jump_seaters_list.append(val)

                # --- UI PREVIEW ---
                st.markdown("---")
                st.subheader(f"📊 Preview: Chuyến bay QH{flt_val}")

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("DATE", date_val)
                c2.metric("REG", reg_val)
                c3.metric("DEP", dep_val)
                c4.metric("ARR", arr_val)

                st.write("")
                col_left, col_right = st.columns(2)

                with col_left:
                    st.markdown("### 👨‍✈️ Crew List")
                    if crew_list:
                        crew_html = "".join([f"<li>{c}</li>" for c in crew_list])
                        st.markdown(f'<div class="crew-box"><ul>{crew_html}</ul></div>', unsafe_allow_html=True)
                    else:
                        st.warning("Không tìm thấy dữ liệu Crew.")

                with col_right:
                    st.markdown("### 💺 JumpSeaters")
                    if jump_seaters_list:
                        js_html = "".join([f"<li>{j}</li>" for j in jump_seaters_list])
                        st.markdown(f'<div class="crew-box" style="border-left-color: #f4b400;"><ul>{js_html}</ul></div>', unsafe_allow_html=True)
                    else:
                        st.info("Không có JumpSeaters.")

                # --- GENERATE WORD ---
                st.write("")
                if st.button("🚀 XUẤT FILE WORD NGAY", use_container_width=True):
                    doc = DocxTemplate(TEMPLATE_FILE)
                    context = {
                        'FLT': flt_val, 'REG': reg_val, 'DEP': dep_val,
                        'ARR': arr_val, 'DATE': date_val,
                        'Crew': "\n".join(crew_list),
                        'JumpSeaters': "\n".join(jump_seaters_list)
                    }
                    doc.render(context)
                    bio = io.BytesIO()
                    doc.save(bio)
                    bio.seek(0)

                    st.download_button(
                        label=f"📥 Tải xuống: GD_QH{flt_val}.docx",
                        data=bio,
                        file_name=f"GD_QH{flt_val}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
            else:
                st.warning(f"❌ Không tìm thấy chuyến bay: {search_flt}")
