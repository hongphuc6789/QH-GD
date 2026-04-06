import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import os
from datetime import datetime

# 1. Cấu hình trang
st.set_page_config(page_title="Bamboo Airways Gendec System", layout="wide")

# --- XỬ LÝ ĐƯỜNG DẪN FILE TEMPLATE ---
base_dir = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_FILE = os.path.join(base_dir, "template.docx")

# 2. CSS Custom
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
    .info-box {
        background-color: #ffffff !important;
        color: #1a1c21 !important;
        padding: 20px;
        border-radius: 10px;
        border-left: 6px solid #1a73e8;
        line-height: 1.6;
        box-shadow: 0 2px 10px rgba(0,0,0,0.05);
    }
    .info-box ul { margin: 0; padding-left: 20px; }
    .info-box ul li {
        color: #1a1c21 !important;
        list-style-type: disc;
        border-bottom: 1px solid #f0f0f0;
        padding: 8px 0;
    }
    h3 { color: #4A90E2 !important; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

st.title("✈️ Bamboo Airways Gendec Generator")

if not os.path.exists(TEMPLATE_FILE):
    st.error(f"⚠️ Không tìm thấy file template tại: {TEMPLATE_FILE}")
else:
    with st.sidebar:
        st.header("Cài đặt")
        uploaded_excel = st.file_uploader("1. Upload file Excel Data", type=["xlsx"])
        st.info("💡 File 'template.docx' đã được tích hợp sẵn trên hệ thống.")

    if uploaded_excel:
        # --- SỬA LỖI TẠI ĐÂY ---
        # Đọc toàn bộ dưới dạng string để tránh Pandas tự parse ngày tháng sai định dạng
        df = pd.read_excel(uploaded_excel, header=2, dtype=str)
        
        df.columns = [str(c).strip() for c in df.columns]

        js_col = None
        crew_col = None
        for col in df.columns:
            if "JumpSeaters" in col: js_col = col
            if "Crew" in col and "Crew #" not in col: crew_col = col

        search_flt = st.text_input("🔍 Nhập số hiệu chuyến bay (ví dụ: 102, 208, 147...)", "")

        if search_flt:
            target_row = df[df['FLT'].astype(str).str.contains(search_flt)]

            if not target_row.empty:
                start_idx = target_row.index[0]

                # Clean FLT
                raw_flt = df.loc[start_idx, 'FLT']
                flt_val = str(raw_flt).split('.')[0] if '.' in str(raw_flt) else str(raw_flt)

                reg_val = str(df.loc[start_idx, 'REG'])
                dep_val = str(df.loc[start_idx, 'DEP'])
                arr_val = str(df.loc[start_idx, 'ARR'])

                # --- XỬ LÝ DATE CHUẨN XÁC ---
                raw_date = df.loc[start_idx, 'DATE']
                try:
                    # Vì đã ép kiểu string ở trên, ta parse với dayfirst=True
                    # Xử lý trường hợp chuỗi có giờ đi kèm (ví dụ: "2026-05-04 00:00:00")
                    clean_date_str = str(raw_date).split(' ')[0]
                    
                    # Ưu tiên parse Ngày trước Tháng
                    date_obj = pd.to_datetime(clean_date_str, dayfirst=True, errors='coerce')
                    
                    if pd.isna(date_obj):
                        # Nếu vẫn lỗi (do file format lạ), giữ nguyên text
                        date_val = clean_date_str
                    else:
                        date_val = date_obj.strftime('%d%b%y').upper() # Kết quả: 05APR26
                except:
                    date_val = str(raw_date).replace('.0', '')

                # --- QUÉT CREW VÀ JUMPSEATERS ---
                crew_list = []
                jump_seaters_list = []
                
                for i in range(start_idx, len(df)):
                    if i > start_idx and pd.notna(df.loc[i, 'FLT']):
                        break
                    
                    if crew_col and pd.notna(df.loc[i, crew_col]):
                        c_val = str(df.loc[i, crew_col]).strip()
                        if c_val.lower() not in ["crew", "name", "nan"]:
                            crew_list.append(c_val)
                    
                    if js_col and pd.notna(df.loc[i, js_col]):
                        j_val = str(df.loc[i, js_col]).strip()
                        if j_val.lower() not in ["jumpseaters", "name", "nan"]:
                            jump_seaters_list.append(j_val)

                # --- GIAO DIỆN PREVIEW ---
                st.markdown("---")
                st.subheader(f"📊 Thông tin chuyến bay QH{flt_val}")

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("DATE", date_val)
                m2.metric("REGISTRATION", reg_val)
                m3.metric("FROM", dep_val)
                m4.metric("TO", arr_val)

                st.write("")
                col_left, col_right = st.columns(2)

                with col_left:
                    st.markdown("### 👨‍✈️ Danh sách Crew")
                    if crew_list:
                        li_items = "".join([f"<li>{c}</li>" for c in crew_list])
                        st.markdown(f'<div class="info-box"><ul>{li_items}</ul></div>', unsafe_allow_html=True)
                    else:
                        st.warning("⚠️ Không tìm thấy dữ liệu Crew.")

                with col_right:
                    st.markdown("### 💺 JumpSeaters")
                    if jump_seaters_list:
                        js_items = "".join([f"<li>{j}</li>" for j in jump_seaters_list])
                        st.markdown(f'<div class="info-box" style="border-left-color: #f4b400;"><ul>{js_items}</ul></div>', unsafe_allow_html=True)
                    else:
                        st.info("ℹ️ Chuyến bay này không có JumpSeaters.")

                # --- NÚT XUẤT FILE ---
                st.write("")
                if st.button("🚀 XUẤT FILE WORD NGAY", use_container_width=True):
                    try:
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
                            label=f"📥 Tải xuống file: GD_QH{flt_val}.docx",
                            data=bio,
                            file_name=f"GD_QH{flt_val}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                        st.success("Tạo file thành công!")
                    except Exception as e:
                        st.error(f"Lỗi khi render file Word: {e}")
            else:
                st.warning(f"❌ Không tìm thấy chuyến bay số: {search_flt}")
