
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

# 2. CSS Custom: Đảm bảo hiển thị tốt trên cả Light/Dark Mode Chrome
st.markdown("""
    <style>
    .main { background-color: transparent; }
    /* Ô thông số chính (Metric) */
    [data-testid="stMetric"] {
        background-color: #f0f2f6;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    [data-testid="stMetricLabel"] { color: #31333F !important; }
    [data-testid="stMetricValue"] { color: #1a73e8 !important; font-weight: bold; }
    
    /* Hộp danh sách Crew & JumpSeaters - Ép nền trắng chữ đen */
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

# Kiểm tra file template
if not os.path.exists(TEMPLATE_FILE):
    st.error(f"⚠️ Không tìm thấy file template tại: {TEMPLATE_FILE}")
else:
    # Sidebar cho thao tác cài đặt
    with st.sidebar:
        st.header("Cài đặt")
        uploaded_excel = st.file_uploader("1. Upload file Excel Data", type=["xlsx"])
        st.info("💡 File 'template.docx' đã được tích hợp sẵn trên hệ thống.")

    if uploaded_excel:
        # Đọc dữ liệu (Dòng tiêu đề nằm ở hàng 3 -> index 2)
        df = pd.read_excel(uploaded_excel, header=2)
        
        # Làm sạch tên cột (xóa khoảng trắng thừa)
        df.columns = [str(c).strip() for c in df.columns]

        # --- DÒ CỘT TỰ ĐỘNG (Dynamic Detection) ---
        js_col = None
        crew_col = None
        for col in df.columns:
            if "JumpSeaters" in col:
                js_col = col
            if "Crew" in col and "Crew #" not in col:
                crew_col = col

        # Ô tìm kiếm chuyến bay
        search_flt = st.text_input("🔍 Nhập số hiệu chuyến bay (ví dụ: 102, 208, 147...)", "")

        if search_flt:
            # Tìm dòng chứa chuyến bay (convert FLT sang string để tìm kiếm)
            target_row = df[df['FLT'].astype(str).str.contains(search_flt)]

            if not target_row.empty:
                start_idx = target_row.index[0]

                # --- XỬ LÝ DỮ LIỆU CƠ BẢN (FIX LỖI .0) ---
                raw_flt = df.loc[start_idx, 'FLT']
                if pd.api.types.is_number(raw_flt):
                    flt_val = str(int(raw_flt)) # Biến 208.0 thành 208
                else:
                    flt_val = str(raw_flt).replace('.0', '')

                reg_val = str(df.loc[start_idx, 'REG'])
                dep_val = str(df.loc[start_idx, 'DEP'])
                arr_val = str(df.loc[start_idx, 'ARR'])

                # Format DATE (DDMMMYY - ví dụ: 14MAR26)
                raw_date = df.loc[start_idx, 'DATE']
                try:
                    if isinstance(raw_date, datetime):
                        date_val = raw_date.strftime('%d%b%y').upper()
                    else:
                        date_obj = pd.to_datetime(raw_date, dayfirst=True)
                        date_val = date_obj.strftime('%d%b%y').upper()
                except:
                    date_val = str(raw_date).replace('.0', '')

                # --- QUÉT CREW VÀ JUMPSEATERS THEO KHỐI ---
                crew_list = []
                jump_seaters_list = []
                
                for i in range(start_idx, len(df)):
                    # Dừng lại nếu chạm tới chuyến bay tiếp theo (cột FLT có dữ liệu mới)
                    if i > start_idx and pd.notna(df.loc[i, 'FLT']):
                        break
                    
                    # Lấy Crew từ cột đã detect
                    if crew_col and pd.notna(df.loc[i, crew_col]):
                        c_val = str(df.loc[i, crew_col]).strip()
                        if c_val.lower() not in ["crew", "name"]:
                            crew_list.append(c_val)
                    
                    # Lấy JumpSeaters từ cột đã detect
                    if js_col and pd.notna(df.loc[i, js_col]):
                        j_val = str(df.loc[i, js_col]).strip()
                        if j_val.lower() not in ["jumpseaters", "name"]:
                            jump_seaters_list.append(j_val)

                # --- HIỂN THỊ GIAO DIỆN XEM TRƯỚC (PREVIEW) ---
                st.markdown("---")
                st.subheader(f"📊 Thông tin chuyến bay QH{flt_val}")

                # Hàng metric chính
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

                # --- NÚT XUẤT FILE WORD ---
                st.write("")
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
                        
                        # Lưu vào buffer để download
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
                        st.success("Tạo file thành công! Bấm nút phía trên để tải về.")
                    except Exception as e:
                        st.error(f"Lỗi khi render file Word: {e}")
            else:
                st.warning(f"❌ Không tìm thấy chuyến bay số: {search_flt}")
