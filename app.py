import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import os
from datetime import datetime

# Cấu hình trang rộng hơn để xem cho sướng
st.set_page_config(page_title="Bamboo Gendec System", layout="wide")

# --- XỬ LÝ ĐƯỜNG DẪN FILE ---
base_dir = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_FILE = os.path.join(base_dir, "template.docx")

# CSS tùy chỉnh để làm giao diện nhìn "xịn" hơn
# CSS tùy chỉnh để fix lỗi Dark Mode
st.markdown("""
    <style>
    /* Tổng thể */
    .main { background-color: transparent; }

    /* Tùy chỉnh ô Metric (Thông tin DATE, REG, DEP, ARR) */
    [data-testid="stMetric"] {
        background-color: #f0f2f6; /* Màu xám nhạt */
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }

    /* Ép màu chữ cho Metric Label (tiêu đề nhỏ) và Value (chữ to) */
    [data-testid="stMetricLabel"] {
        color: #31333F !important; /* Màu đen xám */
    }
    [data-testid="stMetricValue"] {
        color: #1a73e8 !important; /* Màu xanh thương hiệu */
        font-weight: bold;
    }

    /* Tùy chỉnh hộp danh sách Crew và JumpSeaters */
    .crew-box {
        background-color: #ffffff; /* Nền trắng */
        color: #1a1c21 !important; /* Chữ đen đậm */
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #1a73e8;
        line-height: 1.6;
    }

    /* Đảm bảo chữ trong danh sách li cũng màu đen */
    .crew-box ul li {
        color: #1a1c21 !important;
        list-style-type: none;
        border-bottom: 1px solid #eee;
        padding: 5px 0;
    }

    /* Chỉnh lại màu tiêu đề Section cho rõ hơn trong Dark mode */
    h3 {
        color: #4A90E2 !important;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("✈️ Bamboo Airways Gendec Generator")

if not os.path.exists(TEMPLATE_FILE):
    st.error(f"⚠️ Không tìm thấy file template tại: {TEMPLATE_FILE}")
else:
    # Sidebar cho việc upload
    with st.sidebar:
        st.header("Cài đặt")
        uploaded_excel = st.file_uploader("1. Upload Excel Data", type=["xlsx"])
        st.info("File template đã được tích hợp sẵn hệ thống.")

    if uploaded_excel:
        df = pd.read_excel(uploaded_excel, header=2)
        df.columns = [str(c).strip() for c in df.columns]

        # Ô nhập số hiệu chuyến bay nổi bật ở giữa
        search_flt = st.text_input("🔍 Nhập số hiệu chuyến bay để tìm kiếm (ví dụ: 101, 228...)", "")

        if search_flt:
            target_row = df[df['FLT'].astype(str).str.contains(search_flt)]

            if not target_row.empty:
                start_idx = target_row.index[0]

                # --- TRÍCH XUẤT DỮ LIỆU ---
                flt_val = str(df.loc[start_idx, 'FLT'])
                reg_val = str(df.loc[start_idx, 'REG'])
                dep_val = str(df.loc[start_idx, 'DEP'])
                arr_val = str(df.loc[start_idx, 'ARR'])

                # Format DATE (18AUG25)
                raw_date = df.loc[start_idx, 'DATE']
                try:
                    if isinstance(raw_date, datetime):
                        date_val = raw_date.strftime('%d%b%y').upper()
                    else:
                        date_obj = pd.to_datetime(raw_date, dayfirst=True)
                        date_val = date_obj.strftime('%d%b%y').upper()
                except:
                    date_val = str(raw_date)

                # Quét Crew và JumpSeaters
                crew_list = []
                jump_seaters_list = []
                for i in range(start_idx, len(df)):
                    if i > start_idx and pd.notna(df.loc[i, 'FLT']):
                        break
                    if pd.notna(df.loc[i, 'Crew']):
                        crew_list.append(str(df.loc[i, 'Crew']))
                    p_val = df.iloc[i, 15]
                    if pd.notna(p_val):
                        jump_seaters_list.append(str(p_val))

                # --- GIAO DIỆN TRỰC QUAN ---
                st.markdown("---")
                st.subheader(f"📊 Thông tin xem trước: Chuyến bay QH{flt_val}")

                # Hàng 1: Các thông số chính
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("DATE", date_val)
                col2.metric("REG", reg_val)
                col3.metric("DEPARTURE", dep_val)
                col4.metric("ARRIVAL", arr_val)

                # Hàng 2: Chi tiết tổ bay và JumpSeaters
                st.write("")
                col_left, col_right = st.columns(2)

                with col_left:
                    st.markdown("### 👨‍✈️ Danh sách tổ bay (Crew)")
                    crew_html = "".join([f"<li>{c}</li>" for c in crew_list])
                    st.markdown(f"""<div class="crew-box"><ul>{crew_html}</ul></div>""", unsafe_allow_html=True)

                with col_right:
                    st.markdown("### 💺 JumpSeaters")
                    if jump_seaters_list:
                        js_html = "".join([f"<li>{j}</li>" for j in jump_seaters_list])
                        st.markdown(
                            f"""<div class="crew-box" style="border-left-color: #f4b400;"><ul>{js_html}</ul></div>""",
                            unsafe_allow_html=True)
                    else:
                        st.info("Không có JumpSeaters cho chuyến bay này.")

                # Nút bấm xuất file to và nổi bật
                st.write("")
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
                        label="📥 Bấm vào đây để tải file: GD_QH" + flt_val + ".docx",
                        data=bio,
                        file_name=f"GD_QH{flt_val}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
            else:
                st.warning(f"❌ Không tìm thấy dữ liệu cho chuyến bay số: {search_flt}")