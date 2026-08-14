import streamlit as st
import pandas as pd
import os
import sys

# Import custom data processor & excel exporter
sys.path.append(os.path.dirname(__file__))
from data_processor import process_month_folder, analyze_discrepancy_causes, STANDARD_PAIRS, build_human_remark, format_currency
from excel_exporter import export_reconciliation_excel


# Streamlit Page Configuration
st.set_page_config(
    page_title="Đối Soát Tài Khoản 136 - 336",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Clean, Simple & Premium Aesthetics
st.markdown("""
<style>
    .main {
        background-color: #F8F9FA;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    .header-box {
        background: linear-gradient(135deg, #1E3C72 0%, #2A5298 100%);
        padding: 20px 28px;
        border-radius: 10px;
        color: white;
        margin-bottom: 20px;
    }
    .header-title {
        font-size: 24px;
        font-weight: 700;
        margin: 0;
    }
    .header-subtitle {
        font-size: 13.5px;
        opacity: 0.9;
        margin-top: 4px;
    }
    
    /* Simple KPI Cards */
    .kpi-card {
        background-color: white;
        padding: 16px 20px;
        border-radius: 8px;
        border-left: 4px solid #2A5298;
        box-shadow: 0 2px 6px rgba(0,0,0,0.04);
        margin-bottom: 16px;
    }
    .kpi-card-danger {
        border-left: 4px solid #DC3545;
    }
    .kpi-title {
        font-size: 12.5px;
        font-weight: 600;
        color: #6C757D;
        text-transform: uppercase;
    }
    .kpi-value {
        font-size: 22px;
        font-weight: 700;
        color: #212529;
        margin-top: 4px;
    }
    
    /* Simple Remark Box */
    .remark-box {
        background-color: #EBF3FE;
        border-left: 5px solid #1E3C72;
        padding: 14px 18px;
        border-radius: 6px;
        font-size: 14.5px;
        color: #1E3C72;
        font-weight: 500;
        margin-bottom: 20px;
        line-height: 1.5;
    }
</style>
""", unsafe_allow_html=True)

# Base Paths - Dynamic resolution for local and Streamlit Cloud environments
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if os.path.basename(CURRENT_DIR) == "dữ liệu":
    BASE_DIR = os.path.abspath(os.path.join(CURRENT_DIR, ".."))
else:
    BASE_DIR = CURRENT_DIR

if not os.path.exists(os.path.join(BASE_DIR, "đầu vào")):
    if os.path.exists(os.path.join(os.getcwd(), "đầu vào")):
        BASE_DIR = os.getcwd()

INPUT_DIR = os.path.join(BASE_DIR, "đầu vào")
DATA_DIR = os.path.join(BASE_DIR, "dữ liệu")
OUTPUT_DIR = os.path.join(BASE_DIR, "đầu ra")
os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


# Header Section
st.markdown("""
<div class="header-box">
    <div class="header-title">⚡ CÔNG CỤ ĐỐI SOÁT CHÊNH LỆCH TÀI KHOẢN 136 - 336</div>
    <div class="header-subtitle">Tổng công ty Điện lực TP. HCM (080100) ⇄ Công ty Điện lực đối ứng</div>
</div>
""", unsafe_allow_html=True)

# Sidebar - Month Selection & Data Management
st.sidebar.header("⚙️ Chọn Kỳ Số Liệu")

month_folders = []
if os.path.exists(INPUT_DIR):
    month_folders = [d for d in os.listdir(INPUT_DIR) if os.path.isdir(os.path.join(INPUT_DIR, d))]
month_folders.sort()

selected_folder = st.sidebar.selectbox(
    "📁 Chọn Thư Mục Kỳ Tháng:",
    options=month_folders if month_folders else ["tháng 7"],
    index=0
)

target_month_path = os.path.join(INPUT_DIR, selected_folder)

if st.sidebar.button("🔄 Làm Mới & Đọc Lại Dữ Liệu", type="primary", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

st.sidebar.markdown("---")
# Sidebar - File Upload Section for Overwriting or Adding New Month
with st.sidebar.expander("📤 Tải Lên / Cập Nhật 4 File Số Liệu", expanded=False):
    st.caption("Cho phép chép đè file mới hoặc tạo kỳ tháng mới trực tiếp từ giao diện web.")
    upload_month_name = st.text_input("Tên kỳ tháng:", value=selected_folder)
    
    file_hcm_136 = st.file_uploader("1️⃣ TK 136 - HCM.xlsx", type=["xlsx"], key="u_hcm_136")
    file_dl_136 = st.file_uploader("2️⃣ TK 136 - Điện lực.xlsx", type=["xlsx"], key="u_dl_136")
    file_hcm_336 = st.file_uploader("3️⃣ TK 336 - HCM.xlsx", type=["xlsx"], key="u_hcm_336")
    file_dl_336 = st.file_uploader("4️⃣ TK 336 - Điện lực.xlsx", type=["xlsx"], key="u_dl_336")
    
    if st.button("🚀 Ghi Lưu & Cập Nhật Chương Trình", type="primary", use_container_width=True):
        if not upload_month_name.strip():
            st.error("Vui lòng nhập tên kỳ tháng!")
        else:
            save_dir = os.path.join(INPUT_DIR, upload_month_name.strip())
            os.makedirs(save_dir, exist_ok=True)
            saved_count = 0
            
            if file_hcm_136:
                with open(os.path.join(save_dir, "TK 136 - HCM.xlsx"), "wb") as f:
                    f.write(file_hcm_136.getbuffer())
                saved_count += 1
                
            if file_dl_136:
                with open(os.path.join(save_dir, "TK 136 - Điện lực.xlsx"), "wb") as f:
                    f.write(file_dl_136.getbuffer())
                saved_count += 1
                
            if file_hcm_336:
                with open(os.path.join(save_dir, "TK 336 - HCM.xlsx"), "wb") as f:
                    f.write(file_hcm_336.getbuffer())
                saved_count += 1
                
            if file_dl_336:
                with open(os.path.join(save_dir, "TK 336 - Điện lực.xlsx"), "wb") as f:
                    f.write(file_dl_336.getbuffer())
                saved_count += 1
                
            if saved_count > 0:
                st.cache_data.clear()
                st.success(f"✅ Đã lưu {saved_count} file vào kỳ [{upload_month_name}] thành công!")
                st.rerun()
            else:
                st.warning("Bạn chưa chọn file nào để tải lên.")

@st.cache_data(show_spinner=False)
def load_month_data(month_path):
    return process_month_folder(month_path)


# Dialog Modal for Detail View
if hasattr(st, "dialog"):
    @st.dialog("📋 Chi Tiết Bút Toán Chênh Lệch", width="large")
    def render_detail_dialog(pair_key, pair_data, remark_text):
        h_code = pair_data['h_code']
        p_code = pair_data['p_code']
        diff_val = pair_data['diff']
        h_missing = pair_data['hcm_missing']
        p_missing = pair_data['pcvt_missing']
        
        st.markdown(f"### 🎯 Cặp Tài Khoản: **{h_code} (HCM) ⇄ {p_code} (Điện lực)**")
        st.markdown(f"**Chênh lệch cuối kỳ:** <span style='color: #DC3545; font-weight: 700; font-size: 18px;'>{format_currency(diff_val)} VNĐ</span>", unsafe_allow_html=True)
        
        st.markdown(f"""
        <div class="remark-box">
            {remark_text}
        </div>
        """, unsafe_allow_html=True)
        
        col_un_h, col_un_p = st.columns(2)
        
        with col_un_h:
            st.markdown(f"#### 🏛️ Bút Toán HCM Bị Thiếu ({len(h_missing)} chứng từ)")
            if h_missing:
                df_un_h = pd.DataFrame(h_missing)[['date', 'gl_doc', 'net', 'desc', 'creator']].copy()
                df_un_h['Số Tiền (VNĐ)'] = df_un_h['net'].apply(format_currency)
                df_un_h = df_un_h[['date', 'gl_doc', 'Số Tiền (VNĐ)', 'desc', 'creator']]
                df_un_h.columns = ['Ngày CT', 'Số CT GL', 'Số Tiền (VNĐ)', 'Diễn Giải', 'Người Lập']
                st.dataframe(df_un_h, use_container_width=True, hide_index=True)
            else:
                st.success("✅ Phía HCM đã hạch toán đầy đủ.")
                
        with col_un_p:
            st.markdown(f"#### 🏢 Bút Toán Điện Lực Bị Thiếu ({len(p_missing)} chứng từ)")
            if p_missing:
                df_un_p = pd.DataFrame(p_missing)[['date', 'gl_doc', 'net', 'desc', 'creator']].copy()
                df_un_p['Số Tiền (VNĐ)'] = df_un_p['net'].apply(format_currency)
                df_un_p = df_un_p[['date', 'gl_doc', 'Số Tiền (VNĐ)', 'desc', 'creator']]
                df_un_p.columns = ['Ngày CT', 'Số CT GL', 'Số Tiền (VNĐ)', 'Diễn Giải', 'Người Lập']
                st.dataframe(df_un_p, use_container_width=True, hide_index=True)
            else:
                st.success("✅ Phía Điện lực đã hạch toán đầy đủ.")
                
        st.markdown("---")
        if st.button("❌ Đóng Màn Hình Chi Tiết", type="primary", use_container_width=True):
            st.rerun()

if os.path.exists(target_month_path):
    with st.spinner("Đang xử lý số liệu đối soát..."):
        df_th, hcm_accs, pcvt_accs, all_txs = load_month_data(target_month_path)
        cross_matches, unmatched_by_pair = analyze_discrepancy_causes(df_th, hcm_accs, pcvt_accs)
        
    tot_no = df_th['Nợ'].sum()
    tot_co = df_th['Có'].sum()
    tot_diff = df_th['CHÊNH LỆCH'].sum()
    num_mismatches = len(df_th[df_th['CHÊNH LỆCH'].abs() > 0.01])

    # Top KPI summary row
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">Tổng Dư Nợ (HCM)</div>
            <div class="kpi-value">{format_currency(tot_no)}</div>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">Tổng Dư Có (Điện lực)</div>
            <div class="kpi-value">{format_currency(tot_co)}</div>
        </div>
        """, unsafe_allow_html=True)
    with c3:
        st.markdown(f"""
        <div class="kpi-card kpi-card-danger">
            <div class="kpi-title">Tổng Chênh Lệch</div>
            <div class="kpi-value">{format_currency(tot_diff)}</div>
        </div>
        """, unsafe_allow_html=True)
    with c4:
        st.markdown(f"""
        <div class="kpi-card kpi-card-danger">
            <div class="kpi-title">Số Cặp TK Chênh Lệch</div>
            <div class="kpi-value">{num_mismatches} / {len(df_th)}</div>
        </div>
        """, unsafe_allow_html=True)

    # Main Navigation
    tab1, tab2, tab3 = st.tabs([
        "📋 1. Báo Cáo Chênh Lệch (Click Xem Chi Tiết)",
        "🕵️ 2. Phân Tích & Kiểm Dò Bút Toán",
        "📥 3. Xuất Báo Cáo Excel"
    ])

    # -------------------------------------------------------------
    # TAB 1: BÁO CÁO CHÊNH LỆCH & CLICK XEM CHI TIẾT
    # -------------------------------------------------------------
    with tab1:
        st.subheader("📋 Báo Cáo Chênh Lệch 52 Cặp Tài Khoản")
        st.caption("💡 **Hướng dẫn:** Anh bấm trực tiếp vào số tiền **Chênh Lệch** hoặc nút **🔍 Chi Tiết** để mở màn hình popup xem đúng các chứng từ chênh lệch đó. Xem xong anh chỉ cần bấm **Đóng** để quay lại.")

        filter_option = st.radio(
            "Hiển thị danh sách:",
            options=["🔴 Chỉ các cặp có Chênh Lệch", "🌐 Tất cả 52 cặp Tài khoản"],
            horizontal=True
        )

        # Prepare report rows
        report_data = []
        for idx, row in df_th.iterrows():
            h_acc = row['HCM']
            p_acc = row['PCVT']
            pair_key = f"{h_acc} - {p_acc}"
            diff = row['CHÊNH LỆCH']
            
            p_info = unmatched_by_pair.get(pair_key, {
                'h_code': h_acc, 'p_code': p_acc, 'diff': diff,
                'hcm_missing': [], 'pcvt_missing': []
            })
            
            h_missing_cnt = len(p_info['hcm_missing'])
            p_missing_cnt = len(p_info['pcvt_missing'])
            
            report_data.append({
                'stt': idx + 1,
                'pair_key': pair_key,
                'h_acc': h_acc,
                'p_acc': p_acc,
                'no_val': row['Nợ'],
                'co_val': row['Có'],
                'diff': diff,
                'h_cnt': h_missing_cnt,
                'p_cnt': p_missing_cnt,
                'pair_data': p_info
            })

        if "Chỉ các cặp có Chênh Lệch" in filter_option:
            filtered_rows = [r for r in report_data if abs(r['diff']) > 0.01]
        else:
            filtered_rows = report_data

        # Render Header
        col_stt, col_pair, col_diff, col_h, col_p, col_action = st.columns([1, 2.5, 3, 2.2, 2.2, 1.8])
        with col_stt: st.markdown("**STT**")
        with col_pair: st.markdown("**Cặp Tài Khoản**")
        with col_diff: st.markdown("**Chênh Lệch (VNĐ)**")
        with col_h: st.markdown("**Bút Toán HCM**")
        with col_p: st.markdown("**Bút Toán Điện Lực**")
        with col_action: st.markdown("**Thao Tác**")
        st.markdown("---")

        for r in filtered_rows:
            c_stt, c_pair, c_diff, c_h, c_p, c_action = st.columns([1, 2.5, 3, 2.2, 2.2, 1.8])
            
            diff_fmt = format_currency(r['diff'])
            is_diff = abs(r['diff']) > 0.01
            
            with c_stt:
                st.write(f"**{r['stt']}**")
                
            with c_pair:
                st.write(f"🏢 **{r['h_acc']}** ⇄ **{r['p_acc']}**")
                
            with c_diff:
                if is_diff:
                    if st.button(f"🔴 {diff_fmt} VNĐ", key=f"btn_diff_{r['stt']}", type="secondary", help="Click để xem chi tiết chứng từ chênh lệch"):
                        remark_text = build_human_remark(
                            r['h_acc'], r['p_acc'], r['diff'], 
                            r['pair_data']['hcm_missing'], r['pair_data']['pcvt_missing'], 
                            cross_matches
                        )
                        if hasattr(st, "dialog"):
                            render_detail_dialog(r['pair_key'], r['pair_data'], remark_text)
                        else:
                            st.session_state['active_detail_pair'] = r['pair_key']
                else:
                    st.write(f"✅ {diff_fmt}")
                    
            with c_h:
                if r['h_cnt'] > 0:
                    st.markdown(f"⚠️ **{r['h_cnt']} chứng từ** (Thiếu)")
                else:
                    st.markdown("✅ Khớp 100%")
                    
            with c_p:
                if r['p_cnt'] > 0:
                    st.markdown(f"⚠️ **{r['p_cnt']} chứng từ** (Thiếu)")
                else:
                    st.markdown("✅ Khớp 100%")
                    
            with c_action:
                if is_diff:
                    if st.button("🔍 Chi Tiết", key=f"btn_act_{r['stt']}", type="primary"):
                        remark_text = build_human_remark(
                            r['h_acc'], r['p_acc'], r['diff'], 
                            r['pair_data']['hcm_missing'], r['pair_data']['pcvt_missing'], 
                            cross_matches
                        )
                        if hasattr(st, "dialog"):
                            render_detail_dialog(r['pair_key'], r['pair_data'], remark_text)
                        else:
                            st.session_state['active_detail_pair'] = r['pair_key']
                else:
                    st.caption("—")

        # Fallback inline detail display if st.dialog is not used or session_state active
        if 'active_detail_pair' in st.session_state and not hasattr(st, "dialog"):
            active_key = st.session_state['active_detail_pair']
            p_data = unmatched_by_pair.get(active_key)
            if p_data:
                st.markdown("---")
                st.subheader(f"📋 Chi Tiết Chứng Từ Chênh Lệch: {active_key}")
                remark_text = build_human_remark(
                    p_data['h_code'], p_data['p_code'], p_data['diff'],
                    p_data['hcm_missing'], p_data['pcvt_missing'], cross_matches
                )
                st.markdown(f'<div class="remark-box">{remark_text}</div>', unsafe_allow_html=True)
                
                col_un_h, col_un_p = st.columns(2)
                with col_un_h:
                    st.markdown(f"#### 🏛️ Bút Toán HCM Bị Thiếu ({len(p_data['hcm_missing'])} CT)")
                    if p_data['hcm_missing']:
                        df_un_h = pd.DataFrame(p_data['hcm_missing'])[['date', 'gl_doc', 'net', 'desc', 'creator']].copy()
                        df_un_h['Số Tiền (VNĐ)'] = df_un_h['net'].apply(format_currency)
                        df_un_h = df_un_h[['date', 'gl_doc', 'Số Tiền (VNĐ)', 'desc', 'creator']]
                        df_un_h.columns = ['Ngày CT', 'Số CT GL', 'Số Tiền (VNĐ)', 'Diễn Giải', 'Người Lập']
                        st.dataframe(df_un_h, use_container_width=True, hide_index=True)
                with col_un_p:
                    st.markdown(f"#### 🏢 Bút Toán Điện Lực Bị Thiếu ({len(p_data['pcvt_missing'])} CT)")
                    if p_data['pcvt_missing']:
                        df_un_p = pd.DataFrame(p_data['pcvt_missing'])[['date', 'gl_doc', 'net', 'desc', 'creator']].copy()
                        df_un_p['Số Tiền (VNĐ)'] = df_un_p['net'].apply(format_currency)
                        df_un_p = df_un_p[['date', 'gl_doc', 'Số Tiền (VNĐ)', 'desc', 'creator']]
                        df_un_p.columns = ['Ngày CT', 'Số CT GL', 'Số Tiền (VNĐ)', 'Diễn Giải', 'Người Lập']
                        st.dataframe(df_un_p, use_container_width=True, hide_index=True)
                
                if st.button("❌ Đóng Chi Tiết", key="btn_close_inline"):
                    del st.session_state['active_detail_pair']
                    st.rerun()

    # -------------------------------------------------------------
    # TAB 2: KIỂM DÒ BÚT TOÁN GÂY CHÊNH LỆCH
    # -------------------------------------------------------------
    with tab2:
        st.subheader("🕵️ Phân Tích & Kiểm Dò Bút Toán Chi Tiết Theo Cặp Tài Khoản")
        st.info("💡 **Giao diện kiểm dò chọn lọc:** Chọn cặp tài khoản bên dưới để xem diễn giải phân tích nguyên nhân và danh sách chứng từ khuyết thiếu.")
        
        diff_pair_keys = list(unmatched_by_pair.keys())
        
        selected_diff_pair = st.selectbox(
            "🎯 Chọn Cặp Tài Khoản Có Chênh Lệch Để Xem Phân Tích:",
            options=diff_pair_keys,
            index=0 if diff_pair_keys else None
        )
        
        if selected_diff_pair:
            pair_data = unmatched_by_pair[selected_diff_pair]
            h_code = pair_data['h_code']
            p_code = pair_data['p_code']
            diff_val = pair_data['diff']
            h_missing = pair_data['hcm_missing']
            p_missing = pair_data['pcvt_missing']
            
            remark_text = build_human_remark(
                h_code, p_code, diff_val, h_missing, p_missing, cross_matches
            )
            
            st.markdown(f"""
            <div class="remark-box">
                {remark_text}
            </div>
            """, unsafe_allow_html=True)
            
            col_un_h, col_un_p = st.columns(2)
            
            with col_un_h:
                st.markdown(f"#### 🏛️ Bút Toán Phía HCM Chưa Hạch Toán (HCM Bị Thiếu) — {len(h_missing)} chứng từ")
                if h_missing:
                    df_un_h = pd.DataFrame(h_missing)[['date', 'gl_doc', 'net', 'desc', 'creator']].copy()
                    df_un_h['Số Tiền (VNĐ)'] = df_un_h['net'].apply(format_currency)
                    df_un_h = df_un_h[['date', 'gl_doc', 'Số Tiền (VNĐ)', 'desc', 'creator']]
                    df_un_h.columns = ['Ngày CT', 'Số CT GL', 'Số Tiền (VNĐ)', 'Diễn Giải', 'Người Lập']
                    st.dataframe(df_un_h, use_container_width=True, hide_index=True)
                else:
                    st.success("✅ Phía HCM đã hạch toán đầy đủ, không bị thiếu chứng từ nào.")
                    
            with col_un_p:
                st.markdown(f"#### 🏢 Bút Toán Phía Điện Lực Chưa Hạch Toán (Điện Lực Bị Thiếu) — {len(p_missing)} chứng từ")
                if p_missing:
                    df_un_p = pd.DataFrame(p_missing)[['date', 'gl_doc', 'net', 'desc', 'creator']].copy()
                    df_un_p['Số Tiền (VNĐ)'] = df_un_p['net'].apply(format_currency)
                    df_un_p = df_un_p[['date', 'gl_doc', 'Số Tiền (VNĐ)', 'desc', 'creator']]
                    df_un_p.columns = ['Ngày CT', 'Số CT GL', 'Số Tiền (VNĐ)', 'Diễn Giải', 'Người Lập']
                    st.dataframe(df_un_p, use_container_width=True, hide_index=True)
                else:
                    st.success("✅ Phía Điện lực đã hạch toán đầy đủ, không bị thiếu chứng từ nào.")

    # -------------------------------------------------------------
    # TAB 3: EXPORT EXCEL
    # -------------------------------------------------------------
    with tab3:
        st.subheader("📥 Xuất Báo Cáo Excel Kết Quả Đối Soát")
        st.write("Xuất file Excel chứa Sheet TONG_HOP và Sheet PHAN_TICH_CHENH_LECH giải trình chứng từ lệch.")
        
        output_filename = f"Bao_Cao_Doi_Soat_136_336_{selected_folder.replace(' ', '_')}.xlsx"
        full_out_path = os.path.join(OUTPUT_DIR, output_filename)
        
        if st.button("🚀 Bắt Đầu Xuất File Excel", type="primary"):
            with st.spinner("Đang xuất file Excel..."):
                export_reconciliation_excel(df_th, full_out_path, period_text=selected_folder, cross_matches=cross_matches, unmatched_by_pair=unmatched_by_pair)
                
            st.success(f"✅ Đã xuất báo cáo thành công!")
            st.code(full_out_path, language="text")
            
            if os.path.exists(full_out_path):
                with open(full_out_path, "rb") as f:
                    st.download_button(
                        label="💾 Tải File Báo Cáo Excel Về Máy",
                        data=f.read(),
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
