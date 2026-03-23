"""
app.py  –  Streamlit web app, KHÔNG cần API key
Deploy: push lên GitHub (app.py + teacher_core.py + requirements.txt) -> streamlit.io/cloud
"""
import io
import streamlit as st
from teacher_core import process_data

NIEN_KHOA_OPTIONS = ["2025-2026", "2026-2027", "2027-2028"]

st.set_page_config(page_title="PCCM", page_icon="🔞", layout="centered")

st.markdown("""
<style>
/* Màu thích nghi sáng/tối qua CSS custom properties */
:root {
  --step-bg:       rgba(46, 117, 182, 0.10);
  --step-border:   #2E75B6;
  --success-bg:    rgba(67, 160, 71, 0.12);
  --success-border:#43a047;
}
@media (prefers-color-scheme: dark) {
  :root {
    --step-bg:       rgba(46, 117, 182, 0.20);
    --step-border:   #5ba3d9;
    --success-bg:    rgba(67, 160, 71, 0.20);
    --success-border:#66bb6a;
  }
}
[data-theme="dark"] {
  --step-bg:       rgba(46, 117, 182, 0.20);
  --step-border:   #5ba3d9;
  --success-bg:    rgba(67, 160, 71, 0.20);
  --success-border:#66bb6a;
}
.main-header{background:linear-gradient(135deg,#1F4E79,#2E75B6);color:white;
  padding:1.5rem 2rem;border-radius:12px;margin-bottom:1.5rem;text-align:center}
.main-header h1{margin:0;font-size:1.8rem}
.main-header p{margin:.4rem 0 0;opacity:.85;font-size:.95rem}
.step-box{background:var(--step-bg);border-left:4px solid var(--step-border);
  padding:.8rem 1rem;border-radius:0 8px 8px 0;margin-bottom:1rem}
.success-box{background:var(--success-bg);border-left:4px solid var(--success-border);
  padding:.8rem 1rem;border-radius:0 8px 8px 0}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
  <h1>🙃 Tạo file Import PCCM 🙃</h1>
  <p>File Input cần có sheet <b>Data</b></p>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="step-box"><b>1️⃣:</b> Tải lên file Excel chứa sheet <code>Data</code></div>',
            unsafe_allow_html=True)
uploaded = st.file_uploader("Chọn file Excel", type=["xlsx","xls","xlsm"],
                             label_visibility="collapsed")

st.markdown('<div class="step-box";"background-color: #007BFF"><b>2️⃣:</b> Chọn niên khóa</div>',
            unsafe_allow_html=True)
nien_khoa = st.selectbox("Niên khóa", options=NIEN_KHOA_OPTIONS, label_visibility="collapsed")

st.markdown('<div class="step-box"><b>3️⃣:</b> Nhấn nút để xử lý</div>',
            unsafe_allow_html=True)
run_btn = st.button("▶  Chuyển đổi", type="primary", use_container_width=True,
                    disabled=(uploaded is None))

if uploaded is None:
    st.info("✌️Vui lòng tải lên file Excel đầu vào✌️")

if run_btn and uploaded:
    log_area  = st.empty()
    prog_bar  = st.progress(0)
    log_lines = []

    import pandas as pd
    from teacher_core import detect_header_row, find_column
    raw_bytes = uploaded.read()
    try:
        _xl  = pd.ExcelFile(io.BytesIO(raw_bytes))
        _sn  = next((s for s in _xl.sheet_names if s.strip().lower()=="data"), _xl.sheet_names[0])
        _rdf = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=_sn, header=None)
        _hri = detect_header_row(_rdf)
        _df  = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=_sn, header=_hri)
        _df.columns = [str(c).strip() for c in _df.columns]
        _ch  = find_column(_df, ["họ tên","họ và tên","tên","giáo viên","ho ten"])
        total_t = len(_df[_df[_ch].notna()]) if _ch else 1
    except Exception:
        total_t = 1

    processed = [0]
    def progress_cb(msg):
        log_lines.append(msg)
        log_area.code("\n".join(log_lines[-20:]), language=None)
        if "Xử lý giáo viên" in msg:
            processed[0] += 1
            prog_bar.progress(min(int(processed[0]/total_t*90), 90))

    try:
        result_bytes = process_data(io.BytesIO(raw_bytes), nien_khoa, progress_cb=progress_cb)
        prog_bar.progress(100)
        fname = uploaded.name.rsplit(".",1)[0]
        out_name = f"{fname}_output_{nien_khoa}.xlsx"
        st.markdown('<div class="success-box">✅ <b>Chuyển đổi thành công!</b></div>',
                    unsafe_allow_html=True)
        st.download_button("⬇️  Tải xuống file Excel", data=result_bytes,
                           file_name=out_name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
    except Exception as e:
        prog_bar.empty()
        st.error(f"❌ Lỗi: {e}")
