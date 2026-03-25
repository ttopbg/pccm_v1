"""
app.py  –  Streamlit web app, KHÔNG cần API key
Deploy: push lên GitHub (app.py + teacher_core.py + requirements.txt) -> streamlit.io/cloud
"""
import io
import streamlit as st
from teacher_core import process_data

NIEN_KHOA_OPTIONS = ["2025-2026", "2026-2027", "2027-2028"]

st.set_page_config(
    page_title="PCCM",
    page_icon="🔞",
    layout="wide",
)

# ── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
:root {
  --step-bg:        rgba(46, 117, 182, 0.10);
  --step-border:    #2E75B6;
  --success-bg:     rgba(67, 160, 71, 0.12);
  --success-border: #43a047;
  --help-bg:        rgba(46, 117, 182, 0.07);
  --help-border:    rgba(46, 117, 182, 0.30);
  --code-bg:        rgba(0,0,0,0.06);
  --tag-green-bg:   rgba(67,160,71,0.15);
  --tag-green-fg:   #2e7d32;
  --tag-blue-bg:    rgba(46,117,182,0.15);
  --tag-blue-fg:    #1a5296;
  --tag-orange-bg:  rgba(230,119,0,0.15);
  --tag-orange-fg:  #a05000;
}
@media (prefers-color-scheme: dark) {
  :root {
    --step-bg:        rgba(46, 117, 182, 0.20);
    --step-border:    #5ba3d9;
    --success-bg:     rgba(67, 160, 71, 0.20);
    --success-border: #66bb6a;
    --help-bg:        rgba(46, 117, 182, 0.12);
    --help-border:    rgba(91,163,217,0.35);
    --code-bg:        rgba(255,255,255,0.08);
    --tag-green-bg:   rgba(67,160,71,0.22);
    --tag-green-fg:   #81c784;
    --tag-blue-bg:    rgba(46,117,182,0.22);
    --tag-blue-fg:    #90caf9;
    --tag-orange-bg:  rgba(230,119,0,0.22);
    --tag-orange-fg:  #ffb74d;
  }
}
[data-theme="dark"] {
  --step-bg:        rgba(46, 117, 182, 0.20);
  --step-border:    #5ba3d9;
  --success-bg:     rgba(67, 160, 71, 0.20);
  --success-border: #66bb6a;
  --help-bg:        rgba(46, 117, 182, 0.12);
  --help-border:    rgba(91,163,217,0.35);
  --code-bg:        rgba(255,255,255,0.08);
  --tag-green-bg:   rgba(67,160,71,0.22);
  --tag-green-fg:   #81c784;
  --tag-blue-bg:    rgba(46,117,182,0.22);
  --tag-blue-fg:    #90caf9;
  --tag-orange-bg:  rgba(230,119,0,0.22);
  --tag-orange-fg:  #ffb74d;
}

/* Header */
.main-header{background:linear-gradient(135deg,#1F4E79,#2E75B6);color:white;
  padding:1.4rem 2rem;border-radius:12px;margin-bottom:1.4rem;text-align:center}
.main-header h1{margin:0;font-size:1.7rem}
.main-header p{margin:.35rem 0 0;opacity:.85;font-size:.9rem}

/* Step boxes */
.step-box{background:var(--step-bg);border-left:4px solid var(--step-border);
  padding:.75rem 1rem;border-radius:0 8px 8px 0;margin-bottom:1rem}
.success-box{background:var(--success-bg);border-left:4px solid var(--success-border);
  padding:.75rem 1rem;border-radius:0 8px 8px 0}

/* Sidebar help styles */
.help-section{background:var(--help-bg);border:1px solid var(--help-border);
  border-radius:10px;padding:1rem 1.1rem;margin-bottom:.9rem;font-size:.88rem;line-height:1.6}
.help-section h4{margin:0 0 .5rem;font-size:.95rem;font-weight:700}
.help-section ul{margin:.3rem 0 0 1rem;padding:0}
.help-section li{margin:.25rem 0}
.help-section code{background:var(--code-bg);padding:.1rem .35rem;
  border-radius:4px;font-size:.82rem}
.help-section table{width:100%;border-collapse:collapse;font-size:.82rem;margin-top:.5rem}
.help-section th{text-align:left;padding:.3rem .5rem;opacity:.7;font-weight:600;
  border-bottom:1px solid var(--help-border)}
.help-section td{padding:.3rem .5rem;border-bottom:1px solid var(--help-border)}
.tag{display:inline-block;padding:.05rem .45rem;border-radius:20px;font-size:.78rem;font-weight:600}
.tag-green{background:var(--tag-green-bg);color:var(--tag-green-fg)}
.tag-blue {background:var(--tag-blue-bg); color:var(--tag-blue-fg)}
.tag-orange{background:var(--tag-orange-bg);color:var(--tag-orange-fg)}
.example-row{font-family:monospace;font-size:.8rem;word-break:break-all}
</style>
""", unsafe_allow_html=True)

# ── SIDEBAR – Hướng dẫn sử dụng ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📖 Hướng dẫn sử dụng")

    with st.expander("📁  1. Chuẩn bị file đầu vào", expanded=False):
        st.markdown("""
<div class="help-section">
<h4>Yêu cầu file Excel</h4>
<ul>
  <li>Định dạng: <code>.xlsx</code>, <code>.xls</code>, <code>.xlsm</code></li>
  <li>Phải có sheet tên <b>Data</b> (không phân biệt hoa/thường)</li>
  <li>Nếu không có sheet Data, sheet đầu tiên sẽ được dùng</li>
</ul>
<h4>Các cột bắt buộc trong sheet Data</h4>
<table>
  <tr><th>Cột</th><th>Tên chấp nhận được</th></tr>
  <tr><td><span class="tag tag-blue">Họ tên</span></td><td>Họ tên, Họ và tên, Giáo viên…</td></tr>
  <tr><td><span class="tag tag-blue">PCCM</span></td><td>PCCM, Phân công, Môn học giảng dạy, Giảng dạy lớp…</td></tr>
</table>
<h4>Các cột không bắt buộc</h4>
<ul>
  <li><b>STT</b> — Số thứ tự, TT…</li>
  <li><b>Ngày sinh</b> — Ngày sinh, DOB…</li>
</ul>
</div>
""", unsafe_allow_html=True)

    with st.expander("✍️  2. Chú ý về cột PCCM", expanded=False):
        st.markdown("""
<div class="help-section">
<h4>Cấu trúc cơ bản</h4>
<p>Dữ liệu theo dạng <b>Tên môn: danh sách lớp</b>, nhiều môn ngăn cách bằng <code>+</code></p>
<div class="example-row">Hóa: 10A1, 10A2 + Sử: 10D1, 10D2</div>

<h4>Các định dạng lớp được hỗ trợ</h4>
<table>
  <tr><th>Dạng viết</th><th>Kết quả</th></tr>
  <tr><td><code>10A1, 10A2, 10A3</code></td><td>Liệt kê thường</td></tr>
  <tr><td><code>10A123</code></td><td>→ 10A1, 10A2, 10A3</td></tr>
  <tr><td><code>10A1-10A5</code></td><td>→ 10A1 đến 10A5</td></tr>
  <tr><td><code>10A1 đến 10A5</code></td><td>→ 10A1 đến 10A5</td></tr>
  <tr><td><code>11A3,4</code></td><td>→ 11A3, 11A4</td></tr>
  <tr><td><code>11A1(52)</code></td><td>→ 11A1 (bỏ sĩ số)</td></tr>
  <tr><td><code>(11A1, 11A2)</code></td><td>→ 11A1, 11A2 (giữ lớp trong ngoặc)</td></tr>
  <tr><td><code>11C</code></td><td>Lớp không có số index</td></tr>
</table>

<h4>Dấu phân cách được hỗ trợ</h4>
<ul>
  <li>Giữa các lớp: <code>,</code> &nbsp;<code>;</code> &nbsp;khoảng trắng, <b>chưa phân biệt 10A12 là 10A1 và 10A2 hay chỉ là 10A12</b></li>
  <li>Giữa các môn: <code>+</code> &nbsp;hoặc tên môn đứng trước lớp trực tiếp</li>
  <li>Ví dụ: <span class="example-row">Hóa: 10A2, Sử 10D1</span> vẫn được nhận diện đúng</li>
</ul>
</div>
""", unsafe_allow_html=True)

    with st.expander("🔤  3. Nhận diện tên môn học", expanded=False):
        st.markdown("""
<div class="help-section">
<h4>Cách hoạt động</h4>
# <p>Hệ thống nhận diện tên môn theo <b>3 tầng</b>, không cần API:</p>
# <ul>
#   <li><span class="tag tag-green">Tầng 1</span> Khớp chính xác (có dấu): <code>Hóa học</code> → <code>HOAHOC</code></li>
#   <li><span class="tag tag-green">Tầng 2</span> Khớp không dấu + chuỗi con: <code>hoa hoc</code>, <code>hoa</code></li>
#   <li><span class="tag tag-orange">Tầng 3</span> Fuzzy match: nhận diện tên viết gần đúng, sai dấu</li>
# </ul>
<h4>Bảng mã môn học</h4>
<table>
  <tr><th>Tên môn</th><th>Mã</th></tr>
  <tr><td>Ngữ văn / Văn</td><td><code>NGUVAN</code></td></tr>
  <tr><td>Toán / Toán học</td><td><code>TOAN</code></td></tr>
  <tr><td>Tiếng Anh / Anh / NN1</td><td><code>ANH</code></td></tr>
  <tr><td>Lịch sử / Sử</td><td><code>LICHSU</code></td></tr>
  <tr><td>Địa lý / Địa</td><td><code>DIALY</code></td></tr>
  <tr><td>Vật lý / Lý</td><td><code>VATLY</code></td></tr>
  <tr><td>Hóa học / Hóa</td><td><code>HOAHOC</code></td></tr>
  <tr><td>Sinh học / Sinh</td><td><code>SINH</code></td></tr>
  <tr><td>Tin học / Tin</td><td><code>TINHOC</code></td></tr>
  <tr><td>GDTC / Thể dục</td><td><code>GDTC</code></td></tr>
  <tr><td>GDQP / Quốc phòng</td><td><code>GDQP</code></td></tr>
  <tr><td>KTPL / GDKTPL</td><td><code>GDKTPL</code></td></tr>
  <tr><td>GDĐP / GDDP</td><td><code>NDGDDP</code></td></tr>
  <tr><td>HĐTN / TNHN</td><td><code>TNHN</code></td></tr>
  <tr><td>Công nghệ (CN/NN)</td><td><code>CONGNGHE(CN/NN)</code></td></tr>
  <tr><td>KHTN</td><td><code>KHTN</code></td></tr>
  <tr><td>Lịch sử &amp; Địa lý</td><td><code>LICHSUDIALI</code></td></tr>
</table>
</div>
""", unsafe_allow_html=True)

    with st.expander("📊  4. Cấu trúc file đầu ra", expanded=False):
        st.markdown("""
<div class="help-section">
<h4>File output gồm 3 sheet</h4>

<p><span class="tag tag-blue">Sheet 1: Class</span></p>
<ul>
  <li>A1 = <b>Niên khóa</b>, B1 = năm học đã chọn</li>
  <li>A2 = <b>Lớp</b>, B2 = <b>Khối</b></li>
  <li>Từ A3: tổng hợp tất cả lớp, sắp xếp theo khối</li>
</ul>

<p><span class="tag tag-green">Sheet 2: Teachers</span></p>
<table>
  <tr><th>Cột</th><th>Nội dung</th></tr>
  <tr><td>STT</td><td>Số thứ tự</td></tr>
  <tr><td>Họ tên</td><td>Tên giáo viên</td></tr>
  <tr><td>Ngày sinh</td><td>dd/mm/yyyy</td></tr>
  <tr><td>SĐT</td><td>Để trống, nhập sau</td></tr>
  <tr><td>Môn dạy</td><td>Mã môn, cách nhau dấu phẩy</td></tr>
  <tr><td>TBM / CN</td><td>Để trống, nhập sau</td></tr>
  <tr><td>PCCM</td><td>Dạng: <code>10A1-TOAN,11B2-ANH</code></td></tr>
</table>

<p><span class="tag tag-orange">Sheet 3: Students</span></p>
<p>Dòng tiêu đề cố định: STT · Mã HS · Họ tên · Lớp · Giới tính · Ngày sinh · SĐT · Email · Tài khoản</p>
</div>
""", unsafe_allow_html=True)

    with st.expander("⚠️  5. Xử lý trùng lặp & lưu ý", expanded=False):
        st.markdown("""
<div class="help-section">
<h4>Xử lý tổ hợp môn-lớp trùng</h4>
<ul>
  <li><b>Trùng trong cùng 1 GV:</b> chỉ giữ 1 tổ hợp, bỏ trùng tự động</li>
  <li><b>Trùng giữa 2+ GV:</b> thêm tên GV vào cuối để phân biệt</li>
</ul>
<div class="example-row">
  12A2-HOAHOC(Nguyễn Tuấn Anh)<br>
  12A2-HOAHOC(Đoàn Văn Chiến)
</div>

<h4>Ngày sinh</h4>
<ul>
  <li>Chấp nhận: <code>dd/mm/yyyy</code>, <code>yyyy-mm-dd</code>, <code>dd-mm-yyyy</code></li>
  <li>Tự động nhận diện Excel serial date</li>
  <li>Output luôn định dạng <code>dd/mm/yyyy</code></li>
</ul>

<h4>Tên môn không nhận diện được</h4>
<ul>
  <li>Giữ nguyên dạng chữ HOA trong cột PCCM</li>
  <li>Đánh dấu <code>?</code> nếu hoàn toàn trống</li>
</ul>
</div>
""", unsafe_allow_html=True)

#     with st.expander("🚀  6. Hướng dẫn chạy & deploy", expanded=False):
#         st.markdown("""
# <div class="help-section">
# <h4>Chạy local (máy tính)</h4>
# <ol style="margin:.3rem 0 0 1rem;padding:0">
#   <li>Cài thư viện: <code>pip install openpyxl pandas</code></li>
#   <li>Chạy: <code>python convert_teachers_local.py</code></li>
# </ol>

# <h4>Deploy Streamlit Cloud (web)</h4>
# <ol style="margin:.3rem 0 0 1rem;padding:0">
#   <li>Push 3 file lên GitHub:<br>
#       <code>app.py</code> &nbsp;<code>teacher_core.py</code> &nbsp;<code>requirements.txt</code></li>
#   <li>Vào <a href="https://streamlit.io/cloud" target="_blank">streamlit.io/cloud</a>
#       → <b>New app</b></li>
#   <li>Chọn repo, branch <code>main</code>, main file: <code>app.py</code></li>
#   <li>Nhấn <b>Deploy</b> — không cần cấu hình thêm gì</li>
# </ol>

# <h4>Không cần API key</h4>
# <p>Toàn bộ nhận diện môn học chạy offline bằng từ điển + fuzzy matching.</p>
# </div>
# """, unsafe_allow_html=True)

# ── MAIN CONTENT ─────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>🙃 Tạo file Import PCCM 🙃</h1>
  <p>File Input cần có sheet <b>Data</b></p>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="step-box"><b>1️⃣</b> Tải lên file Excel chứa sheet <code>Data</code></div>',
            unsafe_allow_html=True)
uploaded = st.file_uploader("Chọn file Excel", type=["xlsx","xls","xlsm"],
                             label_visibility="collapsed")

st.markdown('<div class="step-box"><b>2️⃣</b> Chọn niên khóa</div>',
            unsafe_allow_html=True)
nien_khoa = st.selectbox("Niên khóa", options=NIEN_KHOA_OPTIONS, label_visibility="collapsed")

st.markdown('<div class="step-box"><b>3️⃣</b> Nhấn nút để xử lý</div>',
            unsafe_allow_html=True)
run_btn = st.button("\u25b6  Chuyển đổi", type="primary", use_container_width=True,
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
        st.markdown('<div class="success-box">\u2705 <b>Chuyển đổi thành công!</b></div>',
                    unsafe_allow_html=True)
        st.download_button("\u2b07\ufe0f  Tải xuống file Excel", data=result_bytes,
                           file_name=out_name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
    except Exception as e:
        prog_bar.empty()
        st.error(f"\u274c Lỗi: {e}")
