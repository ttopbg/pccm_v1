# teacher_core.py  –  logic dùng chung, KHÔNG cần Anthropic API
import re, io, unicodedata, difflib
from datetime import datetime, timedelta, date as date_type
from collections import defaultdict
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SUBJECT_MAP = {
    "ngữ văn":"NGUVAN","ngữ văn học":"NGUVAN","van":"NGUVAN","nguvan":"NGUVAN","nv":"NGUVAN",
    "toán":"TOAN","toán học":"TOAN","toan":"TOAN",
    "tiếng anh":"ANH","ngoại ngữ 1":"ANH","ngoại ngữ 2":"ANH","ngoại ngữ":"ANH",
    "anh":"ANH","nn1":"ANH","nn2":"ANH","english":"ANH",
    "lịch sử":"LICHSU","lich su":"LICHSU","sử":"LICHSU","su":"LICHSU","lichsu":"LICHSU",
    "giáo dục thể chất":"GDTC","giao duc the chat":"GDTC","thể dục":"GDTC",
    "the duc":"GDTC","gdtc":"GDTC","td":"GDTC",
    "giáo dục quốc phòng và an ninh":"GDQP","giáo dục quốc phòng":"GDQP",
    "giao duc quoc phong":"GDQP","quốc phòng":"GDQP","quoc phong":"GDQP",
    "qpan":"GDQP","gdqp":"GDQP",
    "địa lí":"DIALY","địa lý":"DIALY","dia li":"DIALY","dia ly":"DIALY",
    "địa":"DIALY","dia":"DIALY","dialy":"DIALY",
    "giáo dục kinh tế và pháp luật":"GDKTPL","kinh tế pháp luật":"GDKTPL",
    "kinh te phap luat":"GDKTPL","gdktpl":"GDKTPL","ktpl":"GDKTPL",
    "vật lí":"VATLY","vật lý":"VATLY","vat li":"VATLY","vat ly":"VATLY",
    "lí":"VATLY","lý":"VATLY","li":"VATLY","ly":"VATLY","vatly":"VATLY","vl":"VATLY",
    "hóa học":"HOAHOC","hoá học":"HOAHOC","hoa hoc":"HOAHOC",
    "hóa":"HOAHOC","hoá":"HOAHOC","hoa":"HOAHOC","hoahoc":"HOAHOC","hh":"HOAHOC",
    "sinh học":"SINH","sinh hoc":"SINH","sinh":"SINH",
    "cnnn":"CONGNGHE(NN)","nông nghiệp":"CONGNGHE(NN)","nong nghiep":"CONGNGHE(NN)",
    "công nghệ (nn)":"CONGNGHE(NN)","công nghệ(nn)":"CONGNGHE(NN)","cong nghe nn":"CONGNGHE(NN)",
    "cncn":"CONGNGHE(CN)","công nghiệp":"CONGNGHE(CN)","cong nghiep":"CONGNGHE(CN)",
    "công nghệ (cn)":"CONGNGHE(CN)","công nghệ(cn)":"CONGNGHE(CN)","cong nghe cn":"CONGNGHE(CN)",
    "công nghệ":"CONGNGHE","cong nghe":"CONGNGHE",
    "tin học":"TINHOC","tin hoc":"TINHOC","tin":"TINHOC","tinhoc":"TINHOC",
    "nội dung giáo dục địa phương":"NDGDDP","giáo dục địa phương":"NDGDDP",
    "giao duc dia phuong":"NDGDDP","gdđp":"NDGDDP","gddp":"NDGDDP",
    "gd dp":"NDGDDP","gd11dp":"NDGDDP","nd gd dp":"NDGDDP",
    "hoạt động trải nghiệm, hướng nghiệp":"TNHN","hoạt động trải nghiệm":"TNHN",
    "hoat dong trai nghiem":"TNHN","hướng nghiệp":"TNHN","huong nghiep":"TNHN",
    "hđ trải nghiệm":"TNHN","hđtn":"TNHN","hdtn":"TNHN","hđtn hn":"TNHN","tnhn":"TNHN",
    "tiếng pháp":"TIENGPHAP","tieng phap":"TIENGPHAP","pháp":"TIENGPHAP",
    "tiếng nga":"TIENGNGA","tieng nga":"TIENGNGA",
    "tiếng nhật":"TIENGNHAT","tieng nhat":"TIENGNHAT",
    "tiếng trung":"TIENGTRUNG","tieng trung":"TIENGTRUNG",
    "tiếng hàn":"TIENGHAN","tieng han":"TIENGHAN",
    "nghề phổ thông":"NGHEPHOTHONG","nghe pho thong":"NGHEPHOTHONG","nghề":"NGHEPHOTHONG",
    "âm nhạc":"AMNHAC","am nhac":"AMNHAC","nhạc":"AMNHAC","nhac":"AMNHAC",
    "mỹ thuật":"MYTHUAT","mĩ thuật":"MYTHUAT","my thuat":"MYTHUAT",
    "mi thuat":"MYTHUAT","mt":"MYTHUAT",
    "lịch sử và địa lí":"LICHSUDIALI","lịch sử và địa lý":"LICHSUDIALI",
    "lich su va dia ly":"LICHSUDIALI","ls&đl":"LICHSUDIALI",
    "ls & đl":"LICHSUDIALI","lsdl":"LICHSUDIALI",
    "khoa học tự nhiên":"KHTN","khoa hoc tu nhien":"KHTN","khtn":"KHTN",
    "giáo dục công dân":"GDCD","giao duc cong dan":"GDCD","gdcd":"GDCD",
    "hoạt động ngoài giờ lên lớp":"HDNGLL",
    "hoat dong ngoai gio len lop":"HDNGLL","hđngll":"HDNGLL","hdngll":"HDNGLL",
    "tiếng dân tộc thiểu số":"TDTTS","tieng dan toc thieu so":"TDTTS",
    "nghệ thuật":"NGHETHUAT","nghe thuat":"NGHETHUAT",
}

def _remove_accent(s):
    s = unicodedata.normalize("NFD", s)
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower().strip()

_ALL_CODES = list(set(SUBJECT_MAP.values()))
_MAP_NO_ACCENT = {_remove_accent(k): v for k, v in SUBJECT_MAP.items()}
for _c in _ALL_CODES:
    _MAP_NO_ACCENT[_c.lower()] = _c

_CLASS_PAT = r'\d{2}[A-Za-z]+\.?\d{0,2}(?:\.\d+)?'
_SUBJECT_STOPWORDS = {
    "đến","den","và","va","từ","tu","lớp","lop",
    "khối","khoi","tới","toi","to","the","from"
}
_fuzzy_cache: dict = {}

def match_subject_local(raw):
    if not raw: return None
    s  = raw.lower().strip()
    sn = _remove_accent(raw)
    if s in SUBJECT_MAP: return SUBJECT_MAP[s]
    if sn in _MAP_NO_ACCENT: return _MAP_NO_ACCENT[sn]
    if s.upper() in _ALL_CODES: return s.upper()
    best, best_len = None, 0
    for key, code in SUBJECT_MAP.items():
        if key in s or s in key:
            if len(key) > best_len: best, best_len = code, len(key)
    if best: return best
    for key_nd, code in _MAP_NO_ACCENT.items():
        if key_nd in sn or sn in key_nd:
            if len(key_nd) > best_len: best, best_len = code, len(key_nd)
    if best: return best
    if sn in _fuzzy_cache: return _fuzzy_cache[sn]
    matches = difflib.get_close_matches(sn, list(_MAP_NO_ACCENT.keys()), n=1, cutoff=0.72)
    result = _MAP_NO_ACCENT[matches[0]] if matches else None
    _fuzzy_cache[sn] = result
    return result

def get_subject_code(raw, _=None):
    return match_subject_local(raw.strip()) if raw and raw.strip() else None

def _enumerate_splits(grade, alpha, digits, known_classes):
    """
    Liệt kê TẤT CẢ các cách tách hợp lệ của chuỗi `digits` sau prefix `grade+alpha`,
    trong đó mỗi phần đều thuộc known_classes.
    Trả về list of list[str], mỗi phần tử là một cách tách.
    Chỉ trả về các cách mà TOÀN BỘ mảnh đều nằm trong known_classes.
    """
    n = len(digits)
    results = []

    def backtrack(pos, current):
        if pos == n:
            results.append(list(current))
            return
        for length in range(1, n - pos + 1):
            candidate = f"{grade}{alpha}{digits[pos:pos+length]}"
            if candidate in known_classes:
                current.append(candidate)
                backtrack(pos + length, current)
                current.pop()

    backtrack(0, [])
    return results


def _is_ambiguous(grade, alpha, digits, known_classes):
    """
    Kiểm tra xem chuỗi có nhiều hơn 1 cách tách hợp lệ không.
    """
    splits = _enumerate_splits(grade, alpha, digits, known_classes)
    return splits if len(splits) > 1 else None


def expand_class_range(text, known_classes=None, resolved_ambiguities=None):
    """
    Parse chuỗi lớp học thành danh sách lớp.
    known_classes: tập hợp tên lớp hợp lệ (từ cột GVCN).
    resolved_ambiguities: dict {raw_token: [lớp đã chọn]} — lựa chọn của người dùng
                          cho các chuỗi ambiguous. Nếu None, dùng greedy.
    """
    if resolved_ambiguities is None:
        resolved_ambiguities = {}

    text = re.sub(r'(\d{2}[A-Za-z]+\d*)\(\d+\)', r'\1', text)
    classes = []
    rp = re.compile(r'(\d{2})([A-Za-zÀ-ỹ]+)(\d+)\s*(?:đến|den|-)\s*\1\2(\d+)', re.UNICODE)
    for m in rp.finditer(text):
        g,a,s,e = m.groups()
        for i in range(int(s), int(e)+1): classes.append(f"{g}{a}{i}")
    text = rp.sub('', text)

    def _split_digits(grade, alpha, digits):
        raw_token = f"{grade}{alpha}{digits}"
        # Nếu đã có lựa chọn của user cho token này → dùng luôn
        if raw_token in resolved_ambiguities:
            return resolved_ambiguities[raw_token]
        if known_classes:
            # Kiểm tra có ambiguous không
            splits = _enumerate_splits(grade, alpha, digits, known_classes)
            if len(splits) == 1:
                return splits[0]  # chỉ 1 cách → dùng luôn
            elif len(splits) > 1:
                # Ambiguous nhưng chưa có resolved → dùng greedy (sẽ được hỏi ở UI)
                # Greedy: khớp dài nhất từ trái sang
                result = []
                i = 0
                while i < len(digits):
                    matched = False
                    for length in range(len(digits) - i, 0, -1):
                        candidate = f"{grade}{alpha}{digits[i:i+length]}"
                        if candidate in known_classes:
                            result.append(candidate)
                            i += length
                            matched = True
                            break
                    if not matched:
                        result.append(f"{grade}{alpha}{digits[i]}")
                        i += 1
                return result
            else:
                # Không khớp known nào → mỗi ký tự 1 lớp
                return [f"{grade}{alpha}{x}" for x in digits]
        else:
            return [f"{grade}{alpha}{x}" for x in digits]

    def _compact(m):
        g,a,d = m.group(1),m.group(2),m.group(3)
        for c in _split_digits(g, a, d):
            if c not in classes: classes.append(c)
        return ''
    text = re.sub(r'(\d{2})([A-Za-z]+)(\d{3,})', _compact, text)
    text = re.sub(r'(\d{2})([A-Za-z]+)(\d{2})(?![,;.\s])', _compact, text)

    def _suffix(m):
        base,nums = m.group(1),m.group(2)
        for n in re.split(r'[,\s]+',nums):
            if n: classes.append(f"{base}{n.strip()}")
        return ''
    text = re.sub(r'(\d{2}[A-Za-z]+)(\d(?:,\s*\d)+)(?!\d)', _suffix, text)
    classes.extend(re.findall(_CLASS_PAT, text))
    result,seen = [],set()
    for c in classes:
        c=c.strip().strip(',').strip()
        if c and c not in seen: seen.add(c); result.append(c)
    return result


def detect_ambiguous_in_data(df, col_pccm, col_gvcn, known_classes):
    """
    Quét toàn bộ dữ liệu PCCM, tìm tất cả chuỗi token ambiguous.
    Trả về list of dict:
      {
        "token":       "10A123",          # chuỗi gốc trong file
        "grade":       "10",
        "alpha":       "A",
        "digits":      "123",
        "splits":      [["10A1","10A2","10A3"], ["10A12","10A3"], ...],
        "occurrences": ["GV Nguyễn Văn A (Văn: 10A123)", ...]   # mô tả ngữ cảnh
      }
    Mỗi token duy nhất chỉ xuất hiện 1 lần trong kết quả.
    """
    if not known_classes:
        return []

    _ambig_pat = re.compile(r'(\d{2})([A-Za-z]+)(\d{2,})', re.UNICODE)
    found = {}   # token → dict

    col_hoten = None
    for cand in ["họ tên","họ và tên","tên","giáo viên","ho ten","hoten"]:
        col_hoten = find_column(df, [cand])
        if col_hoten: break

    for _, row in df.iterrows():
        praw = str(row.get(col_pccm, "")).strip() if pd.notna(row.get(col_pccm)) else ""
        if not praw:
            continue
        hoten = str(row.get(col_hoten, "")).strip() if col_hoten else ""

        # Chuẩn hoá sơ bộ giống parse_pccm
        text = re.sub(r'\([^)]*\)', '', praw)
        text = text.replace(';', ',').replace('\n', '+')
        # Xoá range đã rõ (10A1-10A5, 10A1 đến 10A5)
        text = re.sub(r'\d{2}[A-Za-z]+\d+\s*(?:đến|den|-)\s*\d{2}[A-Za-z]+\d+', '', text)

        for m in _ambig_pat.finditer(text):
            grade, alpha, digits = m.group(1), m.group(2), m.group(3)
            # Chỉ xét nếu chuỗi ≥ 2 chữ số (10A12 trở lên)
            if len(digits) < 2:
                continue
            token = f"{grade}{alpha}{digits}"
            # Nếu token CHÍNH LÀ 1 lớp hợp lệ trong known → không ambiguous
            if token in known_classes:
                continue
            splits = _enumerate_splits(grade, alpha, digits, known_classes)
            if len(splits) <= 1:
                continue  # không ambiguous

            # Tìm ngữ cảnh môn học xung quanh token
            context_start = max(0, m.start() - 20)
            context = "..." + text[context_start:m.end()+5].strip() + "..."
            occurrence = f"{hoten}: …{context}…" if hoten else context

            if token not in found:
                found[token] = {
                    "token": token,
                    "grade": grade,
                    "alpha": alpha,
                    "digits": digits,
                    "splits": splits,
                    "occurrences": [occurrence],
                }
            else:
                if occurrence not in found[token]["occurrences"]:
                    found[token]["occurrences"].append(occurrence)

    return list(found.values())

def parse_pccm(raw_pccm, known_classes=None, resolved_ambiguities=None):
    if not raw_pccm or (isinstance(raw_pccm,float) and pd.isna(raw_pccm)): return []
    text = str(raw_pccm).strip()
    def ep(m):
        inner=m.group(1).strip()
        return '' if re.fullmatch(r'\d+',inner) else ','+inner+','
    text = re.sub(r'\(([^)]*)\)', ep, text)
    text = text.replace(';',',').replace('\n','+')
    CRP = (r'\d{2}[A-Za-z]+\d+\s*(?:đến|den|-)\s*\d{2}[A-Za-z]+\d+'
           r'|\d{2}[A-Za-z]+\d{3,}' r'|'+_CLASS_PAT)
    tokens,results = [],[]
    tr = re.compile(r'(?P<class>'+CRP+r')|(?P<sep>[+,\s]+)|(?P<colon>:)'
                    r'|(?P<word>[A-Za-zÀ-ỹĐđ][A-Za-zÀ-ỹĐđ\(\)]*)|(?P<other>.)',re.UNICODE)
    for m in tr.finditer(text): tokens.append((m.lastgroup, m.group().strip()))
    merged,i = [],0
    while i < len(tokens):
        kind,val = tokens[i]
        if kind=='word':
            words=[val]; j=i+1
            while j<len(tokens):
                k2,v2=tokens[j]
                if k2=='word': words.append(v2); j+=1
                elif k2=='sep' and j+1<len(tokens) and tokens[j+1][0]=='word':
                    words.append(tokens[j+1][1]); j+=2
                else: break
            merged.append(('word',' '.join(words))); i=j
        elif kind=='sep':
            if val: merged.append(('sep',val))
            i+=1
        elif kind in ('class','colon','other'): merged.append((kind,val)); i+=1
        else: i+=1
    cur_subj,cur_cls = None,[]
    def flush(s,c,o):
        if s and c: o.append((s,c))
        elif c: o.append(("",c))
    idx=0
    while idx<len(merged):
        kind,val = merged[idx]
        if kind=='word':
            if _remove_accent(val) in _SUBJECT_STOPWORDS: idx+=1; continue
            nns=None
            for k2,v2 in merged[idx+1:]:
                if k2!='sep': nns=(k2,v2); break
            if nns and nns[0]=='colon':
                flush(cur_subj,cur_cls,results); cur_subj=val; cur_cls=[]
                idx+=1
                while idx<len(merged) and merged[idx][0] in ('sep','colon'): idx+=1
            elif nns and nns[0]=='class':
                flush(cur_subj,cur_cls,results); cur_subj=val; cur_cls=[]
                idx+=1
                while idx<len(merged) and merged[idx][0]=='sep': idx+=1
            else: idx+=1
        elif kind=='class':
            cur_cls.extend(expand_class_range(val, known_classes, resolved_ambiguities))
            idx+=1
        elif kind in ('sep','colon','other'): idx+=1
        else: idx+=1
    flush(cur_subj,cur_cls,results)
    return results

def format_date(val):
    try:
        if val is None: return None,""
        if isinstance(val,datetime): return val,val.strftime("%d/%m/%Y")
        if isinstance(val,date_type):
            dt=datetime(val.year,val.month,val.day); return dt,dt.strftime("%d/%m/%Y")
        if isinstance(val,(int,float)):
            if pd.isna(val): return None,""
            dt=datetime(1899,12,30)+timedelta(days=int(val)); return dt,dt.strftime("%d/%m/%Y")
        s=str(val).strip()
        if not s or s.lower() in ("nan","nat","none",""): return None,""
        for fmt in ("%d/%m/%Y","%Y-%m-%d","%d-%m-%Y","%m/%d/%Y","%d/%m/%y","%Y/%m/%d"):
            try: dt=datetime.strptime(s,fmt); return dt,dt.strftime("%d/%m/%Y")
            except: pass
        return None,s
    except: return None,""

def find_column(df,candidates):
    cl={c.lower().strip():c for c in df.columns}
    for cand in candidates:
        c=cand.lower().strip()
        if c in cl: return cl[c]
        for key,orig in cl.items():
            if c in key or key in c: return orig
    return None

def detect_header_row(sdf):
    kws=['stt','họ tên','họ và tên','giáo viên','pccm','phân công','ngày sinh']
    for i,row in sdf.iterrows():
        vals=[str(v).lower().strip() for v in row.values if pd.notna(v)]
        if sum(1 for v in vals for k in kws if k in v)>=2: return i
    return 0

def get_grade(cls):
    m=re.match(r'^(\d{2})',str(cls).strip())
    return int(m.group(1)) if m else None

def _sh(ws,row,ncols,color="1F4E79"):
    fill=PatternFill("solid",fgColor=color)
    font=Font(bold=True,color="FFFFFF",name="Arial",size=11)
    align=Alignment(horizontal="center",vertical="center",wrap_text=True)
    for col in range(1,ncols+1):
        cell=ws.cell(row=row,column=col)
        cell.fill=fill; cell.font=font; cell.alignment=align

def _sdr(ws,row,ncols,even,left_cols=()):
    fill=PatternFill("solid",fgColor="EBF3FB" if even else "FFFFFF")
    font=Font(name="Arial",size=10)
    for col in range(1,ncols+1):
        cell=ws.cell(row=row,column=col)
        cell.fill=fill; cell.font=font
        cell.alignment=(Alignment(horizontal="left",vertical="center",wrap_text=True)
                        if col in left_cols
                        else Alignment(horizontal="center",vertical="center"))

def _ab(ws,sr,er,ncols):
    thin=Side(style='thin',color='B0C4DE')
    border=Border(left=thin,right=thin,top=thin,bottom=thin)
    for row in range(sr,er+1):
        for col in range(1,ncols+1): ws.cell(row=row,column=col).border=border

def process_data(input_src, nien_khoa: str, progress_cb=None,
                 resolved_ambiguities=None) -> bytes:
    """
    Xử lý file Excel đầu vào. KHÔNG cần API key.
    input_src: bytes / BytesIO / str path
    """
    def log(m):
        if progress_cb: progress_cb(m)

    src = (io.BytesIO(input_src) if isinstance(input_src,(bytes,bytearray))
           else input_src)

    xl = pd.ExcelFile(src)
    ds = next((s for s in xl.sheet_names if s.strip().lower()=="data"), xl.sheet_names[0])
    log(f"Đọc sheet '{ds}'...")
    rdf = pd.read_excel(src,sheet_name=ds,header=None)
    hri = detect_header_row(rdf)
    df  = pd.read_excel(src,sheet_name=ds,header=hri)
    df.columns = [str(c).strip() for c in df.columns]

    col_stt   = find_column(df,["stt","tt","số thứ tự","no"])
    col_hoten = find_column(df,["họ tên","họ và tên","tên","giáo viên","ho ten","hoten"])
    col_ngay  = find_column(df,["ngày sinh","ngay sinh","sinh ngày","dob","birthday"])
    col_pccm  = find_column(df,["pccm","phân công chuyên môn","phân công",
                                 "giảng dạy lớp","môn học giảng dạy","phan cong","giang day"])
    col_gvcn  = find_column(df,["gvcn","chủ nhiệm","chu nhiem","chủ nhiệm lớp",
                                 "chu nhiem lop","lớp chủ nhiệm","lop chu nhiem","cn"])
    if not col_hoten: raise ValueError("Không tìm thấy cột Họ tên!")
    if not col_pccm:  raise ValueError("Không tìm thấy cột PCCM!")

    df = df[df[col_hoten].notna()&(df[col_hoten].astype(str).str.strip()!="")].copy()
    df = df.reset_index(drop=True)

    # ── Bước 1: Thu thập known_classes từ toàn bộ cột GVCN ───────────────────
    # Đọc raw bằng regex cơ bản — KHÔNG qua expand_class_range để tránh tách sai
    # Mỗi ô GVCN thường chứa tên lớp rõ ràng: "10A1", "10A12", "10A1, 10A2"
    known_classes: set = set()
    if col_gvcn:
        log("Đọc danh sách lớp từ cột GVCN...")
        _raw_cls_pat = re.compile(r'\d{2}[A-Za-z]+\d+', re.UNICODE)
        for val in df[col_gvcn]:
            if pd.notna(val) and str(val).strip():
                for c in _raw_cls_pat.findall(str(val)):
                    known_classes.add(c.strip())
        log(f"  → Nhận diện được {len(known_classes)} lớp: {', '.join(sorted(known_classes))}")

    total = len(df)
    teachers = []

    for idx,row in df.iterrows():
        log(f"Xử lý giáo viên {idx+1}/{total}: {row[col_hoten]}")
        stt   = str(row[col_stt]).strip() if col_stt and pd.notna(row.get(col_stt)) else str(idx+1)
        hoten = str(row[col_hoten]).strip()
        ndt,nstr = (format_date(row[col_ngay]) if col_ngay and pd.notna(row.get(col_ngay))
                    else (None,""))
        praw = str(row[col_pccm]).strip() if pd.notna(row.get(col_pccm)) else ""

        # Đọc lớp chủ nhiệm
        gvcn_str = ""
        if col_gvcn and pd.notna(row.get(col_gvcn)):
            gvcn_raw = str(row[col_gvcn]).strip()
            if gvcn_raw:
                # Parse lớp CN — dùng known_classes để tách chính xác
                cn_classes = expand_class_range(gvcn_raw, known_classes if known_classes else None)
                gvcn_str = ", ".join(cn_classes) if cn_classes else gvcn_raw

        parsed = parse_pccm(praw, known_classes if known_classes else None,
                             resolved_ambiguities or {})
        scodes,mllist = [],[]
        for sr,ll in parsed:
            code = get_subject_code(sr)
            if code:
                if code not in scodes: scodes.append(code)
                for lop in ll:
                    lop=lop.strip()
                    if lop: mllist.append((lop,code))
            else:
                for lop in ll:
                    lop=lop.strip()
                    if lop: mllist.append((lop,sr.upper() if sr else "?"))

        seen=set(); uml=[]
        for lop,code in mllist:
            if (lop,code) not in seen: seen.add((lop,code)); uml.append((lop,code))

        teachers.append({"stt":stt,"ho_ten":hoten,"ngay_dt":ndt,"ngay_str":nstr,
                         "subject_codes":scodes,"mon_lop_list":uml,"gvcn_str":gvcn_str})

    pc = defaultdict(list)
    for t in teachers:
        for lop,code in t["mon_lop_list"]: pc[(lop,code)].append(t["ho_ten"])

    for t in teachers:
        parts=[]
        for lop,code in t["mon_lop_list"]:
            key=(lop,code)
            parts.append(f"{lop}-{code}({t['ho_ten']})" if len(pc[key])>1 else f"{lop}-{code}")
        t["pccm_str"]=",".join(parts)

    all_cls = sorted(set(lop.strip() for t in teachers for lop,_ in t["mon_lop_list"]),
                     key=lambda c:(get_grade(c) or 99,c))

    log("Tạo file Excel đầu ra...")
    wb = openpyxl.Workbook()

    # ── Class sheet ───────────────────────────────────────────────────
    wc=wb.active; wc.title="Class"
    wc["A1"]="Niên khóa"; wc["B1"]=nien_khoa
    wc["A2"]="Lớp";       wc["B2"]="Khối"
    for r in (1,2):
        for col in ("A","B"):
            c=wc[f"{col}{r}"]
            c.fill=PatternFill("solid",fgColor="1F4E79")
            c.font=Font(bold=True,color="FFFFFF",name="Arial",size=11)
            c.alignment=Alignment(horizontal="center",vertical="center")
    for i,cls in enumerate(all_cls):
        r=i+3
        wc.cell(row=r,column=1,value=cls); wc.cell(row=r,column=2,value=get_grade(cls))
        for col in (1,2):
            c=wc.cell(row=r,column=col)
            c.fill=PatternFill("solid",fgColor="EBF3FB" if i%2==0 else "FFFFFF")
            c.font=Font(name="Arial",size=10)
            c.alignment=Alignment(horizontal="center",vertical="center")
    _ab(wc,1,len(all_cls)+2,2)
    wc.column_dimensions["A"].width=14; wc.column_dimensions["B"].width=10
    wc.freeze_panes="A3"

    # ── Teachers sheet ────────────────────────────────────────────────
    wt=wb.create_sheet("Teachers")
    ht=["STT","Họ tên","Ngày sinh","SĐT","Môn dạy","TBM","CN","PCCM"]
    for ci,h in enumerate(ht,1): wt.cell(row=1,column=ci,value=h)
    _sh(wt,1,len(ht)); wt.row_dimensions[1].height=30
    for i,t in enumerate(teachers):
        rn=i+2
        wt.cell(row=rn,column=1,value=t["stt"])
        wt.cell(row=rn,column=2,value=t["ho_ten"])
        dc=wt.cell(row=rn,column=3)
        if t["ngay_dt"]: dc.value=t["ngay_dt"]; dc.number_format="DD/MM/YYYY"
        else: dc.value=t["ngay_str"]
        wt.cell(row=rn,column=4,value="")
        wt.cell(row=rn,column=5,value=", ".join(t["subject_codes"]))
        wt.cell(row=rn,column=6,value="")
        wt.cell(row=rn,column=7,value=t.get("gvcn_str",""))
        wt.cell(row=rn,column=8,value=t["pccm_str"])
        _sdr(wt,rn,len(ht),i%2==0,left_cols=(2,5,8))
    _ab(wt,1,len(teachers)+1,len(ht))
    for ci,w in enumerate([6,25,14,14,30,10,10,80],1):
        wt.column_dimensions[get_column_letter(ci)].width=w
    wt.freeze_panes="A2"

    # ── Students sheet ────────────────────────────────────────────────
    ws=wb.create_sheet("Students")
    hs=["STT","Mã HS","Họ tên","Lớp","Giới tính","Ngày sinh","Số điện thoại","Email","Tài khoản"]
    for ci,h in enumerate(hs,1): ws.cell(row=1,column=ci,value=h)
    _sh(ws,1,len(hs)); ws.row_dimensions[1].height=30
    _ab(ws,1,1,len(hs))
    for ci,w in enumerate([6,14,25,10,12,14,16,28,18],1):
        ws.column_dimensions[get_column_letter(ci)].width=w
    ws.freeze_panes="A2"

    out=io.BytesIO(); wb.save(out); out.seek(0)
    log("Hoàn thành!")
    return out.read()
