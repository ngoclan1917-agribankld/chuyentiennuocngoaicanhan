# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import io, csv, re, unicodedata
import requests
from datetime import date, datetime
from itertools import count
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

try:
    import pycountry
except Exception:
    pycountry = None

# =========================
# ⚙️ CẤU HÌNH & TIÊU ĐỀ
# =========================
st.set_page_config(page_title="TẠO LỆNH CHUYỂN TIỀN QUỐC TẾ", page_icon="💸", layout="wide")
st.markdown(
    """
    <h1 style="text-align:center;color:#8B0000;">
        <span style="padding:6px 12px;border:2px solid #8B0000;border-radius:10px;">
            TẠO LỆNH CHUYỂN TIỀN QUỐC TẾ
        </span>
    </h1>
    """,
    unsafe_allow_html=True
)

# =========================
# 🧩 TIỆN ÍCH CHUNG
# =========================
HTML_TAG_RE = re.compile(r"<[^>]+>")
NBSP = "\u00A0"

def parse_vn_number(s: str) -> float:
    if s is None: return 0.0
    s = str(s).strip().replace(NBSP, " ")
    s = HTML_TAG_RE.sub(" ", s)
    if s == "": return 0.0
    s = s.replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s and "." not in s:
        s = s.replace(",", ".")
    try: return float(s)
    except Exception: return 0.0

def fmt_vn_int(n) -> str:
    try: return f"{int(round(float(n), 0)):,}".replace(",", ".")
    except Exception: return "0"

def fmt_usd(n) -> str:
    try: return f"{float(n):,.2f}"
    except Exception: return "0.00"

def fmt_ddmmyyyy(d) -> str:
    if isinstance(d, (date, datetime)): return d.strftime("%d/%m/%Y")
    return ""

def clean_ccy(val) -> str:
    if val is None: return ""
    s = str(val).strip().replace(NBSP, " ")
    s = HTML_TAG_RE.sub(" ", s).upper()
    return s if re.fullmatch(r"[A-Z]{3}", s) else ""

def normalize_name(s: str) -> list:
    if s is None: return []
    s = str(s).replace(NBSP, " ")
    s = HTML_TAG_RE.sub(" ", s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    toks = [t for t in s.split() if t]
    stop = {"co","ltd","company","the","and","account","acc","fees","fee","university",
            "bank","beneficiary","name","accountname","transfer","payment","inv"}
    return [t for t in toks if t not in stop]

def names_loose_match(a: str, b: str) -> bool:
    A, B = set(normalize_name(a)), set(normalize_name(b))
    if not A or not B: return False
    if A.issubset(B) or B.issubset(A): return True
    inter = len(A & B)
    jacc = inter / max(1, len(A | B))
    return jacc >= 0.7

def to_usd(amount: float, vnd_per_ccy: float, vnd_per_usd: float) -> float:
    if amount is None or pd.isna(amount): return 0.0
    if not (vnd_per_ccy and vnd_per_usd) or vnd_per_ccy <= 0 or vnd_per_usd <= 0: return 0.0
    return float(amount) * float(vnd_per_ccy) / float(vnd_per_usd)

def id_type_value(selected: str, other_text: str) -> str:
    if "Khác" in (selected or "") and (other_text or "").strip(): return other_text.strip()
    if "(Để trống)" in (selected or ""): return ""
    return selected or ""

def get_iso2_country_codes():
    items = []
    if pycountry:
        try:
            for c in pycountry.countries:
                items.append((c.alpha_2.upper(), f"{c.alpha_2.upper()} – {c.name}"))
        except Exception:
            pass
    if not items:
        fallback = {"VN":"Viet Nam","US":"United States","AU":"Australia","JP":"Japan",
                    "KR":"Korea, Republic of","SG":"Singapore","CN":"China","DE":"Germany",
                    "FR":"France","GB":"United Kingdom","TH":"Thailand","CA":"Canada"}
        items = [(k, f"{k} – {v}") for k, v in fallback.items()]
    items.sort(key=lambda x: x[0]); return items

def get_iso4217_codes():
    codes = set()
    if pycountry:
        try:
            for c in pycountry.currencies:
                if getattr(c, "alpha_3", None): codes.add(c.alpha_3.upper())
        except Exception: pass
    if not codes:
        codes = {"USD","EUR","JPY","GBP","AUD","CAD","CHF","CNY","HKD","SGD","KRW",
                 "THB","TWD","MYR","IDR","INR","VND","NZD","SEK","NOK","DKK","RUB",
                 "AED","SAR","QAR","KWD","BHD","TRY","BRL","MXN","ZAR","PLN","HUF"}
    return sorted(list(codes))

def fetch_gdp_per_capita_usd(iso2: str, year: int):
    if not iso2 or not year: return None, None
    for y in [year, year-1, year-2]:
        try:
            u=f"https://api.worldbank.org/v2/country/{iso2.lower()}/indicator/NY.GDP.PCAP.CD?date={y}:{y}&format=json"
            js=requests.get(u,timeout=12).json()
            if isinstance(js,list) and len(js)>1 and js[1] and js[1][0]["value"] is not None:
                return float(js[1][0]["value"]), y
        except Exception: pass
    return None, None

# =========================
# 📥 ĐỌC FILE LỊCH SỬ & CHỌN BẢNG TỐT NHẤT
# =========================
def _read_all_tables(uploaded_file):
    if uploaded_file is None: return []
    bio = io.BytesIO(uploaded_file.read())
    out = []

    # Excel
    try:
        bio.seek(0)
        df = pd.read_excel(bio, engine="openpyxl")
        if isinstance(df, pd.DataFrame) and not df.empty:
            out.append(df)
    except Exception:
        pass

    # CSV
    try:
        bio.seek(0)
        raw = bio.read().decode(errors="ignore")
        try:
            dialect = csv.Sniffer().sniff(raw[:4000])
            df = pd.read_csv(io.StringIO(raw), sep=dialect.delimiter)
        except Exception:
            for sep in [",",";","|","\t"]:
                try:
                    df = pd.read_csv(io.StringIO(raw), sep=sep)
                    break
                except Exception:
                    df = None
        if isinstance(df, pd.DataFrame) and not df.empty:
            out.append(df)
    except Exception:
        pass

    # HTML (cho .xls chứa HTML)
    try:
        bio.seek(0)
        html = bio.read().decode(errors="ignore")
        tables = pd.read_html(html, flavor="bs4")
        out.extend([t for t in tables if isinstance(t, pd.DataFrame) and not t.empty])
    except Exception:
        pass

    return out

def _flatten_header(df):
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [" ".join([str(c) for c in col if str(c) != "nan"]).strip() for col in df.columns]
    else:
        df.columns = [str(c) for c in df.columns]
    return df

def _score_table(df: pd.DataFrame) -> float:
    if df is None or df.empty: return 0.0
    df = _flatten_header(df.copy())
    headers = " ".join([c.lower() for c in df.columns])
    score = 0.0
    for k in ["recipient","người nhận","nguoi nhan","beneficiary","payee","receiver","amount","ccy","currency","prepared date","value date"]:
        if k in headers: score += 1.5

    # tỉ lệ cột parse được số (nghi amount)
    num_best = 0
    for c in df.columns:
        try:
            ser = df[c].astype(str).head(200).apply(parse_vn_number)
            ratio = (ser.notna() & (ser != 0)).mean()
            num_best = max(num_best, ratio)
        except Exception: continue
    score += num_best

    # tỉ lệ cột trông giống CCY
    ccy_best = 0
    for c in df.columns:
        vals = df[c].astype(str).head(200).apply(lambda x: clean_ccy(x))
        ratio = (vals.apply(lambda x: bool(re.fullmatch(r"[A-Z]{3}", x)))).mean()
        ccy_best = max(ccy_best, ratio)
    score += ccy_best
    # kích thước
    score += min(df.shape[1]/20.0, 2.0)
    return score

def _find_col_by_keywords(df, exact, contains=()):
    cols = {str(c).strip().lower(): c for c in df.columns}
    for k in exact:
        if k in cols: return cols[k]
    for k in list(exact)+list(contains):
        for ck, oc in cols.items():
            if k in ck: return oc
    return None

def _infer_name_col(df):
    best_col, best_ratio = None, 0
    for c in df.columns:
        ser = df[c].astype(str).head(400).apply(lambda x: " ".join(normalize_name(x)))
        # điểm “giống tên”: có >=2 token chữ, không phải toàn số
        def is_name(s):
            toks = [t for t in s.split() if t]
            if len(toks) < 2: return False
            # chứa ít nhất 2 token alphabetic
            return sum(t.isalpha() for t in toks) >= 2
        ratio = ser.apply(is_name).mean()
        if ratio > best_ratio:
            best_ratio, best_col = ratio, c
    return best_col if best_ratio >= 0.2 else None

def _infer_amount_col(df):
    best_col, best_ratio = None, 0
    for c in df.columns:
        try:
            ser = df[c].astype(str).head(400).apply(parse_vn_number)
            ratio = (ser.notna()).mean()
            if ratio > best_ratio:
                best_col, best_ratio = c, ratio
        except Exception:
            continue
    return best_col

def _infer_ccy_col(df):
    best_col, best_ratio = None, 0
    for c in df.columns:
        vals = df[c].astype(str).head(400).apply(clean_ccy)
        ratio = (vals.apply(lambda x: bool(re.fullmatch(r"[A-Z]{3}", x)))).mean()
        if ratio > best_ratio:
            best_col, best_ratio = c, ratio
    return best_col if best_ratio >= 0.3 else None

def _infer_date_col(df):
    best_col, best_ratio = None, 0
    for c in df.columns:
        try:
            parsed = pd.to_datetime(df[c], errors="coerce", dayfirst=True)
            ratio = parsed.notna().mean()
            if ratio > best_ratio:
                best_col, best_ratio = c, ratio
        except Exception:
            continue
    return best_col

def read_history(uploaded_file) -> pd.DataFrame:
    empty = pd.DataFrame(columns=["recipient","amount","ccy","prepared date"])
    tables = _read_all_tables(uploaded_file)
    if not tables:
        st.error("Không đọc được file lịch sử. Vui lòng kiểm tra định dạng (Excel/CSV/HTML).")
        return empty

    # Chọn bảng tốt nhất
    best = max(tables, key=_score_table)

    df = _flatten_header(best.copy())

    # Loại dòng header rơi vào data
    def _row_is_header_like(row):
        txt = " ".join(map(str, row.values))
        txt = HTML_TAG_RE.sub(" ", txt).lower()
        keys = ["message key","receiver","amount","người nhận","recipient","prepared date","ccy","currency","remark"]
        return sum(k in txt for k in keys) >= 3
    try:
        df = df[~df.apply(_row_is_header_like, axis=1)]
    except Exception:
        pass

    # Tìm cột theo từ khoá
    recip_exact = ["recipient","người nhận","nguoi nhan","beneficiary","payee","receiver name","creditor name","account name","name"]
    recip_contains = ["nguoi","nhan","beneficiar","payee","receiver","creditor","account","name"]
    amt_exact   = ["amount","số tiền","so tien","value","gia tri","amt"]
    ccy_exact   = ["ccy","currency","mã tiền","ma tien","cur","tiền tệ"]
    date_exact  = ["prepared date","value date","post date","posting date","transaction date","tx date","ngày","date"]

    rcol = _find_col_by_keywords(df, recip_exact, recip_contains) or _infer_name_col(df)
    acol = _find_col_by_keywords(df, amt_exact) or _infer_amount_col(df)
    ccol = _find_col_by_keywords(df, ccy_exact) or _infer_ccy_col(df)
    dcol = _find_col_by_keywords(df, date_exact) or _infer_date_col(df)

    out = pd.DataFrame(columns=["recipient","amount","ccy","prepared date"])
    if rcol is not None:
        out["recipient"] = (
            df[rcol]
            .astype(str)
            .str.replace(NBSP, " ", regex=False)
            .apply(lambda s: HTML_TAG_RE.sub(" ", s))
            .str.strip()
        )
    if acol is not None:
        out["amount"] = df[acol].apply(parse_vn_number).astype(float)
    if ccol is not None:
        out["ccy"] = df[ccol].apply(clean_ccy)
    else:
        out["ccy"] = ""
    if dcol is not None:
        out["prepared date"] = pd.to_datetime(df[dcol], dayfirst=True, errors="coerce")
    else:
        out["prepared date"] = pd.NaT

    # loại rỗng
    out = out[out["recipient"].astype(str).str.strip() != ""]
    out = out[out["amount"].fillna(0).astype(float) != 0]
    return out.reset_index(drop=True)

# =========================
# 🧰 UI HELPERS
# =========================
_key_counter = count(1)
def unique_key(prefix: str) -> str: return f"{prefix}_{next(_key_counter)}"

def inline_input(label_text, widget_fn, *args, key_prefix=None, **kwargs):
    left, right = st.columns([0.38, 0.62])
    with left: st.markdown(f"**{label_text}**")
    with right:
        kwargs.setdefault("label_visibility", "collapsed")
        if "key" not in kwargs:
            base = key_prefix or label_text.replace(" ", "_").lower()
            kwargs["key"] = unique_key(base)
        return widget_fn("", *args, **kwargs)

# =========================
# 1. NGƯỜI GỬI | 2. NGƯỜI NHẬN
# =========================
ISO_COUNTRIES = get_iso2_country_codes()
COUNTRY_LABELS = [x[1] for x in ISO_COUNTRIES]
CURRENCY_CODES = get_iso4217_codes()

left_col, right_col = st.columns(2)

with left_col:
    st.subheader("1. Người gửi")
    send_date = inline_input("Ngày gửi tiền", st.date_input, value=date.today(), format="DD/MM/YYYY", key_prefix="send_date")
    pay_method = inline_input("Hình thức thanh toán", st.radio, options=["Tiền mặt","Chuyển khoản"], horizontal=True, index=0, key_prefix="pay_method")
    s_acc = s_acc_name = s_acc_bank = ""
    if pay_method == "Chuyển khoản":
        s_acc = inline_input("Số tài khoản", st.text_input, key_prefix="sender_acc")
        s_acc_name = inline_input("Tên tài khoản", st.text_input, key_prefix="sender_acc_name")
        s_acc_bank = inline_input("Tại ngân hàng", st.text_input, key_prefix="sender_acc_bank")
    s_full = inline_input("Họ tên", st.text_input, key_prefix="sender_full")
    s_addr = inline_input("Địa chỉ", st.text_area, height=80, key_prefix="sender_addr")
    s_country_label = inline_input("Quốc gia", st.selectbox, options=COUNTRY_LABELS,
                                   index=COUNTRY_LABELS.index("VN – Viet Nam") if "VN – Viet Nam" in COUNTRY_LABELS else 0,
                                   key_prefix="sender_country")
    s_country = s_country_label.split("–")[0].strip()
    s_id_type = inline_input("Loại giấy tờ", st.selectbox, options=["CCCD","CC","Passport","Khác (tự nhập)"], index=0, key_prefix="sender_id_type")
    s_id_type_other = inline_input("Giấy tờ khác", st.text_input, key_prefix="sender_id_type_other") if s_id_type=="Khác (tự nhập)" else ""
    s_id_no = inline_input("Số giấy tờ", st.text_input, key_prefix="sender_id_no")
    s_id_issue = inline_input("Ngày cấp", st.date_input, format="DD/MM/YYYY", key_prefix="sender_id_issue")
    s_phone = inline_input("Số điện thoại", st.text_input, key_prefix="sender_phone")

with right_col:
    st.subheader("2. Người nhận")
    r_full = inline_input("Họ tên", st.text_input, key_prefix="recv_full")
    r_acc = inline_input("Số tài khoản", st.text_input, key_prefix="recv_acc")
    r_addr = inline_input("Địa chỉ", st.text_area, height=80, key_prefix="recv_addr")
    r_cc_choice = inline_input("Mã quốc gia", st.selectbox, options=COUNTRY_LABELS,
                               index=COUNTRY_LABELS.index("VN – Viet Nam") if "VN – Viet Nam" in COUNTRY_LABELS else 0,
                               key_prefix="recv_cc")
    r_cc = r_cc_choice.split("–")[0].strip()
    r_id_type = inline_input("Loại giấy tờ (tuỳ chọn)", st.selectbox,
                             options=["(Để trống)","CCCD","CC","Passport","Khác (tự nhập)"], index=0, key_prefix="recv_id_type")
    r_id_type_other = inline_input("Giấy tờ khác", st.text_input, key_prefix="recv_id_type_other") if r_id_type=="Khác (tự nhập)" else ""
    r_id_no = inline_input("Số giấy tờ (tuỳ chọn)", st.text_input, key_prefix="recv_id_no")

# =========================
# 3–6 (hai cột)
# =========================
secL, secR = st.columns(2)

with secL:
    st.subheader("3. Ngân hàng")
    inter_bank = inline_input("Ngân hàng trung gian", st.text_input, key_prefix="inter_bank")
    inter_swift = inline_input("SWIFT trung gian", st.text_input, key_prefix="inter_swift")
    ben_bank = inline_input("Ngân hàng nhận tiền", st.text_input, key_prefix="ben_bank")
    ben_swift = inline_input("SWIFT nhận tiền", st.text_input, key_prefix="ben_swift")

    st.subheader("4. Hồ sơ cung cấp")
    doc_opts = ["CCCD","Giấy khai sinh","Passport","Visa","Thông báo học phí","Khác"]
    docs = inline_input("Chọn loại hồ sơ", st.multiselect, options=doc_opts, default=[], key_prefix="docs")
    doc_counts = {}
    if docs:
        for d in docs:
            doc_counts[d] = inline_input(f"Số lượng '{d}'", st.number_input, min_value=1, value=1, step=1, key_prefix=f"doc_count_{d}")

with secR:
    st.subheader("5. Mục đích và số tiền")
    CURRENCY_CODES = get_iso4217_codes()
    pay_type = inline_input("Loại thanh toán (Cá nhân)", st.selectbox, options=["Trợ cấp","Học phí","Mục đích khác"], index=0, key_prefix="pay_type")
    purpose_desc = inline_input("Nội dung chuyển tiền", st.text_area, height=80, key_prefix="purpose")
    currency = inline_input("Mã tiền tệ", st.selectbox, options=CURRENCY_CODES,
                            index=CURRENCY_CODES.index("USD") if "USD" in CURRENCY_CODES else 0, key_prefix="currency")
    amt_str = inline_input("Số tiền ngoại tệ (VN: 1.234.567,89)", st.text_input, key_prefix="amt")
    vnd_per_ngt_str = inline_input("Tỷ giá VND/NGT (VND cho 1 NGT)", st.text_input, value="0", key_prefix="vnd_ngt")
    vnd_per_usd_str = inline_input("Tỷ giá VND/USD (VND cho 1 USD)", st.text_input, value="0", key_prefix="vnd_usd")
    fee_str = inline_input("Phí dịch vụ (VND)", st.text_input, value="0", key_prefix="fee")
    telex_str = inline_input("Điện phí (VND)", st.text_input, value="0", key_prefix="telex")

    foreign_amt = parse_vn_number(amt_str or "0")
    vnd_per_ngt = parse_vn_number(vnd_per_ngt_str or "0")
    vnd_per_usd = parse_vn_number(vnd_per_usd_str or "0")
    fee = parse_vn_number(fee_str or "0"); telex = parse_vn_number(telex_str or "0")
    vnd_amount = round(foreign_amt * vnd_per_ngt, 0)
    total_vnd = vnd_amount + fee + telex
    usd_current = to_usd(foreign_amt, vnd_per_ngt, vnd_per_usd)

    c1, c2, c3 = st.columns(3)
    with c1: st.metric("Quy đổi (VND)", fmt_vn_int(vnd_amount))
    with c2: st.metric("Tổng thu (VND)", fmt_vn_int(total_vnd))
    with c3: st.metric("Giá trị hiện tại (USD)", fmt_usd(usd_current))

# =========================
# 6. LỊCH SỬ CHUYỂN TIỀN
# =========================
st.subheader("6. Lịch sử chuyển tiền")
hist_file = st.file_uploader(
    "Tải file lịch sử (.xls/.xlsx/.csv/.html). App sẽ tự bóc HTML embedded trong .xls nếu có.",
    type=["xls","xlsx","csv","html","htm"],
    key=unique_key("hist_upload")
)
hist_df = read_history(hist_file)

# =========================
# 🔎 Kiểm tra hạn mức (chỉ hiện khi Trợ cấp)
# =========================
st.markdown("---")
check_btn = st.button("✅ Kiểm tra hạn mức (GDP/người, quy đổi USD)", key=unique_key("check_btn")) if pay_type == "Trợ cấp" else None

cap_usd = cap_year_used = remain_usd = None
summary_df = pd.DataFrame(columns=["Recipient","CCY","Amount_Total","Amount_Total_USD"])
total_usd_all = 0.0
warning_text = ""

if check_btn and r_full and hist_df is not None:
    # GDP/người theo năm gửi
    cap_usd, cap_year_used = fetch_gdp_per_capita_usd(r_cc, send_date.year) if r_cc else (None, None)
    with st.expander("Hạn mức trợ cấp tối đa (GDP/người, USD)", expanded=True):
        if cap_usd is not None: st.write(f"**{r_cc} – năm {cap_year_used}: {fmt_usd(cap_usd)} USD**")
        else: st.warning("Không lấy được GDP/người từ World Bank.")

    # Lọc lịch sử THEO NGƯỜI NHẬN (không lọc theo năm)
    if not hist_df.empty and "recipient" in hist_df.columns and "amount" in hist_df.columns:
        matched = hist_df[hist_df["recipient"].astype(str).apply(lambda x: names_loose_match(x, r_full))].copy()
    else:
        matched = pd.DataFrame()

    if not matched.empty:
        matched["ccy_eff"] = matched.get("ccy","").apply(lambda x: x if isinstance(x,str) and re.fullmatch(r"[A-Z]{3}", x) else "").replace("", currency)

        # Chỉ hỏi tỷ giá cho CCY ≠ USD xuất hiện thực tế
        nonusd = sorted({c for c in matched["ccy_eff"].unique().tolist() if c != "USD"})
        extra_rates = {}
        if nonusd:
            st.caption("Nhập tỷ giá **VND/CCY** cho các CCY xuất hiện (khác USD):")
            cols = st.columns(min(3, len(nonusd)))
            for i, ccy in enumerate(nonusd):
                with cols[i % len(cols)]:
                    val = st.text_input(f"VND/{ccy}", key=unique_key(f"rate_{ccy}"))
                    extra_rates[ccy] = parse_vn_number(val) if val else 0.0

        def row_to_usd(row):
            amt, ccy_row = row["amount"], row["ccy_eff"]
            if ccy_row == "USD": return float(amt) if pd.notna(amt) else 0.0
            if ccy_row == currency: return to_usd(amt, vnd_per_ngt, vnd_per_usd)
            return to_usd(amt, extra_rates.get(ccy_row, 0.0), vnd_per_usd)

        matched["usd"] = matched.apply(row_to_usd, axis=1)

        grp = matched.groupby("ccy_eff", dropna=False).agg(
            Amount_Total=("amount","sum"),
            Amount_Total_USD=("usd","sum")
        ).reset_index().rename(columns={"ccy_eff":"CCY"})
        grp["Recipient"] = r_full
        summary_df = grp[["Recipient","CCY","Amount_Total","Amount_Total_USD"]]
        total_usd_all = float(summary_df["Amount_Total_USD"].sum())
    else:
        st.info("Không tìm thấy giao dịch nào khớp **tên người nhận** trong lịch sử.")

    with st.expander("Bảng cộng dồn theo CCY (lọc đúng người nhận & quy đổi USD)", expanded=True):
        st.dataframe(summary_df, use_container_width=True)
        st.write(f"**TỔNG ĐÃ CHUYỂN (USD)**: {fmt_usd(total_usd_all)}")

    if cap_usd is not None:
        remain_usd = cap_usd - total_usd_all
        st.write(f"**Số còn được chuyển (USD)** = {fmt_usd(remain_usd)}")
        if to_usd(foreign_amt, vnd_per_ngt, vnd_per_usd) > remain_usd or remain_usd < 0:
            st.error("**🚨 CHUYỂN VƯỢT HẠN MỨC**")
            warning_text = "CHUYỂN VƯỢT HẠN MỨC"

# =========================
# 📤 XUẤT EXCEL (điền ô bên cạnh tiêu đề + sheet Summary)
# =========================
st.markdown("---"); st.subheader("Xuất Excel")
template = st.file_uploader("(Khuyến nghị) Tải file Excel **mẫu in lệnh**. Hệ thống sẽ tìm các ô tiêu đề và điền **ô bên cạnh**.",
                            type=["xlsx","xls"], key=unique_key("template_upload"))

def compose_row_dict():
    docs_list = []
    try:
        for k in (docs or []):
            docs_list.append(f"{k} x{int(st.session_state.get(f'doc_count_{k}',1))}")
    except Exception:
        pass
    docs_str = ", ".join(docs_list)

    return {
        "Ngày gửi": fmt_ddmmyyyy(send_date),
        "Hình thức thanh toán": pay_method,
        "Số tài khoản": s_acc if pay_method=="Chuyển khoản" else "",
        "Tên tài khoản": s_acc_name if pay_method=="Chuyển khoản" else "",
        "Tại ngân hàng": s_acc_bank if pay_method=="Chuyển khoản" else "",
        "Họ tên người gửi": s_full,
        "Địa chỉ người gửi": s_addr,
        "Quốc gia người gửi (mã ISO-2)": s_country,
        "Loại giấy tờ người gửi": s_id_type if s_id_type!="Khác (tự nhập)" else s_id_type_other,
        "Số giấy tờ người gửi": s_id_no,
        "Ngày cấp GTTT người gửi": fmt_ddmmyyyy(s_id_issue),
        "SĐT người gửi": s_phone,
        "Họ tên người nhận": r_full,
        "Số tài khoản người nhận": r_acc,
        "Địa chỉ người nhận": r_addr,
        "Mã quốc gia người nhận": r_cc,
        "Loại giấy tờ người nhận": r_id_type if r_id_type!="Khác (tự nhập)" else r_id_type_other,
        "Số giấy tờ người nhận": r_id_no,
        "Ngân hàng trung gian": inter_bank,
        "SWIFT trung gian": inter_swift,
        "Ngân hàng nhận tiền": ben_bank,
        "SWIFT nhận tiền": ben_swift,
        "Loại thanh toán (Cá nhân)": pay_type,
        "Nội dung chuyển tiền": purpose_desc,
        "Hồ sơ cung cấp": docs_str,
        "Mã tiền tệ": currency,
        "Số tiền ngoại tệ": foreign_amt,
        "Tỷ giá VND/NGT": vnd_per_ngt,
        "Tỷ giá VND/USD": vnd_per_usd,
        "Số tiền quy đổi (VND)": int(round(foreign_amt*vnd_per_ngt,0)),
        "Phí dịch vụ (VND)": int(round(fee,0)),
        "Điện phí (VND)": int(round(telex,0)),
        "Tổng thu (VND)": int(round(foreign_amt*vnd_per_ngt + fee + telex,0)),
        "Giá trị giao dịch hiện tại (USD)": to_usd(foreign_amt, vnd_per_ngt, vnd_per_usd),
        "Hạn mức (GDP/người, USD)": cap_usd if cap_usd is not None else "",
        "Năm áp dụng hạn mức": cap_year_used if cap_year_used is not None else "",
        "TỔNG ĐÃ CHUYỂN (USD)": total_usd_all,
        "Số còn được chuyển (USD)": remain_usd if remain_usd is not None else "",
        "Cảnh báo": warning_text,
    }

def export_excel_fill_template(template_file, mapping: dict, summary: pd.DataFrame | None) -> bytes:
    df_map = pd.DataFrame([mapping])
    df_sum = summary.copy() if isinstance(summary, pd.DataFrame) and not summary.empty else pd.DataFrame(
        columns=["Recipient","CCY","Amount_Total","Amount_Total_USD"])
    if template_file is None:
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            df_map.to_excel(w, index=False, sheet_name="Lenh_Chuyen_Tien")
            df_sum.to_excel(w, index=False, sheet_name="Summary")
        out.seek(0); return out.read()
    bio = io.BytesIO(template_file.read())
    bio.seek(0); wb = load_workbook(bio)
    titles = set(mapping.keys())
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                if isinstance(cell.value, str):
                    key = str(cell.value).strip()
                    if key in titles:
                        ws.cell(row=cell.row, column=cell.column+1, value=mapping[key])
    if "Lenh_Chuyen_Tien" in wb.sheetnames: wb.remove(wb["Lenh_Chuyen_Tien"])
    ws1 = wb.create_sheet("Lenh_Chuyen_Tien")
    for r in dataframe_to_rows(df_map, index=False, header=True): ws1.append(r)
    if "Summary" in wb.sheetnames: wb.remove(wb["Summary"])
    ws2 = wb.create_sheet("Summary")
    for r in dataframe_to_rows(df_sum, index=False, header=True): ws2.append(r)
    out = io.BytesIO(); wb.save(out); out.seek(0); return out.read()

row_dict = compose_row_dict()
excel_bytes = export_excel_fill_template(template, row_dict, summary_df)
st.download_button(
    "⬇️ Tải file Excel (điền ô bên cạnh tiêu đề & sheet Summary)",
    data=excel_bytes,
    file_name=f"lenh_chuyen_tien_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key=unique_key("download_btn")
)
