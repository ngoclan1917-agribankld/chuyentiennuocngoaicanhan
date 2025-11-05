import streamlit as st
import pandas as pd
import io
import requests
from datetime import date, datetime
from unidecode import unidecode
import math
from itertools import count

# Optional nhưng nên có
try:
    import pycountry
except Exception:
    pycountry = None

from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

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
# 🧩 HÀM TIỆN ÍCH
# =========================
def parse_vn_number(s: str) -> float:
    """Parse số kiểu Việt Nam: '1.234.567,89' -> 1234567.89; cũng chấp nhận '1234.56'."""
    if s is None:
        return 0.0
    s = str(s).strip()
    if s == "":
        return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s and "." not in s:
        s = s.replace(",", ".")
    return float(s)

def fmt_vn_int(n: float | int) -> str:
    try:
        return f"{int(round(float(n), 0)):,}".replace(",", ".")
    except Exception:
        return "0"

def fmt_usd(n: float | int) -> str:
    try:
        return f"{float(n):,.2f}"
    except Exception:
        return "0.00"

def fmt_ddmmyyyy(d):
    if isinstance(d, (date, datetime)):
        return d.strftime("%d/%m/%Y")
    return ""

def normalize_name(name: str) -> set:
    if not isinstance(name, str):
        return set()
    name = unidecode(name).lower().strip()
    tokens = [t for t in name.replace(",", " ").split() if t]
    return set(tokens)

def tokens_match(a: str, b: str) -> bool:
    ta, tb = normalize_name(a), normalize_name(b)
    return (ta == tb) and len(ta) > 0

def get_iso2_country_codes():
    items = []
    if pycountry:
        try:
            for c in pycountry.countries:
                items.append((c.alpha_2.upper(), f"{c.alpha_2.upper()} – {c.name}"))
        except Exception:
            pass
    if not items:  # fallback rút gọn
        fallback = {
            "VN": "Viet Nam", "US": "United States", "AU": "Australia", "JP": "Japan",
            "KR": "Korea, Republic of", "SG": "Singapore", "CN": "China", "DE": "Germany",
            "FR": "France", "GB": "United Kingdom", "TH": "Thailand", "CA": "Canada"
        }
        items = [(k, f"{k} – {v}") for k, v in fallback.items()]
    items.sort(key=lambda x: x[0])
    return items

def get_iso4217_codes():
    codes = set()
    if pycountry:
        try:
            for c in pycountry.currencies:
                if getattr(c, "alpha_3", None):
                    codes.add(c.alpha_3.upper())
        except Exception:
            pass
    # fallback tối thiểu
    if not codes:
        codes = {
            "USD","EUR","JPY","GBP","AUD","CAD","CHF","CNY","HKD","SGD","KRW",
            "THB","TWD","MYR","IDR","INR","VND","NZD","SEK","NOK","DKK","RUB",
            "AED","SAR","QAR","KWD","BHD","TRY","BRL","MXN","ZAR","PLN","HUF",
        }
    return sorted(list(codes))

def fetch_gdp_per_capita_usd(iso2: str, year: int):
    """Trả (value_usd, used_year) với fallback year-1, year-2; nếu không có: (None,None)."""
    if not iso2 or not year:
        return None, None
    for y in [year, year - 1, year - 2]:
        url = f"https://api.worldbank.org/v2/country/{iso2.lower()}/indicator/NY.GDP.PCAP.CD?date={y}:{y}&format=json"
        try:
            r = requests.get(url, timeout=12)
            js = r.json()
            if isinstance(js, list) and len(js) > 1 and js[1]:
                val = js[1][0].get("value")
                if val is not None:
                    return float(val), y
        except Exception:
            continue
    return None, None

def safe_read_bytes(uploaded_file):
    """Đọc UploadedFile ra bytes để có thể reset con trỏ & dùng nhiều lần."""
    if uploaded_file is None:
        return None
    b = uploaded_file.read()
    return io.BytesIO(b)

def read_history(file) -> pd.DataFrame:
    """
    Đọc CSV/XLSX robust:
    - Copy sang BytesIO rồi thử read_excel; nếu thất bại -> thử read_csv với nhiều delimiter.
    - Trả về cột chuẩn: recipient, amount, prepared date, currency?
    """
    if file is None:
        return pd.DataFrame(columns=["recipient", "amount", "prepared date", "currency"])

    bio = safe_read_bytes(file)
    if bio is None:
        return pd.DataFrame(columns=["recipient", "amount", "prepared date", "currency"])

    # 1) thử Excel
    try:
        bio.seek(0)
        df = pd.read_excel(bio, engine="openpyxl")
    except Exception:
        # 2) thử CSV với các delimiter
        df = None
        for sep in [",",";","|","\\t"]:
            try:
                bio.seek(0)
                df = pd.read_csv(bio, sep=sep if sep != "\\t" else "\t")
                break
            except Exception:
                continue
        if df is None:
            st.error("Không đọc được file lịch sử. Vui lòng kiểm tra định dạng (CSV hoặc Excel).")
            return pd.DataFrame(columns=["recipient", "amount", "prepared date", "currency"])

    # map cột linh hoạt
    cols = {str(c).strip().lower(): c for c in df.columns}
    def pick(*keys):
        for k in keys:
            for ck, oc in cols.items():
                if ck == k:
                    return oc
        return None

    recipient_col = pick("recipient", "nguoi nhan", "tên người nhận", "ten nguoi nhan")
    amount_col    = pick("amount", "so tien", "giatri", "gia tri")
    date_col      = None
    for ck, oc in cols.items():
        if "prepared" in ck and "date" in ck:
            date_col = oc
            break
    if not date_col:
        for ck, oc in cols.items():
            if ck in ("date", "ngay", "prepared_date"):
                date_col = oc
                break
    currency_col = None
    for ck, oc in cols.items():
        if ck in ("currency", "ma tien", "ma_tien"):
            currency_col = oc
            break

    if not (recipient_col and amount_col and date_col):
        st.warning("File lịch sử cần có cột tối thiểu: recipient, amount, prepared date.")
        return pd.DataFrame(columns=["recipient", "amount", "prepared date", "currency"])

    out = pd.DataFrame({
        "recipient": df[recipient_col].astype(str),
        "amount": df[amount_col],
        "prepared date": pd.to_datetime(df[date_col], dayfirst=True, errors="coerce")
    })
    if currency_col:
        out["currency"] = df[currency_col].astype(str).str.upper().str.strip()
    else:
        out["currency"] = None

    def _amt(x):
        try:
            if isinstance(x, (int, float)) and not pd.isna(x):
                return float(x)
            return parse_vn_number(str(x))
        except Exception:
            return float("nan")
    out["amount"] = out["amount"].apply(_amt)
    return out

def to_usd(amount: float, vnd_per_ccy: float, vnd_per_usd: float) -> float:
    """Quy đổi về USD theo tỷ giá chéo: amount * (VND/CCY) / (VND/USD)."""
    if amount is None or pd.isna(amount):
        return 0.0
    if vnd_per_ccy is None or vnd_per_ccy <= 0 or vnd_per_usd is None or vnd_per_usd <= 0:
        return 0.0
    return float(amount) * float(vnd_per_ccy) / float(vnd_per_usd)

def id_type_value(selected: str, other_text: str) -> str:
    if "Khác" in (selected or "") and (other_text or "").strip():
        return other_text.strip()
    if "(Để trống)" in (selected or ""):
        return ""
    return selected or ""

# =========================
# 🔑 BỘ PHÁT KEY DUY NHẤT
# =========================
_key_counter = count(1)
def unique_key(prefix: str) -> str:
    return f"{prefix}_{next(_key_counter)}"

# =========================
# 🎛️ NHÃN BÊN CẠNH Ô NHẬP (CÓ KEY DUY NHẤT)
# =========================
def inline_input(label_text, widget_fn, *args, key_prefix=None, **kwargs):
    """Hiển thị nhãn bên trái, ô nhập bên phải (cùng hàng) và tự sinh key duy nhất."""
    left, right = st.columns([0.38, 0.62])
    with left:
        st.markdown(f"**{label_text}**")
    with right:
        kwargs.setdefault("label_visibility", "collapsed")
        if "key" not in kwargs:
            base = key_prefix or label_text.replace(" ", "_").lower()
            kwargs["key"] = unique_key(base)
        return widget_fn("", *args, **kwargs)

# =========================
# 🔝 HÀNG TRÊN: 1. NGƯỜI GỬI | 2. NGƯỜI NHẬN
# =========================
left_col, right_col = st.columns(2)

# Danh sách quốc gia & tiền tệ
ISO_COUNTRIES = get_iso2_country_codes()
COUNTRY_LABELS = [x[1] for x in ISO_COUNTRIES]
CURRENCY_CODES = get_iso4217_codes()

with left_col:
    st.subheader("1. Người gửi")
    # date_input format (Streamlit mới hỗ trợ). Nếu bản cũ không hỗ trợ, vẫn hiển thị theo locale,
    # còn khi xuất Excel mình sẽ format dd/mm/yyyy.
    send_date = inline_input("Ngày gửi tiền", st.date_input, value=date.today(),
                             format="DD/MM/YYYY", key_prefix="send_date")
    pay_method = inline_input("Hình thức thanh toán", st.radio,
                              options=["Tiền mặt", "Chuyển khoản"], horizontal=True, index=0, key_prefix="pay_method")
    s_acc = s_acc_name = s_acc_bank = ""
    if pay_method == "Chuyển khoản":
        s_acc = inline_input("Số tài khoản", st.text_input, key_prefix="sender_acc")
        s_acc_name = inline_input("Tên tài khoản", st.text_input, key_prefix="sender_acc_name")
        s_acc_bank = inline_input("Tại ngân hàng", st.text_input, key_prefix="sender_acc_bank")

    s_full = inline_input("Họ tên", st.text_input, key_prefix="sender_full")
    s_addr = inline_input("Địa chỉ", st.text_area, height=80, key_prefix="sender_addr")

    s_country_label = inline_input("Quốc gia", st.selectbox, options=COUNTRY_LABELS, index=COUNTRY_LABELS.index("VN – Viet Nam") if "VN – Viet Nam" in COUNTRY_LABELS else 0, key_prefix="sender_country")
    s_country = s_country_label.split("–")[0].strip()  # mã ISO-2

    s_id_type = inline_input("Loại giấy tờ", st.selectbox,
                             options=["CCCD", "CC", "Passport", "Khác (tự nhập)"], index=0, key_prefix="sender_id_type")
    s_id_type_other = ""
    if s_id_type == "Khác (tự nhập)":
        s_id_type_other = inline_input("Giấy tờ khác", st.text_input, key_prefix="sender_id_type_other")
    s_id_no = inline_input("Số giấy tờ", st.text_input, key_prefix="sender_id_no")
    s_id_issue = inline_input("Ngày cấp", st.date_input, format="DD/MM/YYYY", key_prefix="sender_id_issue")
    s_phone = inline_input("Số điện thoại", st.text_input, key_prefix="sender_phone")

with right_col:
    st.subheader("2. Người nhận")
    r_full = inline_input("Họ tên", st.text_input, key_prefix="recv_full")
    r_acc = inline_input("Số tài khoản", st.text_input, key_prefix="recv_acc")
    r_addr = inline_input("Địa chỉ", st.text_area, height=80, key_prefix="recv_addr")

    r_cc_choice = inline_input("Mã quốc gia", st.selectbox, options=COUNTRY_LABELS, index=COUNTRY_LABELS.index("VN – Viet Nam") if "VN – Viet Nam" in COUNTRY_LABELS else 0, key_prefix="recv_cc")
    r_cc = r_cc_choice.split("–")[0].strip()

    r_id_type = inline_input("Loại giấy tờ (tuỳ chọn)", st.selectbox,
                             options=["(Để trống)", "CCCD", "CC", "Passport", "Khác (tự nhập)"],
                             index=0, key_prefix="recv_id_type")
    r_id_type_other = ""
    if r_id_type == "Khác (tự nhập)":
        r_id_type_other = inline_input("Giấy tờ khác", st.text_input, key_prefix="recv_id_type_other")
    r_id_no = inline_input("Số giấy tờ (tuỳ chọn)", st.text_input, key_prefix="recv_id_no")

# =========================
# ⬇️ HÀNG DƯỚI: 3–6 CHIA 2 BÊN CHO CÂN ĐỐI
# =========================
secL, secR = st.columns(2)

with secL:
    st.subheader("3. Ngân hàng")
    inter_bank = inline_input("Ngân hàng trung gian", st.text_input, key_prefix="inter_bank")
    inter_swift = inline_input("SWIFT trung gian", st.text_input, key_prefix="inter_swift")
    ben_bank = inline_input("Ngân hàng nhận tiền", st.text_input, key_prefix="ben_bank")
    ben_swift = inline_input("SWIFT nhận tiền", st.text_input, key_prefix="ben_swift")

    st.subheader("4. Hồ sơ cung cấp")
    doc_opts = ["CCCD", "Giấy khai sinh", "Passport", "Visa", "Thông báo học phí", "Khác"]
    docs = inline_input("Chọn loại hồ sơ", st.multiselect, options=doc_opts, default=[], key_prefix="docs")
    doc_counts = {}
    if docs:
        for d in docs:
            doc_counts[d] = inline_input(f"Số lượng '{d}'", st.number_input,
                                         min_value=1, value=1, step=1, key_prefix=f"doc_count_{d}")

with secR:
    st.subheader("5. Mục đích và số tiền")
    pay_type = inline_input("Loại thanh toán (Cá nhân)", st.selectbox,
                            options=["Trợ cấp", "Học phí", "Mục đích khác"], index=0, key_prefix="pay_type")
    purpose_desc = inline_input("Nội dung chuyển tiền", st.text_area, height=80, key_prefix="purpose")

    currency = inline_input("Mã tiền tệ", st.selectbox, options=CURRENCY_CODES,
                            index=CURRENCY_CODES.index("USD") if "USD" in CURRENCY_CODES else 0,
                            key_prefix="currency")

    amt_str = inline_input("Số tiền ngoại tệ (VN: 1.234.567,89)", st.text_input, key_prefix="amt")
    vnd_per_ngt_str = inline_input("Tỷ giá VND/NGT (VND cho 1 NGT)", st.text_input, value="0", key_prefix="vnd_ngt")
    vnd_per_usd_str = inline_input("Tỷ giá VND/USD (VND cho 1 USD)", st.text_input, value="0", key_prefix="vnd_usd")
    fee_str = inline_input("Phí dịch vụ (VND)", st.text_input, value="0", key_prefix="fee")
    telex_str = inline_input("Điện phí (VND)", st.text_input, value="0", key_prefix="telex")

    # Parse
    try:
        foreign_amt = parse_vn_number(amt_str) if amt_str else 0.0
        vnd_per_ngt = parse_vn_number(vnd_per_ngt_str) if vnd_per_ngt_str else 0.0
        vnd_per_usd = parse_vn_number(vnd_per_usd_str) if vnd_per_usd_str else 0.0
        fee = parse_vn_number(fee_str) if fee_str else 0.0
        telex = parse_vn_number(telex_str) if telex_str else 0.0
    except Exception:
        st.error("Vui lòng kiểm tra lại định dạng số (dùng '.' cho nghìn và ',' cho thập phân).")
        foreign_amt, vnd_per_ngt, vnd_per_usd, fee, telex = 0.0, 0.0, 0.0, 0.0, 0.0

    vnd_amount = round(foreign_amt * vnd_per_ngt, 0)
    total_vnd = vnd_amount + fee + telex
    usd_current = to_usd(foreign_amt, vnd_per_ngt, vnd_per_usd)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Quy đổi (VND)", fmt_vn_int(vnd_amount))
    with c2:
        st.metric("Tổng thu (VND)", fmt_vn_int(total_vnd))
    with c3:
        st.metric("Giá trị hiện tại (USD)", fmt_usd(usd_current))

# =========================
# 6. LỊCH SỬ CHUYỂN TIỀN & TỶ GIÁ PHỤ
# =========================
st.subheader("6. Lịch sử chuyển tiền")
hist_file = st.file_uploader(
    "Tải file CSV/XLSX có cột: recipient, amount, prepared date (tuỳ chọn: currency)",
    type=["csv", "xlsx", "xls"],
    key=unique_key("hist_upload")
)
hist_df = read_history(hist_file)

rates_map = {}
if not hist_df.empty and hist_df["currency"].notna().any():
    st.info("Đã phát hiện cột 'currency' trong lịch sử—hãy nhập tỷ giá VND/<mã> cho từng loại tiền.")
    uniq_ccy = sorted([c for c in hist_df["currency"].dropna().unique().tolist() if c and c != "None"])
    cols = st.columns(min(3, len(uniq_ccy)) if uniq_ccy else 1)
    for idx, ccy in enumerate(uniq_ccy):
        with cols[idx % len(cols)]:
            val = st.text_input(f"VND/{ccy}", key=unique_key(f"rate_{ccy}"))
            try:
                rates_map[ccy] = parse_vn_number(val) if val else 0.0
            except Exception:
                rates_map[ccy] = 0.0

# =========================
# 🔎 NÚT KIỂM TRA HẠN MỨC (chỉ hiện khi Trợ cấp)
# =========================
st.markdown("---")
check_btn = None
if pay_type == "Trợ cấp":
    check_btn = st.button("✅ Kiểm tra hạn mức (GDP/người, quy đổi USD)", key=unique_key("check_btn"))

cap_usd = cap_year_used = sent_sum_usd = remain_usd = None
warning_text = ""

if check_btn and r_full and r_cc and send_date:
    # Lấy GDP/người
    cap_usd, cap_year_used = fetch_gdp_per_capita_usd(r_cc, send_date.year)
    with st.expander("Hạn mức trợ cấp tối đa một năm (GDP/người, USD)", expanded=True):
        if cap_usd is not None:
            st.write(f"**GDP/người** của **{r_cc}** cho **năm {cap_year_used}**: **{fmt_usd(cap_usd)} USD**")
        else:
            st.error("Không lấy được GDP/người từ World Bank cho mã quốc gia/năm này.")

    # Cộng dồn USD theo năm
    if not hist_df.empty:
        same_year = hist_df[hist_df["prepared date"].dt.year == send_date.year]
        mask = same_year["recipient"].astype(str).apply(lambda x: tokens_match(x, r_full))
        matched = same_year.loc[mask].copy()

        def row_to_usd(row):
            amt = row["amount"]
            row_ccy = row.get("currency", None)
            if pd.isna(row_ccy) or not row_ccy or row_ccy == "None":
                # mặc định cùng loại nguyên tệ NGT
                return to_usd(amt, vnd_per_ngt, vnd_per_usd)
            # có currency riêng -> cần VND/<row_ccy>
            v_row = rates_map.get(str(row_ccy).upper(), 0.0)
            return to_usd(amt, v_row, vnd_per_usd)

        matched["usd"] = matched.apply(row_to_usd, axis=1)
        sent_sum_usd = float(matched["usd"].sum())
    else:
        sent_sum_usd = 0.0

    with st.expander("Số tiền đã chuyển trong năm (sau quy đổi USD)", expanded=True):
        st.write(f"**ĐÃ CHUYỂN NĂM {send_date.year}: {fmt_usd(sent_sum_usd)} USD**")

    if cap_usd is not None:
        remain_usd = cap_usd - sent_sum_usd
        st.write(f"**Số còn được chuyển (USD)** = {fmt_usd(remain_usd)}")
        if usd_current > remain_usd or (remain_usd is not None and remain_usd < 0):
            st.error("**🚨 CHUYỂN VƯỢT HẠN MỨC**")
            warning_text = "CHUYỂN VƯỢT HẠN MỨC"

# =========================
# ⬇️ XUẤT EXCEL (ĐIỀN NGAY BÊN CẠNH Ô TIÊU ĐỀ)
# =========================
st.markdown("---")
st.subheader("Xuất Excel")

template = st.file_uploader(
    "(Khuyến nghị) Tải file Excel **mẫu in lệnh**. Hệ thống sẽ tìm 'ô tiêu đề' và điền **ô bên cạnh** cùng hàng.",
    type=["xlsx", "xls"],
    key=unique_key("template_upload")
)

# Lập bảng dữ liệu xuất
def compose_row_dict():
    return {
        "Ngày gửi": fmt_ddmmyyyy(send_date),
        "Hình thức thanh toán": pay_method,
        "Số tài khoản": s_acc if pay_method == "Chuyển khoản" else "",
        "Tên tài khoản": s_acc_name if pay_method == "Chuyển khoản" else "",
        "Tại ngân hàng": s_acc_bank if pay_method == "Chuyển khoản" else "",
        "Họ tên người gửi": s_full,
        "Địa chỉ người gửi": s_addr,
        "Quốc gia người gửi (mã ISO-2)": s_country,
        "Loại giấy tờ người gửi": id_type_value(s_id_type, s_id_type_other),
        "Số giấy tờ người gửi": s_id_no,
        "Ngày cấp GTTT người gửi": fmt_ddmmyyyy(s_id_issue),
        "SĐT người gửi": s_phone,

        "Họ tên người nhận": r_full,
        "Số tài khoản người nhận": r_acc,
        "Địa chỉ người nhận": r_addr,
        "Mã quốc gia người nhận": r_cc,
        "Loại giấy tờ người nhận": id_type_value(r_id_type, r_id_type_other),
        "Số giấy tờ người nhận": r_id_no,

        "Ngân hàng trung gian": inter_bank,
        "SWIFT trung gian": inter_swift,
        "Ngân hàng nhận tiền": ben_bank,
        "SWIFT nhận tiền": ben_swift,

        "Loại thanh toán (Cá nhân)": pay_type,
        "Nội dung chuyển tiền": purpose_desc,
        "Hồ sơ cung cấp": ", ".join([f"{k} x{doc_counts.get(k,1)}" for k in (docs or [])]),

        "Mã tiền tệ": currency,
        "Số tiền ngoại tệ": foreign_amt,
        "Tỷ giá VND/NGT": vnd_per_ngt,
        "Tỷ giá VND/USD": vnd_per_usd,
        "Số tiền quy đổi (VND)": int(round(vnd_amount, 0)) if not math.isnan(vnd_amount) else 0,
        "Phí dịch vụ (VND)": int(round(fee, 0)) if not math.isnan(fee) else 0,
        "Điện phí (VND)": int(round(telex, 0)) if not math.isnan(telex) else 0,
        "Tổng thu (VND)": int(round(total_vnd, 0)) if not math.isnan(total_vnd) else 0,

        "Giá trị giao dịch hiện tại (USD)": usd_current if usd_current is not None else "",

        "Hạn mức (GDP/người, USD)": cap_usd if cap_usd is not None else "",
        "Năm áp dụng hạn mức": cap_year_used if cap_year_used is not None else "",
        "Đã chuyển trong năm (USD)": sent_sum_usd if sent_sum_usd is not None else "",
        "Số còn được chuyển (USD)": remain_usd if remain_usd is not None else "",
        "Cảnh báo": warning_text or "",
    }

# Ghi sheet dữ liệu phụ &/hoặc điền bên cạnh tiêu đề trong template
def export_excel_fill_template(template_file, mapping: dict) -> bytes:
    """
    - Nếu có template: mở workbook, tìm TỪNG ô có text == 'tiêu đề' và ghi GIÁ TRỊ vào ô bên phải (cột+1).
      Ghi cho tất cả các sheet (sheet nào có tiêu đề thì điền).
      Đồng thời tạo/ghi sheet 'Lenh_Chuyen_Tien' chứa toàn bộ mapping dạng bảng (tham khảo/đối soát).
    - Nếu không có template: tạo file mới chỉ có sheet 'Lenh_Chuyen_Tien'.
    """
    df = pd.DataFrame([mapping])

    if template_file is None:
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Lenh_Chuyen_Tien")
        out.seek(0)
        return out.read()

    # Đọc template vào workbook
    bio = safe_read_bytes(template_file)
    bio.seek(0)
    wb = load_workbook(bio)

    # Điền dữ liệu: tìm đúng tiêu đề -> ghi sang ô bên phải (cùng hàng, col+1)
    titles = list(mapping.keys())
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                val = cell.value
                if isinstance(val, str):
                    key = val.strip()
                    if key in mapping:
                        target_col = cell.column + 1  # ngay bên cạnh
                        ws.cell(row=cell.row, column=target_col, value=mapping[key])

    # Ghi thêm sheet dữ liệu tổng hợp
    if "Lenh_Chuyen_Tien" in wb.sheetnames:
        ws_old = wb["Lenh_Chuyen_Tien"]
        wb.remove(ws_old)
    ws = wb.create_sheet("Lenh_Chuyen_Tien")
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()

row_dict = compose_row_dict()
excel_bytes = export_excel_fill_template(template, row_dict)

st.download_button(
    label="⬇️ Tải file Excel (điền ngay bên cạnh ô tiêu đề)",
    data=excel_bytes,
    file_name=f"lenh_chuyen_tien_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key=unique_key("download_btn")
)

st.success("Hoàn tất. Định dạng ngày hiển thị và xuất file: dd/mm/yyyy. Đọc file lịch sử đã được vá lỗi (Excel/CSV).")
