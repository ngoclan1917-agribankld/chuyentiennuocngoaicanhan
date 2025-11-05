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
    if not items:
        fallback = {
            "VN": "Viet Nam", "US": "United States", "AU": "Australia", "JP": "Japan",
            "KR": "Korea, Republic of", "SG": "Singapore", "CN": "China", "DE": "Germany",
            "FR": "France", "GB": "United Kingdom", "TH": "Thailand", "CA": "Canada"
        }
        items = [(k, f"{k} – {v}") for k, v in fallback.items()]
    items.sort(key=lambda x: x[0])
    return items

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

def read_history(file) -> pd.DataFrame:
    """Đọc CSV/XLSX, trả về cột chuẩn: recipient, amount, prepared date, currency?"""
    if file is None:
        return pd.DataFrame(columns=["recipient", "amount", "prepared date", "currency"])
    ext = file.name.lower().split(".")[-1]
    if ext in ("xlsx", "xls"):
        df = pd.read_excel(file)
    else:
        df = pd.read_csv(file)

    cols = {c.strip().lower(): c for c in df.columns}
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
            if ck in ("date", "ngay"):
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
    """
    Hiển thị nhãn bên trái, ô nhập bên phải (cùng hàng) và tự sinh key duy nhất.
    Dùng cho mọi widget để tránh StreamlitDuplicateElementId.
    """
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

with left_col:
    st.subheader("1. Người gửi")
    send_date = inline_input("Ngày gửi tiền", st.date_input, value=date.today(), key_prefix="send_date")
    pay_method = inline_input("Hình thức thanh toán", st.radio,
                              options=["Tiền mặt", "Chuyển khoản"], horizontal=True, index=0, key_prefix="pay_method")
    s_acc = ""
    s_acc_name = ""
    s_acc_bank = ""
    if pay_method == "Chuyển khoản":
        s_acc = inline_input("Số tài khoản", st.text_input, key_prefix="sender_acc")
        s_acc_name = inline_input("Tên tài khoản", st.text_input, key_prefix="sender_acc_name")
        s_acc_bank = inline_input("Tại ngân hàng", st.text_input, key_prefix="sender_acc_bank")

    s_full = inline_input("Họ tên", st.text_input, key_prefix="sender_full")
    s_addr = inline_input("Địa chỉ", st.text_area, height=80, key_prefix="sender_addr")
    s_country = inline_input("Quốc gia", st.text_input, key_prefix="sender_country")
    s_id_type = inline_input("Loại giấy tờ", st.selectbox,
                             options=["CCCD", "CC", "Passport", "Khác (tự nhập)"], index=0, key_prefix="sender_id_type")
    s_id_type_other = ""
    if s_id_type == "Khác (tự nhập)":
        s_id_type_other = inline_input("Giấy tờ khác", st.text_input, key_prefix="sender_id_type_other")
    s_id_no = inline_input("Số giấy tờ", st.text_input, key_prefix="sender_id_no")
    s_id_issue = inline_input("Ngày cấp", st.date_input, key_prefix="sender_id_issue")
    s_phone = inline_input("Số điện thoại", st.text_input, key_prefix="sender_phone")

with right_col:
    st.subheader("2. Người nhận")
    r_full = inline_input("Họ tên", st.text_input, key_prefix="recv_full")
    r_acc = inline_input("Số tài khoản", st.text_input, key_prefix="recv_acc")
    r_addr = inline_input("Địa chỉ", st.text_area, height=80, key_prefix="recv_addr")

    iso_list = get_iso2_country_codes()
    r_cc_label = [x[1] for x in iso_list]
    r_cc_choice = inline_input("Mã quốc gia", st.selectbox, options=r_cc_label, index=0, key_prefix="recv_cc")
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

    currency = (inline_input("Mã tiền tệ (ISO-4217)", st.text_input, key_prefix="currency") or "").upper().strip() or "USD"
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
# 🔎 NÚT KIỂM TRA HẠN MỨC (Trợ cấp)
# =========================
st.markdown("---")
check_btn = st.button("✅ Kiểm tra hạn mức (áp dụng khi Loại thanh toán = Trợ cấp)", key=unique_key("check_btn"))

cap_usd = cap_year_used = sent_sum_usd = remain_usd = None
warning_text = ""

if check_btn and pay_type == "Trợ cấp" and r_full and r_cc and send_date:
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
# ⬇️ XUẤT EXCEL (KÈM THEO MẪU)
# =========================
st.markdown("---")
st.subheader("Xuất Excel")
template = st.file_uploader(
    "(Tuỳ chọn) Tải file Excel **mẫu in lệnh** để chèn dữ liệu",
    type=["xlsx", "xls"],
    key=unique_key("template_upload")
)

def compose_row_dict():
    return {
        "send_date": send_date.isoformat() if isinstance(send_date, (date, datetime)) else "",
        "pay_method": pay_method,
        "sender_fullname": s_full,
        "sender_account": s_acc if pay_method == "Chuyển khoản" else "",
        "sender_account_name": s_acc_name if pay_method == "Chuyển khoản" else "",
        "sender_account_bank": s_acc_bank if pay_method == "Chuyển khoản" else "",
        "sender_addr": s_addr,
        "sender_country": s_country,
        "sender_id_type": id_type_value(s_id_type, s_id_type_other),
        "sender_id_no": s_id_no,
        "sender_id_issue_date": s_id_issue.isoformat() if isinstance(s_id_issue, (date, datetime)) else "",
        "sender_phone": s_phone,

        "recipient_fullname": r_full,
        "recipient_account": r_acc,
        "recipient_addr": r_addr,
        "recipient_country_code": r_cc,
        "recipient_id_type": id_type_value(r_id_type, r_id_type_other),
        "recipient_id_no": r_id_no,

        "intermediary_bank": inter_bank,
        "intermediary_swift": inter_swift,
        "beneficiary_bank": ben_bank,
        "beneficiary_swift": ben_swift,

        "pay_type_personal": pay_type,
        "purpose_desc": purpose_desc,
        "docs_selected": ", ".join([f"{k} x{doc_counts.get(k,1)}" for k in (docs or [])]),

        "currency": currency,
        "foreign_amount": foreign_amt,
        "vnd_per_ngt": vnd_per_ngt,
        "vnd_per_usd": vnd_per_usd,
        "vnd_amount_rounded": int(round(vnd_amount, 0)) if not math.isnan(vnd_amount) else 0,
        "service_fee_vnd": int(round(fee, 0)) if not math.isnan(fee) else 0,
        "telex_fee_vnd": int(round(telex, 0)) if not math.isnan(telex) else 0,
        "total_vnd": int(round(total_vnd, 0)) if not math.isnan(total_vnd) else 0,

        "usd_current": usd_current if usd_current is not None else "",
        "cap_usd": cap_usd if cap_usd is not None else "",
        "cap_year_used": cap_year_used if cap_year_used is not None else "",
        "sent_sum_usd_year": sent_sum_usd if sent_sum_usd is not None else "",
        "remain_usd": remain_usd if remain_usd is not None else "",
        "warning": warning_text or "",
    }

def export_excel_with_template(template_file, row_dict: dict) -> bytes:
    """
    Nếu có file mẫu: giữ nguyên các sheet, thêm/ghi sheet 'Lenh_Chuyen_Tien' với dữ liệu dạng bảng.
    Nếu không có template: tạo file mới chỉ có sheet 'Lenh_Chuyen_Tien'.
    """
    df = pd.DataFrame([row_dict])

    if template_file is None:
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Lenh_Chuyen_Tien")
        out.seek(0)
        return out.read()

    wb = load_workbook(template_file)
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

row = compose_row_dict()
excel_bytes = export_excel_with_template(template, row)

st.download_button(
    label="⬇️ Tải file Excel (theo mẫu nếu có)",
    data=excel_bytes,
    file_name=f"lenh_chuyen_tien_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key=unique_key("download_btn")
)

st.success("Đã khởi tạo giao diện mới với key duy nhất cho mọi widget — lỗi DuplicateElementId đã được xử lý.")
