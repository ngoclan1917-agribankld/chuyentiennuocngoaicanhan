import streamlit as st
import pandas as pd
import io
import requests
from datetime import datetime, date
from unidecode import unidecode
import math

# Optional but recommended
try:
    import pycountry
except Exception:
    pycountry = None

# =========================
# ⚙️ CẤU HÌNH GIAO DIỆN
# =========================
st.set_page_config(page_title="💸 Lệnh chuyển tiền quốc tế", page_icon="💸", layout="wide")
st.title("💸 Trình tạo Lệnh chuyển tiền quốc tế (tỷ giá chéo USD & kiểm tra hạn mức Trợ cấp)")

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
    # Trả về [(code, "code – country name")]
    items = []
    if pycountry:
        try:
            for c in pycountry.countries:
                items.append((c.alpha_2.upper(), f"{c.alpha_2.upper()} – {c.name}"))
        except Exception:
            pass
    # fallback rút gọn nếu pycountry không có
    if not items:
        fallback = {
            "VN": "Viet Nam", "US": "United States", "AU": "Australia", "JP": "Japan",
            "KR": "Korea, Republic of", "SG": "Singapore", "CN": "China", "DE": "Germany",
            "FR": "France", "GB": "United Kingdom", "TH": "Thailand", "CA": "Canada"
        }
        items = [(k, f"{k} – {v}") for k, v in fallback.items()]
    items.sort(key=lambda x: x[0])
    return items

def country_name_from_code(code: str) -> str | None:
    code = (code or "").upper().strip()
    if not code or len(code) != 2:
        return None
    if pycountry:
        try:
            c = pycountry.countries.get(alpha_2=code)
            if c:
                return c.name
        except Exception:
            pass
    fallback = {
        "VN": "Viet Nam", "US": "United States", "AU": "Australia", "JP": "Japan",
        "KR": "Korea, Republic of", "SG": "Singapore", "CN": "China", "DE": "Germany",
        "FR": "France", "GB": "United Kingdom", "TH": "Thailand", "CA": "Canada"
    }
    return fallback.get(code)

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
        return pd.DataFrame(columns=["recipient", "amount", "prepared date"])
    ext = file.name.lower().split(".")[-1]
    if ext in ("xlsx", "xls"):
        df = pd.read_excel(file)
    else:
        df = pd.read_csv(file)

    # map cột linh hoạt
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
    # Ưu tiên 'prepared date', nếu không có thì bất kỳ 'date'
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

    # parse amount theo chuẩn VN nếu là chuỗi
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

# =========================
# 🧭 BỐ CỤC 2 CỘT
# =========================
colL, colR = st.columns([1.1, 1])

# =========================
# 📝 NHẬP THÔNG TIN
# =========================
with colL:
    st.subheader("1) Người gửi")
    send_date = st.date_input("Ngày gửi tiền", value=date.today())
    s_full = st.text_input("Họ tên người gửi")
    s_acc = st.text_input("Số tài khoản người gửi")
    s_addr = st.text_area("Địa chỉ người gửi")
    s_country = st.text_input("Quốc gia người gửi")
    s_id_type = st.selectbox("Loại giấy tờ người gửi", ["CCCD", "CC", "Passport", "Khác (tự nhập)"], index=0)
    s_id_type_other = st.text_input("Loại giấy tờ khác (người gửi)", disabled=(s_id_type != "Khác (tự nhập)"))
    s_id_no = st.text_input("Số giấy tờ người gửi")
    s_id_issue = st.date_input("Ngày cấp giấy tờ người gửi")
    s_phone = st.text_input("Số điện thoại người gửi")

    st.subheader("2) Người nhận")
    r_full = st.text_input("Họ tên người nhận")
    r_acc = st.text_input("Số tài khoản người nhận")
    r_addr = st.text_area("Địa chỉ người nhận")

    st.markdown("**Mã quốc gia người nhận (ISO-3166 alpha-2)**")
    cc_mode = st.radio("Cách nhập mã quốc gia", ["Chọn từ danh sách", "Nhập tay"], horizontal=True)
    iso_list = get_iso2_country_codes()
    if cc_mode == "Chọn từ danh sách":
        cc_choice = st.selectbox("Chọn mã quốc gia", options=[x[1] for x in iso_list], index=0)
        r_cc = cc_choice.split("–")[0].strip()
    else:
        r_cc = st.text_input("Nhập mã quốc gia 2 ký tự (ví dụ: VN, AU, US)").upper().strip()

    suggested_country_name = country_name_from_code(r_cc)
    if suggested_country_name:
        st.info(f"➡️ Mã quốc gia **{r_cc}** gợi ý: **{suggested_country_name}**")
    else:
        if r_cc:
            st.warning("⚠️ Mã quốc gia nhập tay không hợp lệ theo ISO-2. Vui lòng kiểm tra.")

    r_id_type = st.selectbox("Loại giấy tờ người nhận (tuỳ chọn)", ["(Để trống)", "CCCD", "CC", "Passport", "Khác (tự nhập)"], index=0)
    r_id_type_other = st.text_input("Loại giấy tờ khác (người nhận)", disabled=(r_id_type != "Khác (tự nhập)"))
    r_id_no = st.text_input("Số giấy tờ người nhận (tuỳ chọn)")

    st.subheader("3) Ngân hàng")
    inter_bank = st.text_input("Ngân hàng trung gian")
    inter_swift = st.text_input("SWIFT CODE ngân hàng trung gian")
    ben_bank = st.text_input("Ngân hàng nhận tiền")
    ben_swift = st.text_input("SWIFT CODE ngân hàng nhận tiền")

    st.subheader("4) Hồ sơ cung cấp")
    doc_opts = ["CCCD", "Giấy khai sinh", "Passport", "Visa", "Thông báo học phí", "Khác"]
    docs = st.multiselect("Chọn loại hồ sơ", options=doc_opts, default=[])
    doc_counts = {}
    for d in docs:
        doc_counts[d] = st.number_input(f"Số lượng '{d}'", min_value=1, value=1, step=1)

with colR:
    st.subheader("5) Mục đích & số tiền")
    pay_type = st.selectbox("Loại thanh toán (Cá nhân)", ["Trợ cấp", "Học phí", "Mục đích khác"], index=0)
    purpose_desc = st.text_area("Nội dung chuyển tiền")

    st.markdown("**Tiền tệ giao dịch & tỷ giá**")
    currency = st.text_input("Mã tiền tệ (ISO-4217, ví dụ: USD, EUR, JPY, VND)").upper().strip() or "USD"
    amt_str = st.text_input("Số tiền ngoại tệ (định dạng VN: 1.234.567,89)")
    vnd_per_ngt_str = st.text_input("Tỷ giá VND/NGT (VND cho 1 đơn vị nguyên tệ)", value="0")
    vnd_per_usd_str = st.text_input("Tỷ giá VND/USD (VND cho 1 USD)", value="0")
    fee_str = st.text_input("Phí dịch vụ (VND)", value="0")
    telex_str = st.text_input("Điện phí (VND)", value="0")

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

    c1, c2 = st.columns(2)
    with c1:
        st.metric("Số tiền quy đổi (VND, làm tròn 0)", fmt_vn_int(vnd_amount))
    with c2:
        st.metric("Tổng số tiền phải thu (VND)", fmt_vn_int(total_vnd))
    st.metric("Giá trị giao dịch hiện tại (USD, theo tỷ giá chéo)", fmt_usd(usd_current))

# =========================
# 📂 LỊCH SỬ GIAO DỊCH
# =========================
st.subheader("6) Lịch sử chuyển tiền (để cộng dồn USD theo năm gửi)")
hist_file = st.file_uploader("Tải file CSV/XLSX có cột: recipient, amount, prepared date (tuỳ chọn: currency)", type=["csv", "xlsx", "xls"])
hist_df = read_history(hist_file)

# Nếu file có nhiều loại tiền tệ, yêu cầu nhập tỷ giá VND/<mã> cho từng mã
rates_map = {}
if not hist_df.empty and hist_df["currency"].notna().any():
    st.info("Phát hiện file lịch sử có cột 'currency'. Vui lòng nhập tỷ giá VND/<mã> cho từng loại xuất hiện.")
    uniq_ccy = sorted([c for c in hist_df["currency"].dropna().unique().tolist() if c and c != "None"])
    cols = st.columns(min(3, len(uniq_ccy)) if uniq_ccy else 1)
    for idx, ccy in enumerate(uniq_ccy):
        with cols[idx % len(cols)]:
            val = st.text_input(f"VND/{ccy}", key=f"vnd_per_{ccy}")
            try:
                rates_map[ccy] = parse_vn_number(val) if val else 0.0
            except Exception:
                rates_map[ccy] = 0.0

# =========================
# 🧮 TRỢ CẤP: LẤY GDPpc & CỘNG DỒN USD
# =========================
cap_usd, cap_year_used, sent_sum_usd, remain_usd = None, None, None, None
warning_text = ""

if pay_type == "Trợ cấp" and r_full and r_cc and suggested_country_name and send_date:
    # GDP per capita
    cap_usd, cap_year_used = fetch_gdp_per_capita_usd(r_cc, send_date.year)
    with st.expander("Hạn mức trợ cấp (GDP/người)"):
        if cap_usd is not None:
            st.write(f"**GDP/người (USD)** của **{suggested_country_name}** cho **năm {cap_year_used}**: **{fmt_usd(cap_usd)} USD**")
        else:
            st.error("Không lấy được GDP/người từ World Bank cho mã quốc gia/năm này.")

    # Cộng dồn đã chuyển trong năm (USD)
    if not hist_df.empty:
        same_year = hist_df[hist_df["prepared date"].dt.year == send_date.year]
        mask = same_year["recipient"].astype(str).apply(lambda x: tokens_match(x, r_full))
        matched = same_year.loc[mask].copy()

        def row_to_usd(row):
            amt = row["amount"]
            row_ccy = row.get("currency", None)
            if pd.isna(row_ccy) or not row_ccy or row_ccy == "None":
                # giả định cùng nguyên tệ NGT với giao dịch hiện tại
                return to_usd(amt, vnd_per_ngt, vnd_per_usd)
            # có currency riêng -> cần VND/<row_ccy>
            v_row = rates_map.get(str(row_ccy).upper(), 0.0)
            return to_usd(amt, v_row, vnd_per_usd)

        matched["usd"] = matched.apply(row_to_usd, axis=1)
        sent_sum_usd = float(matched["usd"].sum())
    else:
        sent_sum_usd = 0.0

    with st.expander("Số tiền đã chuyển trong năm (sau quy đổi USD)"):
        st.write(f"**ĐÃ CHUYỂN NĂM {send_date.year}: {fmt_usd(sent_sum_usd)} USD**")

    if cap_usd is not None:
        remain_usd = cap_usd - sent_sum_usd
        st.write(f"**Số còn được chuyển (USD)** = {fmt_usd(remain_usd)}")
        if usd_current > remain_usd or (remain_usd is not None and remain_usd < 0):
            st.error("**🚨 CHUYỂN VƯỢT HẠN MỨC**")
            warning_text = "CHUYỂN VƯỢT HẠN MỨC"
        else:
            warning_text = ""

# =========================
# ⬇️ XUẤT EXCEL
# =========================
st.subheader("7) Tải về Excel dữ liệu lệnh")

def id_type_value(selected: str, other_text: str) -> str:
    if "Khác" in (selected or "") and (other_text or "").strip():
        return other_text.strip()
    if "(Để trống)" in (selected or ""):
        return ""
    return selected or ""

row = {
    "send_date": send_date.isoformat() if send_date else "",
    "sender_fullname": s_full, "sender_account": s_acc, "sender_addr": s_addr, "sender_country": s_country,
    "sender_id_type": id_type_value(s_id_type, s_id_type_other),
    "sender_id_no": s_id_no, "sender_id_issue_date": s_id_issue.isoformat() if s_id_issue else "", "sender_phone": s_phone,

    "recipient_fullname": r_full, "recipient_account": r_acc, "recipient_addr": r_addr,
    "recipient_country_code": r_cc, "recipient_country_suggested": suggested_country_name or "",
    "recipient_id_type": id_type_value(r_id_type, r_id_type_other), "recipient_id_no": r_id_no,

    "intermediary_bank": inter_bank, "intermediary_swift": inter_swift,
    "beneficiary_bank": ben_bank, "beneficiary_swift": ben_swift,

    "pay_type_personal": pay_type, "purpose_desc": purpose_desc,
    "docs_selected": ", ".join([f"{k} x{doc_counts.get(k,1)}" for k in docs]),

    "currency": currency, "foreign_amount": foreign_amt,
    "vnd_per_ngt": vnd_per_ngt, "vnd_per_usd": vnd_per_usd,
    "vnd_amount_rounded": int(round(vnd_amount, 0)) if not math.isnan(vnd_amount) else 0,
    "service_fee_vnd": fee, "telex_fee_vnd": telex,
    "total_vnd": int(round(total_vnd, 0)) if not math.isnan(total_vnd) else 0,

    "usd_current": usd_current if usd_current is not None else "",
    "cap_usd": cap_usd if cap_usd is not None else "",
    "cap_year_used": cap_year_used if cap_year_used is not None else "",
    "sent_sum_usd_year": sent_sum_usd if sent_sum_usd is not None else "",
    "remain_usd": remain_usd if remain_usd is not None else "",
    "warning": warning_text,
}

df_out = pd.DataFrame([row])

buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    df_out.to_excel(writer, index=False, sheet_name="remittance")
buf.seek(0)

st.download_button(
    label="⬇️ Tải Excel dữ liệu lệnh",
    data=buf,
    file_name=f"remittance_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.success("Hoàn tất. Lưu ý: để so sánh chuẩn theo USD, hãy nhập đúng **VND/NGT** và **VND/USD**; nếu file lịch sử có nhiều đồng tiền, cần nhập đầy đủ **VND/<mã tiền>** tương ứng.")
