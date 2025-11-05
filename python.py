# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import io, csv, re, unicodedata
import requests
from datetime import date, datetime
from itertools import count
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# optional lists
try:
    import pycountry
except Exception:
    pycountry = None

# ============== CONFIG & TITLE ==============
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

# ============== HELPERS ==============
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

def fmt_vn_int(n): 
    try: return f"{int(round(float(n),0)):,}".replace(",", ".")
    except: return "0"

def fmt_usd(n):
    try: return f"{float(n):,.2f}"
    except: return "0.00"

def fmt_ddmmyyyy(d):
    if isinstance(d,(date,datetime)): return d.strftime("%d/%m/%Y")
    return ""

def clean_ccy(v)->str:
    if v is None: return ""
    s=str(v).strip().replace(NBSP," ")
    s=HTML_TAG_RE.sub(" ", s).upper()
    return s if re.fullmatch(r"[A-Z]{3}", s) else ""

def normalize_name(s:str)->list:
    if s is None: return []
    s=str(s).replace(NBSP," ")
    s=HTML_TAG_RE.sub(" ", s)
    s=unicodedata.normalize("NFKD", s)
    s="".join(ch for ch in s if not unicodedata.combining(ch))
    s=s.lower()
    s=re.sub(r"[^a-z0-9]+", " ", s)
    toks=[t for t in s.split() if t]
    stop={"co","ltd","company","the","and","account","acc","fees","fee","university",
          "bank","beneficiary","name","accountname","transfer","payment","inv"}
    return [t for t in toks if t not in stop]

def names_loose_match(a,b)->bool:
    A,B=set(normalize_name(a)), set(normalize_name(b))
    if not A or not B: return False
    if A.issubset(B) or B.issubset(A): return True
    inter=len(A&B); jacc=inter/max(1,len(A|B))
    return jacc>=0.7

def to_usd(amount, vnd_per_ccy, vnd_per_usd):
    if amount is None or pd.isna(amount): return 0.0
    if not (vnd_per_ccy and vnd_per_usd) or vnd_per_ccy<=0 or vnd_per_usd<=0: return 0.0
    return float(amount)*float(vnd_per_ccy)/float(vnd_per_usd)

def id_type_value(selected, other_text):
    if "Khác" in (selected or "") and (other_text or "").strip(): return other_text.strip()
    if "(Để trống)" in (selected or ""): return ""
    return selected or ""

def get_iso2_country_codes():
    items=[]
    if pycountry:
        try:
            for c in pycountry.countries:
                items.append((c.alpha_2.upper(), f"{c.alpha_2.upper()} – {c.name}"))
        except: pass
    if not items:
        fallback={"VN":"Viet Nam","US":"United States","AU":"Australia","JP":"Japan",
                  "KR":"Korea, Republic of","SG":"Singapore","CN":"China","DE":"Germany",
                  "FR":"France","GB":"United Kingdom","TH":"Thailand","CA":"Canada"}
        items=[(k,f"{k} – {v}") for k,v in fallback.items()]
    items.sort(key=lambda x:x[0]); return items

def get_iso4217_codes():
    codes=set()
    if pycountry:
        try:
            for c in pycountry.currencies:
                if getattr(c,"alpha_3",None): codes.add(c.alpha_3.upper())
        except: pass
    if not codes:
        codes={"USD","EUR","JPY","GBP","AUD","CAD","CHF","CNY","HKD","SGD","KRW",
               "THB","TWD","MYR","IDR","INR","VND","NZD","SEK","NOK","DKK","RUB",
               "AED","SAR","QAR","KWD","BHD","TRY","BRL","MXN","ZAR","PLN","HUF"}
    return sorted(list(codes))

def fetch_gdp_per_capita_usd(iso2, year):
    if not iso2 or not year: return None, None
    for y in [year, year-1, year-2]:
        try:
            u=f"https://api.worldbank.org/v2/country/{iso2.lower()}/indicator/NY.GDP.PCAP.CD?date={y}:{y}&format=json"
            js=requests.get(u,timeout=12).json()
            if isinstance(js,list) and len(js)>1 and js[1] and js[1][0]["value"] is not None:
                return float(js[1][0]["value"]), y
        except: pass
    return None, None

# ============== READ HISTORY (.xlsx / .xls / .csv / .html) ==============
def _flatten_header(df):
    if isinstance(df.columns, pd.MultiIndex):
        df.columns=[" ".join([str(c) for c in col if str(c)!="nan"]).strip() for col in df.columns]
    else:
        df.columns=[str(c) for c in df.columns]
    return df

def _row_is_header_like(row):
    txt=" ".join(map(str,row.values))
    txt=HTML_TAG_RE.sub(" ", txt).lower()
    keys=["message key","receiver","amount","người nhận","recipient","prepared date","ccy","currency","remark"]
    return sum(k in txt for k in keys) >= 3

def _find_col(df, exact, contains=()):
    cols={str(c).strip().lower():c for c in df.columns}
    for k in exact:
        if k in cols: return cols[k]
    for k in list(exact)+list(contains):
        for ck,oc in cols.items():
            if k in ck: return oc
    return None

def _infer_name_col(df):
    best,best_ratio=None,0
    for c in df.columns:
        ser=df[c].astype(str).head(400).apply(lambda x:" ".join(normalize_name(x)))
        def is_name(s):
            toks=[t for t in s.split() if t]
            return len(toks)>=2 and sum(t.isalpha() for t in toks)>=2
        ratio=ser.apply(is_name).mean()
        if ratio>best_ratio: best_ratio, best=c, ratio
    return best if best_ratio>=0.2 else None

def _infer_amount_col(df):
    best,best_ratio=None,0
    for c in df.columns:
        ser=df[c].astype(str).head(400).apply(parse_vn_number)
        ratio=ser.notna().mean()
        if ratio>best_ratio: best_ratio, best=ratio, c
    return best

def _infer_ccy_col(df):
    best,best_ratio=None,0
    for c in df.columns:
        vals=df[c].astype(str).head(400).apply(clean_ccy)
        ratio=vals.apply(lambda x:bool(re.fullmatch(r"[A-Z]{3}",x))).mean()
        if ratio>best_ratio: best_ratio,best=ratio,c
    return best if best_ratio>=0.3 else None

def _infer_date_col(df):
    best,best_ratio=None,0
    for c in df.columns:
        try:
            parsed=pd.to_datetime(df[c], errors="coerce", dayfirst=True)
            ratio=parsed.notna().mean()
            if ratio>best_ratio: best_ratio,best=ratio,c
        except: continue
    return best

def read_history(uploaded_file)->pd.DataFrame:
    empty=pd.DataFrame(columns=["recipient","amount","ccy","prepared date"])
    if uploaded_file is None: return empty

    # read once -> bytes to reuse
    raw = uploaded_file.read()
    name = getattr(uploaded_file, "name", "") or ""

    frames=[]

    # 1) xlsx by openpyxl
    try:
        if name.lower().endswith((".xlsx",".xlsm",".xltx",".xltm")):
            df=pd.read_excel(io.BytesIO(raw), engine="openpyxl")
            if isinstance(df,pd.DataFrame) and not df.empty: frames.append(df)
    except: pass

    # 2) xls by xlrd (BIFF8 legacy)
    try:
        if name.lower().endswith(".xls"):
            # IMPORTANT: xlrd==1.2.0 is required in requirements
            df=pd.read_excel(io.BytesIO(raw), engine="xlrd")
            if isinstance(df,pd.DataFrame) and not df.empty: frames.append(df)
    except: 
        pass

    # 3) CSV
    try:
        txt=raw.decode(errors="ignore")
        try:
            dialect=csv.Sniffer().sniff(txt[:4000])
            df=pd.read_csv(io.StringIO(txt), sep=dialect.delimiter)
        except Exception:
            df=None
            for sep in [",",";","|","\t"]:
                try:
                    df=pd.read_csv(io.StringIO(txt), sep=sep); break
                except Exception: pass
        if isinstance(df,pd.DataFrame) and not df.empty: frames.append(df)
    except: pass

    # 4) HTML table (xls export as HTML)
    try:
        html=raw.decode(errors="ignore")
        if "<table" in html.lower() or "<td" in html.lower():
            tables=pd.read_html(html, flavor="bs4")
            frames.extend([t for t in tables if isinstance(t,pd.DataFrame) and not t.empty])
    except: pass

    if not frames:
        st.error("Không đọc được file lịch sử (.xls/.xlsx/.csv/.html).")
        return empty

    # chọn frame đầu (ưu tiên đã đọc thành công theo thứ tự trên)
    df = frames[0].copy()
    df = _flatten_header(df)

    # loại dòng header lẫn trong data
    try: df = df[~df.apply(_row_is_header_like, axis=1)]
    except: pass

    # dò cột
    recip_exact=["recipient","người nhận","nguoi nhan","beneficiary","payee","receiver name","creditor name","account name","name"]
    recip_contains=["nguoi","nhan","beneficiar","payee","receiver","creditor","account","name"]
    amt_exact=["amount","số tiền","so tien","value","gia tri","amt"]
    ccy_exact=["ccy","currency","mã tiền","ma tien","cur","tiền tệ"]
    date_exact=["prepared date","value date","post date","posting date","transaction date","tx date","ngày","date"]

    rcol=_find_col(df,recip_exact,recip_contains) or _infer_name_col(df)
    acol=_find_col(df,amt_exact) or _infer_amount_col(df)
    ccol=_find_col(df,ccy_exact) or _infer_ccy_col(df)
    dcol=_find_col(df,date_exact) or _infer_date_col(df)

    out=pd.DataFrame(columns=["recipient","amount","ccy","prepared date"])
    if rcol is not None:
        out["recipient"]=df[rcol].astype(str).str.replace(NBSP," ",regex=False)\
            .apply(lambda s:HTML_TAG_RE.sub(" ",s)).str.strip()
    if acol is not None:
        out["amount"]=df[acol].apply(parse_vn_number).astype(float)
    if ccol is not None:
        out["ccy"]=df[ccol].apply(clean_ccy)
    else:
        out["ccy"]=""
    if dcol is not None:
        out["prepared date"]=pd.to_datetime(df[dcol], dayfirst=True, errors="coerce")
    else:
        out["prepared date"]=pd.NaT

    out=out[out["recipient"].astype(str).str.strip()!=""]
    out=out[out["amount"].fillna(0).astype(float)!=0]
    return out.reset_index(drop=True)

# ============== UI HELPERS ==============
_key_counter = count(1)
def unique_key(prefix:str)->str: return f"{prefix}_{next(_key_counter)}"

def inline_input(label_text, widget_fn, *args, key_prefix=None, **kwargs):
    left, right = st.columns([0.38, 0.62])
    with left: st.markdown(f"**{label_text}**")
    with right:
        kwargs.setdefault("label_visibility","collapsed")
        if "key" not in kwargs:
            base = key_prefix or label_text.replace(" ","_").lower()
            kwargs["key"]=unique_key(base)
        return widget_fn("", *args, **kwargs)

# ============== 1. NGƯỜI GỬI | 2. NGƯỜI NHẬN ==============
ISO_COUNTRIES = get_iso2_country_codes()
COUNTRY_LABELS = [x[1] for x in ISO_COUNTRIES]
CURRENCY_CODES = get_iso4217_codes()

left_col, right_col = st.columns(2)

with left_col:
    st.subheader("1. Người gửi")
    send_date = inline_input("Ngày gửi tiền", st.date_input, value=date.today(), format="DD/MM/YYYY", key_prefix="send_date")
    pay_method = inline_input("Hình thức thanh toán"
