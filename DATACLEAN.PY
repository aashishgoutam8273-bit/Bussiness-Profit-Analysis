"""
BUSINESS PROFIT ANALYSIS - messy Excel ko clean karne ka code
-------------------------------------------------------------
VS Code me is file ko kholo aur upar right wala  Run (>)  button dabao.

Ye code:
  1. Excel file ko clean karta hai
  2. Clean CSV bana deta hai  ->  Business_Profit_Cleaned.csv
  3. Saari details terminal me dikhata hai
  4. Phir poochta hai: Business ka ID ya naam daalo -> uski poori detail dikhata hai
"""
import os
import re
import subprocess
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Library na ho to apne aap install kar leta hai
try:
    import pandas as pd
    import openpyxl  # noqa: F401
except ImportError:
    print("pandas / openpyxl install ho raha hai, thoda wait karo...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas", "openpyxl"])
    import pandas as pd

# ---- Excel file dhundo (script wale folder me, kisi bhi naam se jisme 'business' ho) ----
BASE = os.path.dirname(os.path.abspath(__file__))
candidates = [f for f in os.listdir(BASE)
              if f.lower().endswith((".xlsx", ".xlsm")) and "business" in f.lower()
              and "clean" not in f.lower() and not f.startswith("~$")]
if not candidates:
    sys.exit("Excel file nahi mili! Excel file ko is .py file ke SAATH, EK HI FOLDER me rakho.\n"
             "Folder: " + BASE)
INPUT_FILE = os.path.join(BASE, candidates[0])
OUTPUT_CSV = os.path.join(BASE, "Business_Profit_Cleaned.csv")
print("Input file :", candidates[0])

# ---------- 1. FILE PADHO (title rows skip karke asli header dhundo) ----------
raw = pd.read_excel(INPUT_FILE, header=None, dtype=object)
header_row = raw.index[raw.apply(
    lambda r: r.astype(str).str.strip().str.lower().eq("biz id").any(), axis=1)][0]
df = raw.iloc[header_row + 1:].copy()
df.columns = ["Biz ID", "Business Name", "Category", "City", "Start Date",
              "Revenue", "Expenses", "Net Profit", "Owner Contact", "Status"]
log = {"Raw rows (header ke baad)": len(df)}

# ---------- 2. JUNK ROWS HATAO ----------
df = df.dropna(how="all")                                           # khali rows
df = df[~df["Biz ID"].astype(str).str.strip().str.lower().eq("biz id")]   # repeated header
df = df[df["Biz ID"].astype(str).str.match(r"^\s*b[-\s]?\d+\s*$", case=False, na=False)]  # footer/total rows
log["Junk rows hatayi"] = log["Raw rows (header ke baad)"] - len(df)

# ---------- 3. TEXT COLUMNS ----------
df["Biz ID"] = df["Biz ID"].astype(str).str.extract(r"(\d+)")[0].astype(int).map("B{:03d}".format)

def clean_text(x):
    if pd.isna(x) or str(x).strip().lower() in ("", "n/a", "na", "?", "-", "--"):
        return pd.NA
    return re.sub(r"\s+", " ", str(x)).strip()

df["Business Name"] = df["Business Name"].map(clean_text).str.title()
df["Business Name"] = (df["Business Name"].str.replace(r"\bPvt Ltd\b", "Pvt Ltd", regex=True)
                                           .str.replace(r"\bCo\b\.?", "Co.", regex=True))

CATEGORY_MAP = {
    "retail": "Retail", "retial": "Retail",
    "food": "Food", "food & beverage": "Food", "f&b": "Food",
    "services": "Services", "service": "Services",
    "manufacturing": "Manufacturing", "mfg": "Manufacturing",
    "tech": "Tech", "it": "Tech", "technology": "Tech",
    "healthcare": "Healthcare", "health care": "Healthcare",
}
CITY_MAP = {
    "ghaziabad": "Ghaziabad", "gzb": "Ghaziabad", "ghazibad": "Ghaziabad",
    "delhi": "Delhi", "new delhi": "Delhi",
    "noida": "Noida", "meerut": "Meerut",
    "lucknow": "Lucknow", "lko": "Lucknow",
    "gurugram": "Gurugram", "gurgaon": "Gurugram",
}
df["Category"] = df["Category"].map(clean_text).str.lower().map(CATEGORY_MAP).fillna("Unknown")
df["City"]     = df["City"].map(clean_text).str.lower().map(CITY_MAP).fillna("Unknown")
df["Status"]   = df["Status"].map(clean_text).str.title().fillna("Unknown")

# ---------- 4. DATE ----------
def parse_date(x):
    if pd.isna(x): return pd.NaT
    if hasattr(x, "year"): return pd.Timestamp(x)
    s = str(x).strip()
    for fmt in ("%d/%m/%Y", "%d-%b-%y", "%Y.%m.%d", "%B %d, %Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try: return pd.to_datetime(s, format=fmt)
        except ValueError: continue
    return pd.NaT
df["Start Date"] = df["Start Date"].map(parse_date)

# ---------- 5. MONEY (₹, Rs., commas, lakh, (brackets)=negative) ----------
def parse_money(x):
    if pd.isna(x): return float("nan")
    if isinstance(x, (int, float)): return float(x)
    s = str(x).strip().lower()
    if s in ("", "n/a", "na", "?", "-", "--"): return float("nan")
    neg = s.startswith("(") or s.startswith("-")
    mult = 100000 if "lakh" in s else 1
    s = re.sub(r"rs\.?|inr|lakhs?", "", s)        # "Rs." ka dot hatao
    num = re.sub(r"[^\d.]", "", s)
    if not num: return float("nan")
    v = float(num) * mult
    v = round(v)
    return -v if neg else v

for col in ["Revenue", "Expenses", "Net Profit"]:
    df[col] = df[col].map(parse_money)

# Missing values ko baaki do columns se wapas nikalo (Profit = Revenue - Expenses)
m = df["Net Profit"].isna() & df["Revenue"].notna() & df["Expenses"].notna()
df.loc[m, "Net Profit"] = df["Revenue"] - df["Expenses"]
log["Net Profit recalculate kiya"] = int(m.sum())
m = df["Expenses"].isna() & df["Revenue"].notna() & df["Net Profit"].notna()
df.loc[m, "Expenses"] = df["Revenue"] - df["Net Profit"]
log["Expenses wapas nikala"] = int(m.sum())
m = df["Revenue"].isna() & df["Expenses"].notna() & df["Net Profit"].notna()
df.loc[m, "Revenue"] = df["Expenses"] + df["Net Profit"]
log["Revenue wapas nikala"] = int(m.sum())
df["Profit Margin %"] = (df["Net Profit"] / df["Revenue"]).round(4)
df["Result"] = df["Net Profit"].apply(lambda p: "Profit" if p > 0 else ("Loss" if p < 0 else "Break-even")
                                     if pd.notna(p) else "Unknown")

# ---------- 6. PHONE (sirf 10 digit) ----------
def clean_phone(x):
    d = re.sub(r"\D", "", str(x)) if not pd.isna(x) else ""
    d = d[-10:] if len(d) >= 10 else ""
    return d if d else pd.NA
df["Owner Contact"] = df["Owner Contact"].map(clean_phone)

# ---------- 7. DUPLICATES ----------
before = len(df)
df["_key"] = df["Business Name"].str.lower().str.replace(r"[^a-z0-9]", "", regex=True)
df = df.drop_duplicates(subset="_key", keep="first").drop(columns="_key")
log["Duplicate rows hatayi"] = before - len(df)

# ---------- 8. FINAL ----------
df = df.sort_values("Biz ID").reset_index(drop=True)
log["Final clean records"] = len(df)



# ---------- CSV SAVE ----------
df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")

# ================= REPORT (terminal) =================
pd.set_option("display.width", 250)
pd.set_option("display.max_columns", None)
pd.set_option("display.max_rows", None)
pd.set_option("display.float_format", lambda v: f"{v:,.0f}")

def pct(v): return f"{v*100:.1f}%" if pd.notna(v) else ""

print("\n" + "=" * 100)
print("CLEANING SUMMARY")
print("=" * 100)
for k, v in log.items(): print(f"  {k}: {v}")

print("\n" + "=" * 100)
print("BUSINESS PROFIT ANALYSIS - SUMMARY")
print("=" * 100)
print(f"Total businesses : {len(df)}")
print(f"Total Revenue    : Rs. {df['Revenue'].sum():,.0f}")
print(f"Total Expenses   : Rs. {df['Expenses'].sum():,.0f}")
print(f"Total Net Profit : Rs. {df['Net Profit'].sum():,.0f}")
print(f"Average Margin   : {pct(df['Profit Margin %'].mean())}")
print("\nProfit / Loss  :", df["Result"].value_counts().to_dict())
print("Status         :", df["Status"].value_counts().to_dict())

for col in ["Category", "City"]:
    g = (df.groupby(col).agg(Businesses=("Biz ID", "count"), Revenue=("Revenue", "sum"),
                             Expenses=("Expenses", "sum"), Net_Profit=("Net Profit", "sum"))
           .sort_values("Net_Profit", ascending=False))
    print(f"\n--- {col} ke hisaab se ---"); print(g.to_string())

show = ["Biz ID", "Business Name", "Net Profit", "Profit Margin %", "City", "Category"]
top = df.dropna(subset=["Net Profit"]).sort_values("Net Profit", ascending=False).copy()
top["Profit Margin %"] = top["Profit Margin %"].map(pct)
print("\n--- TOP 10 PROFIT WALE BUSINESS ---"); print(top.head(10)[show].to_string(index=False))
print("\n--- TOP 10 LOSS WALE BUSINESS ---");   print(top.tail(10).iloc[::-1][show].to_string(index=False))

view = df.copy()
view["Start Date"] = view["Start Date"].dt.strftime("%d-%b-%Y")
view["Profit Margin %"] = view["Profit Margin %"].map(pct)
print("\n" + "=" * 100); print("SAARE BUSINESS KI POORI DETAIL"); print("=" * 100)
print(view.to_string(index=False))
print(f"\nClean CSV ban gayi: {OUTPUT_CSV}")

# ================= INTERACTIVE SEARCH =================
def show_one(row):
    print("\n" + "-" * 50)
    for col in view.columns:
        val = row[col]
        if col in ("Revenue", "Expenses", "Net Profit") and pd.notna(val):
            val = f"Rs. {val:,.0f}"
        print(f"{col:<16}: {val if pd.notna(val) and val != '' else '-'}")
    print("-" * 50)

print("\n" + "=" * 100)
print("DETAIL SEARCH: Business ka ID (jaise B005 ya 5) ya naam ka koi bhi hissa likho.")
print("'all' = saari list dobara | 'q' = band")
print("=" * 100)
while True:
    try:
        q = input("\nDetail daalo > ").strip()
    except (EOFError, KeyboardInterrupt):
        break
    if not q: continue
    if q.lower() in ("q", "quit", "exit"): break
    if q.lower() == "all":
        print(view.to_string(index=False)); continue
    m = re.fullmatch(r"b?-?\s*(\d+)", q.lower())
    if m:
        hit = view[view["Biz ID"] == f"B{int(m.group(1)):03d}"]
    else:
        hit = view[view["Business Name"].str.lower().str.contains(re.escape(q.lower()), na=False)]
    if hit.empty:
        print("Koi business nahi mila. Dobara try karo.")
    else:
        print(f"{len(hit)} business mila:")
        for _, r in hit.iterrows(): show_one(r)
print("Bye!")