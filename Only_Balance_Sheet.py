import argparse
import time
from datetime import date, timedelta

import pandas as pd
import requests

HEADERS = {"User-Agent": "Adam adamsaputra007@gmail.com"}

# Label -> XBRL tags to try in order (companies tag the same item differently)
METRICS = {
    "Current Assets": ["AssetsCurrent"],
    "Intangibles (excl. goodwill)": ["IntangibleAssetsNetExcludingGoodwill",
                                     "FiniteLivedIntangibleAssetsNet"],
    "Total Assets": ["Assets"],
    "Current Liabilities": ["LiabilitiesCurrent"],
    "Long-Term Debt": ["LongTermDebtNoncurrent", "LongTermDebt",
                       "LongTermDebtAndCapitalLeaseObligations"],
    "Total Liabilities": ["Liabilities"],
}


def latest_quarter():
    """Most recent calendar quarter-end that is at least 50 days old
    (so most companies have filed). Returns e.g. 'CY2026Q2I'."""
    cutoff = date.today() - timedelta(days=50)
    for year in (date.today().year, date.today().year - 1):
        for q, (m, d) in reversed(list(enumerate([(3, 31), (6, 30), (9, 30), (12, 31)], 1))):
            if date(year, m, d) <= cutoff:
                return f"CY{year}Q{q}I"


def get(url):
    time.sleep(0.2)
    r = requests.get(url, headers=HEADERS, timeout=60)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.json()


def frame(tag, period):
    data = get(f"https://data.sec.gov/api/xbrl/frames/us-gaap/{tag}/USD/{period}.json")
    if not data or not data.get("data"):
        return pd.DataFrame()
    return pd.DataFrame(data["data"]).drop_duplicates("cik").set_index("cik")


p = argparse.ArgumentParser()
p.add_argument("--period", default=latest_quarter(),
               help="e.g. CY2026Q2I (default: latest quarter)")
p.add_argument("--tickers-only", action="store_true",
               help="Drop companies that have no stock ticker")
args = p.parse_args()
print(f"Period: {args.period}")

# Pull each metric for ALL companies at once
columns = {}
info = None
for label, tags in METRICS.items():
    series = None
    for tag in tags:
        df = frame(tag, args.period)
        if df.empty:
            continue
        s = df["val"]
        series = s if series is None else series.combine_first(s)
        if tag == "Assets":
            info = df[["entityName", "end"]]
    columns[label] = series
    print(f"  {label}: {0 if series is None else len(series)} companies")

if info is None:
    raise SystemExit("No data returned. Try an older period, e.g. --period CY2026Q1I")

result = pd.DataFrame({k: v for k, v in columns.items() if v is not None})
result = info.join(result, how="left")
result = result.rename(columns={"entityName": "Company", "end": "Period End"})

# Add tickers (not every company has one)
tickers = get("https://www.sec.gov/files/company_tickers.json")
tmap = {}
for v in tickers.values():
    tmap.setdefault(int(v["cik_str"]), v["ticker"])
result.insert(0, "Ticker", result.index.map(tmap))
result.index.name = "CIK"
result = result.reset_index()

if args.tickers_only:
    result = result[result["Ticker"].notna()]

result = result.sort_values("Total Assets", ascending=False)

# Save to Excel
out = "all_companies_balance_sheet.xlsx"
with pd.ExcelWriter(out, engine="openpyxl") as writer:
    result.to_excel(writer, index=False, sheet_name="Balance Sheet")
    ws = writer.sheets["Balance Sheet"]
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col in ws.columns:
        width = max(len(str(c.value)) for c in col[:200] if c.value is not None)
        ws.column_dimensions[col[0].column_letter].width = min(width + 3, 40)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            if isinstance(c.value, (int, float)) and c.column >= 5:
                c.number_format = '#,##0;(#,##0);"-"'

print(f"\nSaved {len(result)} companies to {out}")