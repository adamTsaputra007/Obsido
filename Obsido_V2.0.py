import argparse
import io
import os
import re
import time
import zipfile
from datetime import date, timedelta

import pandas as pd
import requests

HEADERS = {"User-Agent": "Adam adamsaputra007@gmail.com"}

OCF_TAGS = ["NetCashProvidedByUsedInOperatingActivities",
            "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations"]
CAPEX_TAGS = ["PaymentsToAcquirePropertyPlantAndEquipment",
              "PaymentsToAcquireProductiveAssets", "PaymentsForCapitalImprovements"]

# Label -> (kind, namespace, unit, [XBRL tags to try in order])
# kind "I" = balance sheet point in time, "D" = income statement over a period
METRICS = {
    # ---- Balance sheet ----
    "Current Assets (M)": ("I", "us-gaap", "USD", ["AssetsCurrent"]),
    "Intangibles excl. Goodwill (M)": ("I", "us-gaap", "USD",
        ["IntangibleAssetsNetExcludingGoodwill", "FiniteLivedIntangibleAssetsNet"]),
    "Total Assets (M)": ("I", "us-gaap", "USD", ["Assets"]),
    "Current Liabilities (M)": ("I", "us-gaap", "USD", ["LiabilitiesCurrent"]),
    "Long-Term Debt (M)": ("I", "us-gaap", "USD",
        ["LongTermDebtNoncurrent", "LongTermDebt",
         "LongTermDebtAndCapitalLeaseObligations"]),
    "Total Liabilities (M)": ("I", "us-gaap", "USD", ["Liabilities"]),
    "Retained Earnings (M)": ("I", "us-gaap", "USD", ["RetainedEarningsAccumulatedDeficit"]),
    # ---- Shareholders' equity section ----
    # Equity attributable to the parent; the fallback (used only when that's missing)
    # includes noncontrolling interest
    "Total Shareholders' Equity (M)": ("I", "us-gaap", "USD",
        ["StockholdersEquity",
         "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"]),
    "Preferred Stock (M)": ("I", "us-gaap", "USD",
        ["PreferredStockValue", "PreferredStockValueOutstanding",
         "PreferredStockIncludingAdditionalPaidInCapitalNetOfDiscount",
         "PreferredStockIncludingAdditionalPaidInCapital"]),
    "Noncontrolling Interest (M)": ("I", "us-gaap", "USD", ["MinorityInterest"]),
    "Treasury Stock (M)": ("I", "us-gaap", "USD",
        ["TreasuryStockValue", "TreasuryStockCommonValue"]),
    "Shares Issued (M)": ("I", "us-gaap", "shares", ["CommonStockSharesIssued"]),
    "Treasury Shares (M)": ("I", "us-gaap", "shares",
        ["TreasuryStockCommonShares", "TreasuryStockShares"]),
    # Fallback is the cover-page count, which is dated a few weeks after quarter end
    "Shares Outstanding (M)": ("I", "us-gaap", "shares",
        ["CommonStockSharesOutstanding", "dei:EntityCommonStockSharesOutstanding"]),
    # ---- Income statement ----
    "Sales (M)": ("D", "us-gaap", "USD",
        ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax",
         "RevenueFromContractWithCustomerIncludingAssessedTax", "SalesRevenueNet",
         "RevenuesNetOfInterestExpense"]),  # big banks
    "Litigation (M)": ("D", "us-gaap", "USD",
        ["LitigationSettlementExpense", "LossContingencyLossInPeriod",
         "LossContingencyAccrualProvisionNet"]),
    "R&D (M)": ("D", "us-gaap", "USD",
        ["ResearchAndDevelopmentExpense",
         "ResearchAndDevelopmentExpenseExcludingAcquiredInProcessCost"]),
    "Depreciation & Amortization (M)": ("D", "us-gaap", "USD",
        ["DepreciationDepletionAndAmortization", "DepreciationAndAmortization",
         "DepreciationAmortizationAndAccretionNet", "DepreciationAmortizationAndOther"]),
    "Operating Income (M)": ("D", "us-gaap", "USD", ["OperatingIncomeLoss"]),
    "Pre-Tax Income (M)": ("D", "us-gaap", "USD",
        ["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
         "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"]),
    "Income Tax (M)": ("D", "us-gaap", "USD", ["IncomeTaxExpenseBenefit"]),
    "Net Income (M)": ("D", "us-gaap", "USD", ["NetIncomeLoss", "ProfitLoss"]),
    "Preferred Dividends (M)": ("D", "us-gaap", "USD",
        ["DividendsPreferredStock", "DividendsPreferredStockCash",
         "PreferredStockDividendsIncomeStatementImpact",
         "PreferredStockDividendsAndOtherAdjustments",
         "PaymentsOfDividendsPreferredStockAndPreferenceStock"]),
    # Fallback (below) is Net Income - Preferred Dividends
    "Net Income to Common Stockholders (M)": ("D", "us-gaap", "USD",
        ["NetIncomeLossAvailableToCommonStockholdersBasic"]),
    # Only the income statement / EPS reconciliation tags (not the equity statement)
    "Preferred Dividends - EPS Reconciliation (M)": ("D", "us-gaap", "USD",
        ["PreferredStockDividendsIncomeStatementImpact",
         "PreferredStockDividendsAndOtherAdjustments"]),
    "Weighted Avg Shares Basic (M)": ("D", "us-gaap", "shares",
        ["WeightedAverageNumberOfSharesOutstandingBasic"]),
    # EPS is per share, so it stays in dollars rather than millions
    "EPS Basic ($)": ("D", "us-gaap", "USD-per-shares",
        ["EarningsPerShareBasic", "EarningsPerShareBasicAndDiluted"]),
    "Dividends (M)": ("D", "us-gaap", "USD",
        ["DividendsCommonStockCash", "Dividends", "PaymentsOfDividendsCommonStock",
         "PaymentsOfDividends"]),
    # ---- Cash flow ----
    # kind "P" = the quarter three quarters before the balance sheet quarter
    "Operating Cash Flow (M)": ("D", "us-gaap", "USD", OCF_TAGS),
    "CapEx (M)": ("D", "us-gaap", "USD", CAPEX_TAGS),
    "Operating Cash Flow 3Q Prior (M)": ("P", "us-gaap", "USD", OCF_TAGS),
    "CapEx 3Q Prior (M)": ("P", "us-gaap", "USD", CAPEX_TAGS),
}

# Any column whose label ends in "(M)" is shown in millions
IN_MILLIONS = {label for label in METRICS if label.endswith("(M)")}
IN_MILLIONS.add("Preferred + Treasury Stock (M)")

# SIC code ranges -> broad sector, first match wins (so specific ranges come first)
SECTOR_RANGES = [
    (6770, 6770, "Blank Check / SPAC"),
    (6798, 6798, "Real Estate"),  # REITs
    (1300, 1399, "Energy"), (2900, 2999, "Energy"), (4610, 4619, "Energy"),
    (2830, 2836, "Health Care"), (3841, 3851, "Health Care"),
    (8000, 8099, "Health Care"), (8731, 8731, "Health Care"),
    (3630, 3639, "Consumer Discretionary"), (3650, 3652, "Consumer Discretionary"),
    (3812, 3812, "Industrials"),
    (3570, 3579, "Technology"), (3600, 3699, "Technology"),
    (3820, 3829, "Technology"), (7370, 7379, "Technology"),
    (2710, 2799, "Communication Services"), (4800, 4899, "Communication Services"),
    (7800, 7899, "Communication Services"),
    (4950, 4959, "Industrials"),  # waste management
    (4900, 4999, "Utilities"),
    (6500, 6599, "Real Estate"),
    (6000, 6499, "Financials"), (6700, 6799, "Financials"),
    (100, 999, "Consumer Staples"), (2000, 2199, "Consumer Staples"),
    (2840, 2844, "Consumer Staples"), (5140, 5149, "Consumer Staples"),
    (5400, 5499, "Consumer Staples"), (5912, 5912, "Consumer Staples"),
    (5331, 5331, "Consumer Staples"), (5399, 5399, "Consumer Staples"),
    (1000, 1499, "Materials"), (2400, 2499, "Materials"), (2600, 2699, "Materials"),
    (2800, 2899, "Materials"), (3000, 3099, "Materials"), (3200, 3399, "Materials"),
    (1531, 1531, "Consumer Discretionary"),  # homebuilders
    (2200, 2399, "Consumer Discretionary"), (2500, 2599, "Consumer Discretionary"),
    (3100, 3199, "Consumer Discretionary"), (3710, 3716, "Consumer Discretionary"),
    (3751, 3751, "Consumer Discretionary"), (3940, 3949, "Consumer Discretionary"),
    (5200, 5999, "Consumer Discretionary"), (7000, 7099, "Consumer Discretionary"),
    (7200, 7299, "Consumer Discretionary"), (7500, 7599, "Consumer Discretionary"),
    (7900, 7999, "Consumer Discretionary"), (8200, 8299, "Consumer Discretionary"),
    (1500, 1799, "Industrials"), (3400, 3599, "Industrials"), (3700, 3799, "Industrials"),
    (3800, 3999, "Industrials"), (4000, 4799, "Industrials"), (5000, 5199, "Industrials"),
    (7300, 7399, "Industrials"), (8100, 8199, "Industrials"), (8700, 8799, "Industrials"),
    (9995, 9995, "Shell / Non-operating"),
]


def sector(sic):
    if pd.isna(sic):
        return None
    for lo, hi, name in SECTOR_RANGES:
        if lo <= sic <= hi:
            return name
    return "Other"


def sic_codes_for(name):
    """CIK -> SIC code from one SEC Financial Statement Data Set, e.g. '2026q3'
    (~65 MB zip). Cached to a small CSV so later runs skip the download."""
    cache = f"sic_codes_{name}.csv"
    if os.path.exists(cache):
        return pd.read_csv(cache, index_col="cik")["sic"]
    url = f"https://www.sec.gov/files/dera/data/financial-statement-data-sets/{name}.zip"
    print(f"Downloading SIC codes from {url} ...")
    r = requests.get(url, headers=HEADERS, timeout=300)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    with zipfile.ZipFile(io.BytesIO(r.content)).open("sub.txt") as f:
        sub = pd.read_csv(f, sep="\t", usecols=["cik", "sic"], encoding_errors="replace")
    sic = sub.dropna().drop_duplicates("cik", keep="last").set_index("cik")["sic"]
    try:
        sic.to_csv(cache)
    except OSError:
        pass  # caching is optional
    return sic


def load_sic_codes():
    """CIK -> SIC code from the newest published data set."""
    t = date.today()
    q = (t.month - 1) // 3 + 1
    for back in range(1, 5):  # a quarter's data set is published after it ends
        y, qq = divmod(t.year * 4 + (q - 1) - back, 4)
        sic = sic_codes_for(f"{y}q{qq + 1}")
        if sic is not None:
            return sic
    return pd.Series(dtype=float)


def latest_quarter():
    """Most recent calendar quarter-end at least 50 days old (so most
    companies have filed). Returns e.g. 'CY2026Q2I'."""
    cutoff = date.today() - timedelta(days=50)
    for year in (date.today().year, date.today().year - 1):
        for q, (m, d) in reversed(list(enumerate([(3, 31), (6, 30), (9, 30), (12, 31)], 1))):
            if date(year, m, d) <= cutoff:
                return f"CY{year}Q{q}I"


def get(url, tries=4):
    for attempt in range(tries):
        time.sleep(0.2 if attempt == 0 else 5 * attempt)  # back off on retries
        try:
            r = requests.get(url, headers=HEADERS, timeout=60)
        except (requests.Timeout, requests.ConnectionError):
            if attempt == tries - 1:
                raise
            continue
        if r.status_code == 404:
            return None
        if r.status_code in (429, 500, 502, 503, 504) and attempt < tries - 1:
            continue
        r.raise_for_status()
        return r.json()


def frame(ns, tag, unit, period):
    data = get(f"https://data.sec.gov/api/xbrl/frames/{ns}/{tag}/{unit}/{period}.json")
    if not data or not data.get("data"):
        return pd.DataFrame()
    return pd.DataFrame(data["data"]).drop_duplicates("cik").set_index("cik")


p = argparse.ArgumentParser()
p.add_argument("--period", default=latest_quarter(),
               help="Balance sheet date, e.g. CY2026Q2I (default: latest quarter)")
p.add_argument("--annual", action="store_true",
               help="Use full-year income statement figures instead of the quarter")
p.add_argument("--tickers-only", action="store_true",
               help="Drop companies that have no stock ticker")
p.add_argument("--out", default="all_companies_financials.xlsx",
               help="Output Excel file path")
args = p.parse_args()

m = re.fullmatch(r"CY(\d{4})Q([1-4])I", args.period)
if not m:
    raise SystemExit("Period must look like CY2026Q2I")
year, qtr = int(m.group(1)), int(m.group(2))

balance_period = args.period
if args.annual:
    # Full year ending at (or before) the chosen quarter
    income_period = f"CY{year if qtr == 4 else year - 1}"
else:
    income_period = f"CY{year}Q{qtr}"
    if qtr == 4:
        print("Note: companies don't report a standalone Q4 income statement in a "
              "10-Q, so Q4 income data will be sparse. Consider --annual.")
# The single quarter three quarters before the balance sheet quarter (Q2 2026 -> Q3 2025)
py, pq = divmod(year * 4 + (qtr - 1) - 3, 4)
prior_period = f"CY{py}Q{pq + 1}"
print(f"Balance sheet period: {balance_period} | Income statement period: {income_period} "
      f"| 3Q prior cash flow period: {prior_period}")

# Pull each metric for ALL companies at once
columns = {}
info = None
for label, (kind, ns, unit, tags) in METRICS.items():
    period = {"I": balance_period, "D": income_period, "P": prior_period}[kind]
    series = None
    for tag in tags:
        # A tag can override the namespace, e.g. "dei:EntityCommonStockSharesOutstanding"
        tag_ns, _, tag = tag.rpartition(":")
        df = frame(tag_ns or ns, tag, unit, period)
        if df.empty:
            continue
        s = df["val"]
        series = s if series is None else series.combine_first(s)
        if tag == "Assets":
            info = df[["entityName", "end"]]
    if label == "Sales (M)":
        # Other banks: revenue = net interest income + noninterest income
        nii = frame("us-gaap", "InterestIncomeExpenseNet", "USD", period)
        nonii = frame("us-gaap", "NoninterestIncome", "USD", period)
        if not nii.empty and not nonii.empty:
            bank = (nii["val"] + nonii["val"]).dropna()
            series = bank if series is None else series.combine_first(bank)
    if series is not None and label in IN_MILLIONS:
        series = series / 1_000_000
    columns[label] = series
    print(f"  {label}: {0 if series is None else len(series)} companies")

if info is None:
    raise SystemExit("No data returned. Try an older period, e.g. --period CY2026Q1I")

result = pd.DataFrame({k: v for k, v in columns.items() if v is not None})
result = info.join(result, how="left")
result = result.rename(columns={"entityName": "Company", "end": "Balance Sheet Date"})

# Most companies without preferred stock don't report income to common separately,
# so fill it in as Net Income - Preferred Dividends (none reported counts as 0)
nic = "Net Income to Common Stockholders (M)"
if {nic, "Net Income (M)", "Preferred Dividends (M)"} <= set(result.columns):
    derived = result["Net Income (M)"] - result["Preferred Dividends (M)"].fillna(0)
    result[nic] = result[nic].fillna(derived)

# Preferred + Treasury combined (a missing one counts as 0 if the other exists)
pt = [c for c in ("Preferred Stock (M)", "Treasury Stock (M)") if c in result.columns]
if pt:
    result["Preferred + Treasury Stock (M)"] = result[pt].sum(axis=1, min_count=1)

# Add tickers (not every company has one)
tickers = get("https://www.sec.gov/files/company_tickers.json")
tmap, nmap = {}, {}
for v in tickers.values():
    tmap.setdefault(int(v["cik_str"]), v["ticker"])
    nmap.setdefault(int(v["cik_str"]), v["title"])
result.insert(0, "Ticker", result.index.map(tmap))
# The ticker list has cleaner names than the XBRL data (e.g. "BofA Finance LLC" for BAC)
result["Company"] = pd.Series(result.index.map(nmap), index=result.index).fillna(result["Company"])

# Sector from the SEC's SIC industry code (last column)
sic = load_sic_codes()
result["Sector"] = result.index.map(sic).map(sector)
result.index.name = "CIK"
result = result.reset_index()

if args.tickers_only:
    result = result[result["Ticker"].notna()]

result = result.sort_values("Total Assets (M)", ascending=False)

# Save to Excel
out = args.out
with pd.ExcelWriter(out, engine="openpyxl") as writer:
    result.to_excel(writer, index=False, sheet_name="Financials")
    ws = writer.sheets["Financials"]
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    headers = [c.value for c in ws[1]]
    for col in ws.columns:
        width = max(len(str(c.value)) for c in col[:200] if c.value is not None)
        ws.column_dimensions[col[0].column_letter].width = min(width + 3, 40)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            if isinstance(c.value, (int, float)) and c.column >= 5:
                header = headers[c.column - 1]
                if header in IN_MILLIONS or header.endswith("($)"):
                    c.number_format = '#,##0.00;(#,##0.00);"-"'
                else:
                    c.number_format = '#,##0;(#,##0);"-"'
    note = ws.max_row + 2
    ws.cell(note, 1, f"Source: SEC EDGAR XBRL frames API. Balance sheet: {balance_period}. "
                     f"Income statement and cash flow: {income_period}. "
                     f"3Q prior cash flow: {prior_period}. All figures in millions "
                     "(USD, or shares for share counts) except EPS, which is dollars per share. Blank = not reported under the tags checked "
                     "(cash flow is mostly reported year-to-date, so single-quarter figures are "
                     "sparse except for Q1). Net Income to Common Stockholders is Net Income - Preferred "
                     "Dividends where not reported directly. Sector is grouped from the company's SEC SIC code.")

print(f"\nSaved {len(result)} companies to {out}")