import argparse
import re
import time
from datetime import date, timedelta

import pandas as pd
import requests

HEADERS = {"User-Agent": "Adam adamsaputra007@gmail.com"}

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
    "Total Equity (M)": ("I", "us-gaap", "USD",
        ["StockholdersEquity",
         "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"]),
    # Fallback is the cover-page count, which is dated a few weeks after quarter end
    "Shares Outstanding (M)": ("I", "us-gaap", "shares",
        ["CommonStockSharesOutstanding", "dei:EntityCommonStockSharesOutstanding"]),
    "Preferred Stock": ("I", "us-gaap", "USD", ["PreferredStockValue"]),
    "Treasury Stock": ("I", "us-gaap", "USD",
        ["TreasuryStockValue", "TreasuryStockCommonValue"]),
    # ---- Income statement ----
    "Sales (M)": ("D", "us-gaap", "USD",
        ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax",
         "RevenueFromContractWithCustomerIncludingAssessedTax", "SalesRevenueNet",
         "RevenuesNetOfInterestExpense"]),  # big banks
    "Litigation": ("D", "us-gaap", "USD",
        ["LitigationSettlementExpense", "LossContingencyLossInPeriod",
         "LossContingencyAccrualProvisionNet"]),
    "R&D": ("D", "us-gaap", "USD",
        ["ResearchAndDevelopmentExpense",
         "ResearchAndDevelopmentExpenseExcludingAcquiredInProcessCost"]),
    "Depreciation & Amortization": ("D", "us-gaap", "USD",
        ["DepreciationDepletionAndAmortization", "DepreciationAndAmortization",
         "DepreciationAmortizationAndAccretionNet", "DepreciationAmortizationAndOther"]),
    "Operating Income": ("D", "us-gaap", "USD", ["OperatingIncomeLoss"]),
    "Pre-Tax Income": ("D", "us-gaap", "USD",
        ["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
         "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"]),
    "Income Tax": ("D", "us-gaap", "USD", ["IncomeTaxExpenseBenefit"]),
    "Net Income": ("D", "us-gaap", "USD", ["NetIncomeLoss", "ProfitLoss"]),
    "Dividends": ("D", "us-gaap", "USD",
        ["DividendsCommonStockCash", "Dividends", "PaymentsOfDividendsCommonStock",
         "PaymentsOfDividends"]),
}

IN_MILLIONS = {"Current Assets (M)", "Intangibles excl. Goodwill (M)", "Total Assets (M)",
               "Current Liabilities (M)", "Long-Term Debt (M)", "Total Liabilities (M)",
               "Retained Earnings (M)", "Total Equity (M)", "Shares Outstanding (M)",
               "Sales (M)"}


def latest_quarter():
    """Most recent calendar quarter-end at least 50 days old (so most
    companies have filed). Returns e.g. 'CY2026Q2I'."""
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
print(f"Balance sheet period: {balance_period} | Income statement period: {income_period}")

# Pull each metric for ALL companies at once
columns = {}
info = None
for label, (kind, ns, unit, tags) in METRICS.items():
    period = balance_period if kind == "I" else income_period
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

# Preferred + Treasury combined (a missing one counts as 0 if the other exists)
pt = [c for c in ("Preferred Stock", "Treasury Stock") if c in result.columns]
if pt:
    result["Preferred + Treasury Stock"] = result[pt].sum(axis=1, min_count=1)

# Add tickers (not every company has one)
tickers = get("https://www.sec.gov/files/company_tickers.json")
tmap, nmap = {}, {}
for v in tickers.values():
    tmap.setdefault(int(v["cik_str"]), v["ticker"])
    nmap.setdefault(int(v["cik_str"]), v["title"])
result.insert(0, "Ticker", result.index.map(tmap))
# The ticker list has cleaner names than the XBRL data (e.g. "BofA Finance LLC" for BAC)
result["Company"] = pd.Series(result.index.map(nmap), index=result.index).fillna(result["Company"])
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
                if headers[c.column - 1] in IN_MILLIONS:
                    c.number_format = '#,##0.00;(#,##0.00);"-"'
                else:
                    c.number_format = '#,##0;(#,##0);"-"'
    note = ws.max_row + 2
    ws.cell(note, 1, f"Source: SEC EDGAR XBRL frames API. Balance sheet: {balance_period}. "
                     f"Income statement: {income_period}. Columns marked (M) are in millions; "
                     "all other values in whole USD. Blank = not reported under the tags checked.")

print(f"\nSaved {len(result)} companies to {out}")
