import requests
import pandas as pd

headers = {"User-Agent": "Adam adamsaputra007@gmail.com"}
url = "https://www.sec.gov/files/company_tickers.json"

data = requests.get(url, headers=headers, timeout=30).json()
df = pd.DataFrame({"Ticker": sorted(v["ticker"] for v in data.values())})

df.to_excel("tickers.xlsx", index=False)
print(f"Saved {len(df)} tickers to tickers.xlsx")