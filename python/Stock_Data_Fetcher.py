import requests
import pandas as pd
import time
from datetime import datetime

# ==============================
# Alpha Vantage API Key
# ==============================
API_KEY = "IZ2VHZ33YPLLE6QF"

# ==============================
# List of Stocks
# ==============================
STOCKS = [
    "AAPL",
    "MSFT",
    "GOOGL",
    "AMZN",
    "TSLA",
    "META",
    "NFLX",
    "ASML",
    "NVDA",
    "AMD"
]

print("Fetching stock data...\n")

data = []

# ==============================
# Fetch Data
# ==============================
for i, stock in enumerate(STOCKS, start=1):

    print(f"[{i}/10] Fetching {stock}...")

    url = (
        f"https://www.alphavantage.co/query?"
        f"function=GLOBAL_QUOTE&symbol={stock}&apikey={API_KEY}"
    )

    try:
        response = requests.get(url, timeout=30)
        json_data = response.json()

        if "Global Quote" in json_data and json_data["Global Quote"]:

            quote = json_data["Global Quote"]

            symbol = quote.get("01. symbol", stock)

            price = float(quote.get("05. price", 0) or 0)

            change = float(quote.get("09. change", 0) or 0)

            change_percent = float(
                quote.get("10. change percent", "0")
                .replace("%", "")
                or 0
            )

            volume = int(quote.get("06. volume", 0) or 0)

            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            data.append({
                "Symbol": symbol,
                "Price": price,
                "Change": change,
                "ChangePercent": change_percent,
                "Volume": volume,
                "Timestamp": timestamp
            })

            print(f"✓ {stock} fetched successfully")

        else:

            print(f"✗ {stock} - No data returned")

            data.append({
                "Symbol": stock,
                "Price": None,
                "Change": None,
                "ChangePercent": None,
                "Volume": None,
                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })

    except Exception as e:

        print(f"✗ Error fetching {stock}: {e}")

        data.append({
            "Symbol": stock,
            "Price": None,
            "Change": None,
            "ChangePercent": None,
            "Volume": None,
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })

    # Wait to avoid Alpha Vantage rate limits
    if i != len(STOCKS):
        print("Waiting 12 seconds...\n")
        time.sleep(12)

# ==============================
# Create DataFrame
# ==============================
df = pd.DataFrame(data)

# ==============================
# Save to Excel
# ==============================
excel_file = "Stock_Data.xlsx"

df.to_excel(
    excel_file,
    index=False,
    sheet_name="Data"
)

print("\n====================================")
print("Stock data saved successfully!")
print("====================================")
print(f"File: {excel_file}")
print(f"Rows Written: {len(df)}")

print("\nPreview:")
print(df)
