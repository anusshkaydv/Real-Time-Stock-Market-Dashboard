import requests
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime

# Your API Key (Replace with your actual key)
API_KEY = "IZ2VHZ33YPLLE6QF"

# List of 10 stocks to fetch
STOCKS = ["AAPL", "MSFT", "GOOGL", "AMZN", "TSLA", "META", "NFLX", "ASML", "NVDA", "AMD"]

# Fetch stock data from Alpha Vantage
print("Fetching stock data...")
data = []

for stock in STOCKS:
    url = f"https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol={stock}&apikey={API_KEY}"
    
    try:
        response = requests.get(url)
        json_data = response.json()
        
        # Check if API returned valid data
        if "Global Quote" in json_data:
            quote = json_data["Global Quote"]
            
            # Extract fields (handle missing data gracefully)
            symbol = quote.get("01. symbol", stock)
            price = quote.get("05. price", "N/A")
            change = quote.get("09. change", "N/A")
            change_percent = quote.get("10. change percent", "N/A").replace("%", "")
            volume = quote.get("06. volume", "N/A")
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            data.append({
                "Symbol": symbol,
                "Price": price,
                "Change": change,
                "ChangePercent": change_percent,
                "Volume": volume,
                "Timestamp": timestamp
            })
            
            print(f"✓ Fetched {stock}")
        else:
            print(f"✗ Failed to fetch {stock} - API returned no data")
    
    except Exception as e:
        print(f"✗ Error fetching {stock}: {str(e)}")

# Create DataFrame
df = pd.DataFrame(data)

# Write to Excel with formatting
excel_file = "Stock_Data.xlsx"
df.to_excel(excel_file, index=False, sheet_name="Data")

print(f"\n✓ Data saved to {excel_file}")
print(f"Total stocks fetched: {len(data)}/10")
print("\nData Preview:")
print(df)