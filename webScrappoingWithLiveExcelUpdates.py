# make sure excel application is active(open) before running the code.
import requests
import pandas as pd
import time
import xlwings as xw
import os

def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": "false"
    }
    response = requests.get(url, params=params)
    
    if response.status_code == 200:
        data = response.json()
        crypto_data = []
        for coin in data:
            crypto_data.append([
                coin["name"],
                coin["symbol"].upper(),
                coin["current_price"],
                coin["market_cap"],
                coin["total_volume"],
                coin["price_change_percentage_24h"]
            ])

        # Create a DataFrame
        df = pd.DataFrame(crypto_data, columns=["Name", "Symbol", "Price (USD)", "Market Cap", "24h Volume", "24h Change (%)"])
        return df
    else:
        print("Error fetching data:", response.status_code)
        return None

def excel_exists(filename="crypto_data.xlsx"):
    if not os.path.exists(filename):
        df = pd.DataFrame(columns=["Name", "Symbol", "Price (USD)", "Market Cap", "24h Volume", "24h Change (%)"])
        df.to_excel(filename, index=False)
        print(f"Created new file: {filename}")

def update_excel_live(df, filename="crypto_data.xlsx"):
    if df is not None:
        excel_exists(filename)
        try:
            # connect excel app
            app = xw.apps.active if xw.apps else xw.App(visible=True) 
            wb = None
            for book in app.books:
                if book.name == filename:
                    wb = book
                    break
            if not wb:
                wb = app.books.open(filename)
            
            sheet = wb.sheets[0]
            sheet.range("A1").value = df
            wb.save(filename)
            print("Excel updated successfully.")
        except Exception as e:
            print(f"Error interacting with Excel: {e}")


def main():
    excel_exists()
    while True:
        try:
            df = fetch_crypto_data()
            if df is not None:
                update_excel_live(df)
        except Exception as e:
            print(f"Error occurred: {e}")
        print("Waiting 2 minutes before next update...")
        time.sleep(5) 

if __name__ == "__main__":
    main()