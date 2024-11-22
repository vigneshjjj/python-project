import requests
import pandas as pd
import xlwings as xw
from time import sleep
from datetime import datetime

# Function to fetch live cryptocurrency data from CoinGecko API
def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": False
    }
    response = requests.get(url, params=params)
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Error fetching data: {response.status_code}")
        return []

# Convert the fetched data into a pandas DataFrame
def get_crypto_dataframe(data):
    df = pd.DataFrame(data)
    return df[[
        "name", "symbol", "current_price", "market_cap", "total_volume", 
        "price_change_percentage_24h"
    ]]

# Perform analysis on the data
def analyze_data(df):
    # Top 5 by Market Cap
    top_5 = df.nlargest(5, "market_cap")[["name", "market_cap"]].values.tolist()
    
    # Average Price of Top 50
    average_price = df["current_price"].mean()
    
    # Highest and Lowest 24-hour Price Change
    max_change = df.loc[df["price_change_percentage_24h"].idxmax()]["name"]
    min_change = df.loc[df["price_change_percentage_24h"].idxmin()]["name"]
    
    return {
        "Top 5 Cryptos by Market Cap": top_5,
        "Average Price": average_price,
        "Highest 24h Change": max_change,
        "Lowest 24h Change": min_change
    }

# Update Excel file with data and analysis
def update_excel(df, analysis):
    # Connect to the active Excel workbook
    wb = xw.Book.caller()
    
    # Update "Crypto Data" sheet with current data
    sheet_data = wb.sheets["Crypto Data"]
    sheet_data.clear()  # Clear previous data
    sheet_data.range("A1").value = df.columns.tolist()  # Write headers
    sheet_data.range("A2").value = df.values  # Write data
    
    # Apply formatting to the data sheet
    sheet_data.range("A1:F1").color = (255, 255, 0)  # Yellow header color
    sheet_data.range("A1:F1").api.Font.Bold = True  # Bold headers
    sheet_data.range("A1:F100").api.Borders.LineStyle = 1  # Add borders
    sheet_data.range("A1:F100").api.Font.Name = "Calibri"
    sheet_data.range("A1:F100").api.Font.Size = 10
    sheet_data.range("A1:F100").api.HorizontalAlignment = 3  # Center align text
    
    # Conditional formatting for the 'price_change_percentage_24h' column
    change_column = sheet_data.range("F2:F51")  # Assuming 50 rows of data
    for cell in change_column:
        if cell.value < 0:
            cell.color = (255, 0, 0)  # Red for negative change
        elif cell.value > 0:
            cell.color = (0, 255, 0)  # Green for positive change
    
    # Update "Analysis Summary" sheet
    sheet_analysis = wb.sheets["Analysis Summary"]
    sheet_analysis.clear()  # Clear previous analysis
    
    # Table for Top 5 Cryptos by Market Cap
    sheet_analysis.range("A1").value = ["Top 5 Cryptos by Market Cap"]
    sheet_analysis.range("A2").value = [["Name", "Market Cap"]]  # Table header
    sheet_analysis.range("A3").value = analysis["Top 5 Cryptos by Market Cap"]
    sheet_analysis.range("A1:B1").color = (0, 255, 255)  # Cyan header color
    sheet_analysis.range("A1:B1").api.Font.Bold = True  # Bold headers
    sheet_analysis.range("A2:B6").api.Borders.LineStyle = 1  # Add borders
    
    # Table for Average Price of Top 50 Cryptos
    sheet_analysis.range("D1").value = ["Average Price of Top 50"]
    sheet_analysis.range("D2").value = [["Metric", "Value"]]  # Table header
    sheet_analysis.range("D3").value = [["Average Price", f"${analysis['Average Price']:.2f}"]]
    sheet_analysis.range("D1:E1").color = (255, 255, 0)  # Yellow header color
    sheet_analysis.range("D1:E1").api.Font.Bold = True  # Bold headers
    sheet_analysis.range("D2:E3").api.Borders.LineStyle = 1  # Add borders

    # Table for Highest and Lowest 24-Hour Change
    sheet_analysis.range("G1").value = ["Highest and Lowest 24-Hour Change"]
    sheet_analysis.range("G2").value = [["Metric", "Crypto Name"]]  # Table header
    sheet_analysis.range("G3").value = [["Highest 24h Change", analysis["Highest 24h Change"]]]
    sheet_analysis.range("G4").value = [["Lowest 24h Change", analysis["Lowest 24h Change"]]]
    sheet_analysis.range("G1:H1").color = (255, 255, 255)  # White header color
    sheet_analysis.range("G1:H1").api.Font.Bold = True  # Bold headers
    sheet_analysis.range("G2:H4").api.Borders.LineStyle = 1  # Add borders

    # Add Timestamp for the last update
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sheet_analysis.range("A6").value = "Data Last Updated"
    sheet_analysis.range("B6").value = timestamp

# Create the Excel file if it doesn't exist
def create_excel_file():
    try:
        # Try to open the existing workbook
        wb = xw.Book("LiveCryptoData.xlsx")
    except FileNotFoundError:
        # If the file doesn't exist, create it
        wb = xw.Book()

        # Create "Crypto Data" sheet and add headers
        sheet_data = wb.sheets[0]
        sheet_data.name = "Crypto Data"
        sheet_data.range("A1").value = ["Name", "Symbol", "Current Price", "Market Cap", "Total Volume", "24h Price Change"]

        # Create "Analysis Summary" sheet
        sheet_analysis = wb.sheets.add("Analysis Summary")
        sheet_analysis.range("A1").value = ["Metric", "Value"]

        # Save the file
        wb.save(r"C:\Users\Administrator\Desktop\LiveCryptoData.xlsx")
        print("LiveCryptoData.xlsx created and saved.")

    return wb

# Automate live updates every 3 minutes
def live_update():
    while True:
        crypto_data = fetch_crypto_data()
        if crypto_data:
            df_crypto = get_crypto_dataframe(crypto_data)
            analysis = analyze_data(df_crypto)
            update_excel(df_crypto, analysis)
            print("Data updated in Excel.")
        sleep(180)  # Update every 3 minutes

# Entry point for the script
if __name__ == "__main__":
    wb = create_excel_file()  # Create or open the workbook
    wb.set_mock_caller()  # Set the workbook as the caller for xlwings
    live_update()  # Start live update process
