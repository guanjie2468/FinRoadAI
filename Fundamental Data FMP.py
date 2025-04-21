import requests
import pandas as pd
import yfinance as yf

# ------------------- SETUP -------------------
API_KEY = "wgghgc24g9lxkn8Ngle8EmE2A6zz0EDk"  # Replace with your FinancialModelingPrep API key
ticker = "GOOG"  # Example stock

# ------------------- FETCH FINANCIAL DATA -------------------
def fetch_fmp_data(statement_type, ticker):
    """Fetch financial statement data from FMP API."""
    url = f"https://financialmodelingprep.com/api/v3/{statement_type}/{ticker}?apikey={API_KEY}"
    response = requests.get(url)
    return response.json() if response.status_code == 200 else None

# Fetch statements
income_statement = pd.DataFrame(fetch_fmp_data("income-statement", ticker)).set_index("date").T
balance_sheet = pd.DataFrame(fetch_fmp_data("balance-sheet-statement", ticker)).set_index("date").T
cash_flow = pd.DataFrame(fetch_fmp_data("cash-flow-statement", ticker)).set_index("date").T

# Convert values to numeric
for df in [income_statement, balance_sheet, cash_flow]:
    df.replace('-', 0, inplace=True)
    df[:] = df.apply(pd.to_numeric, errors='coerce')

# ------------------- SAVE TO EXCEL -------------------
# Define the Excel file name
excel_filename = f"{ticker}_financial_data.xlsx"

# Save the financial statements to an Excel file
with pd.ExcelWriter(excel_filename) as writer:
    income_statement.to_excel(writer, sheet_name="Income Statement")
    balance_sheet.to_excel(writer, sheet_name="Balance Sheet")
    cash_flow.to_excel(writer, sheet_name="Cash Flow")

print(f"Financial data saved to '{excel_filename}'")
