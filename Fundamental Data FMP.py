# This script fetches financial data from FinancialModelingPrep (FMP) API, processes it, and saves it to an Excel file.
import requests
import pandas as pd

# ------------------- SETUP -------------------
API_KEY = "wgghgc24g9lxkn8Ngle8EmE2A6zz0EDk"  # Replace with your FinancialModelingPrep API key
ticker = "GOOG"  # Set your ticker

# ------------------- FETCH DATA -------------------
def fetch_fmp_data(statement_type, ticker, period="annual"):
    url = f"https://financialmodelingprep.com/api/v3/{statement_type}/{ticker}?period={period}&limit=5&apikey={API_KEY}"
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        if isinstance(data, list) and data and 'date' in data[0]:
            return pd.DataFrame(data).set_index("date").T
    return pd.DataFrame()

def to_numeric_billions(df):
    df.replace("-", 0, inplace=True)
    return df.apply(pd.to_numeric, errors='coerce') / 1e9

# Fetch data and convert to billions
income_df = to_numeric_billions(fetch_fmp_data("income-statement", ticker))
balance_df = to_numeric_billions(fetch_fmp_data("balance-sheet-statement", ticker))
cashflow_df = to_numeric_billions(fetch_fmp_data("cash-flow-statement", ticker))

# ------------------- SAFE EXTRACTION -------------------
def safe_get(df, label):
    return df.loc[label] if label in df.index else pd.Series([None]*df.shape[1], index=df.columns)

# ------------------- ANALYSIS VARIABLES -------------------
revenue = safe_get(income_df, "revenue")
net_income = safe_get(income_df, "netIncome")
operating_income = safe_get(income_df, "operatingIncome")
pretax_income = safe_get(income_df, "incomeBeforeTax")
tax_provision = safe_get(income_df, "incomeTaxExpense")
interest_expense = safe_get(income_df, "interestExpense")
ebitda = safe_get(income_df, "ebitda")
diluted_eps = safe_get(income_df, "epsdiluted")

capex = safe_get(cashflow_df, "capitalExpenditure")
d_and_a = safe_get(cashflow_df, "depreciationAndAmortization")
change_wc = safe_get(cashflow_df, "changeInWorkingCapital")
operating_cf = safe_get(cashflow_df, "netCashProvidedByOperatingActivities")
dividends_paid = safe_get(cashflow_df, "dividendsPaid")

cash = safe_get(balance_df, "cashAndCashEquivalents")
total_debt = safe_get(balance_df, "totalDebt")
total_assets = safe_get(balance_df, "totalAssets")
shares_out = safe_get(balance_df, "commonStockSharesOutstanding")

# ------------------- CALCULATIONS -------------------
rev_growth = revenue[::-1].pct_change()[::-1] * 100
ni_growth = net_income[::-1].pct_change()[::-1] * 100
capex_growth = capex[::-1].pct_change()[::-1] * 100
cash_growth = cash[::-1].pct_change()[::-1] * 100
net_margin = (net_income / revenue) * 100
op_margin = (operating_income / revenue) * 100
tax_rate = (tax_provision / pretax_income) * 100
debt_ratio = (total_debt / total_assets) * 100
interest_coverage = ebitda / interest_expense
ocf_margin = (operating_cf / revenue) * 100
free_cf = operating_cf + capex + d_and_a - change_wc
dividend_per_share = dividends_paid / shares_out

# ------------------- RATIO SHEET -------------------
# Reverse the columns so oldest year is left, latest is right
ratios = pd.DataFrame({
    "Revenue (B)": revenue,
    "Revenue Growth (%)": rev_growth,
    "Net Income (B)": net_income,
    "Net Income Growth (%)": ni_growth,
    "Net Margin (%)": net_margin,
    "Operating Margin (%)": op_margin,
    "Pretax Income (B)": pretax_income,
    "Tax Provision (B)": tax_provision,
    "Tax Rate (%)": tax_rate,
    "EBITDA (B)": ebitda,
    "Interest Expense (B)": interest_expense,
    "Interest Coverage (x)": interest_coverage,
    "Total Debt (B)": total_debt,
    "Total Assets (B)": total_assets,
    "Debt Ratio (%)": debt_ratio,
    "Operating Cash Flow (B)": operating_cf,
    "OCF Margin (%)": ocf_margin,
    "CapEx (B)": capex,
    "CapEx Growth (%)": capex_growth,
    "D&A (B)": d_and_a,
    "D&A as % of CapEx": (d_and_a / capex) * 100,
    "Change in Working Capital (B)": change_wc,
    "Change in WC as % of Revenue": (change_wc / revenue) * 100,
    "Free Cash Flow (B)": free_cf,
    "Cash (B)": cash,
    "Cash Growth (%)": cash_growth,
    "Diluted EPS": diluted_eps,
    "Shares Outstanding (B)": shares_out / 1e9,
    "Dividend Per Share ($)": dividend_per_share
}).T

# Reverse the columns so oldest year is left, latest is right
ratios = ratios[ratios.columns[::-1]]

# ------------------- SAVE TO EXCEL -------------------
excel_filename = f"{ticker}_FMP_financial_analysis.xlsx"
with pd.ExcelWriter(excel_filename) as writer:
    income_df.to_excel(writer, sheet_name="Income Statement")
    balance_df.to_excel(writer, sheet_name="Balance Sheet")
    cashflow_df.to_excel(writer, sheet_name="Cash Flow")
    ratios.to_excel(writer, sheet_name="Ratio Analysis")

print(f"Financial analysis saved to '{excel_filename}'")
