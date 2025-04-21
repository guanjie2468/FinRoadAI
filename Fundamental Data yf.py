import yfinance as yf
import pandas as pd

# Define the stock ticker
ticker = 'JD'  # TO INPUT STOCK OF YOUR CHOICE!

# Create a Yahoo Finance object
stock = yf.Ticker(ticker)

# Extract financial statements
balance_sheet = stock.balance_sheet
cash_flow = stock.cashflow
income_statement = stock.financials

# Print the available labels in the financial statements
print("Income Statement Labels:", income_statement.index)
print("Balance Sheet Labels:", balance_sheet.index)
print("Cash Flow Labels:", cash_flow.index)

# Check if the required labels exist before accessing them
required_labels = {
    'income_statement': ['Total revenue', 'Net Income', 'Operating Income', 'Pretax Income', 'Tax Provision', 'Interest Expense', 'EBITDA', 'Diluted EPS'],
    'balance_sheet': ['Ordinary Shares Number', 'Total Debt', 'Total Assets', 'Cash And Cash Equivalents'],
    'cash_flow': ['Capital Expenditure', 'Depreciation Amortization Depletion', 'Change In Working Capital', 'Operating Cash Flow', 'Cash Dividends Paid']
}

for label in required_labels['income_statement']:
    if label not in income_statement.index:
        raise KeyError(f"Label '{label}' not found in income statement")

for label in required_labels['balance_sheet']:
    if label not in balance_sheet.index:
        raise KeyError(f"Label '{label}' not found in balance sheet")

for label in required_labels['cash_flow']:
    if label not in cash_flow.index:
        raise KeyError(f"Label '{label}' not found in cash flow statement")

# Calculate the required metrics
# Convert all data to billions
income_statement = income_statement / 1e9
balance_sheet = balance_sheet / 1e9
cash_flow = cash_flow / 1e9

# Financial Metrics and data for DCF Valuation
revenue = income_statement.loc['Total Revenue']
revenue_growth = revenue[::-1].pct_change()[::-1].dropna() * 100
net_income = income_statement.loc['Net Income']
net_income_growth = net_income[::-1].pct_change()[::-1].dropna() * 100
net_income_margin = (net_income / revenue) * 100
operating_income = income_statement.loc['Operating Income']
operating_margin = (operating_income / revenue) * 100
pretax_income = income_statement.loc['Pretax Income']
tax_provision = income_statement.loc['Tax Provision']
tax_rate = (tax_provision / pretax_income) * 100
capital_expenditure = cash_flow.loc['Capital Expenditure']
capital_expenditure_growth = capital_expenditure[::-1].pct_change()[::-1].dropna() * 100
depreciation_amortization = cash_flow.loc['Depreciation Amortization Depletion']
D_and_A_percent_of_capital_expenditure = (depreciation_amortization / capital_expenditure) * 100
change_in_working_capital = cash_flow.loc['Change In Working Capital']
change_in_working_capital_percent_of_revenue = (change_in_working_capital / revenue) * 100
free_cash_flow = income_statement.loc['Operating Income']*(1-tax_rate/100) + capital_expenditure - change_in_working_capital + depreciation_amortization

# Financial Metrics and data for WACC Calculation
ordinary_shares = balance_sheet.loc['Ordinary Shares Number']
Ordinary_shares_growth = ordinary_shares[::-1].pct_change()[::-1].dropna() * 100
latest_share_price = stock.history(period='1d')['Close'].iloc[-1]
market_cap = ordinary_shares * latest_share_price
interest_expense = income_statement.loc['Interest Expense']
total_debt = balance_sheet.loc['Total Debt']
cost_of_debt_pre_tax = (interest_expense / total_debt) * 100
cost_of_debt_after_tax = cost_of_debt_pre_tax * (1 - tax_rate / 100)

risk_free_rate_ticker = '^TNX'
risk_free_rate_data = yf.Ticker(risk_free_rate_ticker).history(period='1d')
risk_free_rate = risk_free_rate_data['Close'].iloc[-1] / 100  # Convert to a decimal

# Financial Metrics and data for WACC Calculation
equity_beta = stock.info['beta']
market_return = 0.08  # Assuming an 8% market return
cost_of_equity = (risk_free_rate + equity_beta * (market_return - risk_free_rate)) * 100

# Calculate WACC
total_equity = market_cap
total_capital = total_debt + total_equity
weight_of_debt = (total_debt / total_capital) * 100
weight_of_equity = (total_equity / total_capital) * 100
wacc = (weight_of_debt * cost_of_debt_after_tax / 100) + (weight_of_equity * cost_of_equity / 100)

# Financial Metrics and data for Financial Health Analysis
total_assets = balance_sheet.loc['Total Assets']
debt_ratio = (total_debt / total_assets) * 100
interest_expense = income_statement.loc['Interest Expense']
ebitda = income_statement.loc['EBITDA']
interest_coverage = (ebitda / interest_expense)

operating_cash_flow = cash_flow.loc['Operating Cash Flow']
operating_cash_flow_margin = (operating_cash_flow / revenue) * 100

cash_and_cash_equivalents = balance_sheet.loc['Cash And Cash Equivalents']
cash_and_cash_equivalents_growth = cash_and_cash_equivalents[::-1].pct_change()[::-1].dropna() * 100

diluted_eps = income_statement.loc['Diluted EPS'] * 1e9
diluted_eps_growth = diluted_eps[::-1].pct_change()[::-1].dropna() * 100

# Common Stock Dividend Paid Per Share
Dividend = cash_flow.loc['Cash Dividends Paid']
dividend_per_share = Dividend / ordinary_shares
dividend_per_share_growth = dividend_per_share[::-1].pct_change()[::-1].dropna() * 100

# Create a DataFrame to store the data
data = {
    'Revenue': revenue,
    'Revenue Growth (%)': revenue_growth,
    'Net Income': net_income,
    'Net Income Growth (%)': net_income_growth,
    'Net Income Margin (%)': net_income_margin,
    'Operating Income': operating_income,
    'Operating Margin (%)': operating_margin,
    'Pretax Income': pretax_income,
    'Tax Provision': tax_provision,
    'Tax Rate (%)': tax_rate,
    'Capital Expenditure': capital_expenditure,
    'Capital Expenditure Growth (%)': capital_expenditure_growth,
    'Depreciation & Amortization': depreciation_amortization,
    'D&A / CapEx (%)': D_and_A_percent_of_capital_expenditure,
    'Change in Working Capital': change_in_working_capital,
    'Change in Working Capital / Revenue (%)': change_in_working_capital_percent_of_revenue,
    'Free Cash Flow': free_cash_flow,
    'Cash and Cash Equivalents': cash_and_cash_equivalents,
    'Cash and Cash Equivalents Growth (%)': cash_and_cash_equivalents_growth,
    'Ordinary Shares': ordinary_shares,
    'Ordinary_shares_growth': Ordinary_shares_growth,    
    'Latest Share Price': latest_share_price,
    'Market Cap': market_cap,
    'Interest Expense': interest_expense,
    'Total Debt': total_debt,
    'Cost of Debt (Pre-Tax) (%)': cost_of_debt_pre_tax,
    'Cost of Debt (After-Tax) (%)': cost_of_debt_after_tax,
    'Risk-Free Rate (%)': risk_free_rate * 100,
    'Equity Beta': equity_beta,
    'Market Return (%)': market_return * 100,
    'Cost of Equity (%)': cost_of_equity,
    'WACC (%)': wacc,
    'Total Assets': total_assets,
    'Debt Ratio (%)': debt_ratio,
    'EBITDA': ebitda,
    'Interest Coverage': interest_coverage,
    'Operating Cash Flow': operating_cash_flow,
    'Operating Cash Flow Margin (%)': operating_cash_flow_margin,
    'Dividend Per Share': dividend_per_share,
    'Dividend Per Share Growth (%)': dividend_per_share_growth,
    'Diluted EPS': diluted_eps,
    'Diluted EPS Growth (%)': diluted_eps_growth
}

df = pd.DataFrame(data)

# Transpose the DataFrame to have years in rows and data in columns
df_transposed = df.T

# Save the transposed DataFrame to an Excel file with the ticker name in the filename
excel_filename = f'{ticker}_financial_analysis.xlsx'
df_transposed.to_excel(excel_filename, sheet_name=f'{ticker} Financial Analysis')

print(f"Financial analysis saved to '{excel_filename}'")