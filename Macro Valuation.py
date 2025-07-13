#Valuation

from ib_insync import *
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

ib = IB()
ib.connect('127.0.0.1', 7497, clientId=1)  # Change to 7496 for live trading

if ib.isConnected():
    print("✅ IBKR API connection successful!")
else:
    print("❌ Connection failed. Check API settings.")

# Define asset for AIAE proxy
spy = Stock('SPY', 'SMART', 'USD')  # SPY ETF (equity market inflow)

# Function to fetch historical data
def fetch_data(contract, duration='20 Y', bar_size='1 month', what_to_show='TRADES'):
    data = ib.reqHistoricalData(
        contract,
        endDateTime='',
        durationStr=duration,
        barSizeSetting=bar_size,
        whatToShow=what_to_show,
        useRTH=True,
        formatDate=1
    )
    return data

# Retrieve historical data
spy_data = fetch_data(spy)  # SPY Trading Volume

# Inspect the data structure
print("SPY Data Structure:")
print(spy_data)

# Convert to DataFrame
def to_dataframe(data, col_name):
    # Debugging: Print the raw data
    print(f"Raw data for {col_name}:")
    print(data)
    
    if not data:
        return pd.DataFrame()  # Return an empty DataFrame if data is empty
    
    df = pd.DataFrame(data)
    if 'date' in df.columns:
        df['date'] = pd.to_datetime(df['date'])
        df.set_index('date', inplace=True)
    elif 'time' in df.columns:
        df['time'] = pd.to_datetime(df['time'])
        df.set_index('time', inplace=True)
    df.rename(columns={'close': col_name}, inplace=True)
    return df

spy_df = to_dataframe(spy_data, 'SPY Inflow')

# Debugging: Print the DataFrame
print("SPY DataFrame:")
print(spy_df.head())

# ✅ Plot AIAE Proxy
plt.figure(figsize=(12, 6))
sns.lineplot(data=spy_df, x=spy_df.index, y='SPY Inflow', label='SPY Inflow (Equity Activity)', color='blue')

plt.title('Aggregate Investor Allocation to Equities (AIAE) - SPY Inflow Proxy')
plt.xlabel('Date')
plt.ylabel('Value')
plt.grid(True)
plt.legend()
plt.show()

# ✅ Disconnect IBKR API
ib.disconnect()