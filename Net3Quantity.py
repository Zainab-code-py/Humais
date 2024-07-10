import pandas as pd
from datetime import datetime, timedelta

# Load the data with specified dtypes to avoid mixed type warning
net_sales = pd.read_csv('NetSales_modified.csv', dtype={'Country Code': str, 'Brand': str}, low_memory=False)
bcodes = pd.read_csv('BCodes.csv')
calendar = pd.read_csv('Calendar.csv')

# Load the WTD.xlsx file
wtd_file = 'WTD.xlsx'


# Helper function to get the closest Sunday
def get_closest_sunday(date):
    return (date - timedelta(days=(date.weekday() + 1) % 7)).replace(hour=0, minute=0, second=0, microsecond=0)


# Get the current date and determine the TW and LW date ranges
today = datetime.today()
closest_sunday = get_closest_sunday(today)

# Calculate TW and LW ranges
tw_start = closest_sunday - timedelta(days=7)
tw_end = closest_sunday - timedelta(days=1)
lw_start = closest_sunday - timedelta(days=14)
lw_end = closest_sunday - timedelta(days=8)

# Debug print to check the date ranges
print(f"Closest Sunday: {closest_sunday}")
print(f"TW Range: {tw_start} to {tw_end}")
print(f"LW Range: {lw_start} to {lw_end}")


# Convert Finance_Date to datetime and handle different formats
def parse_date(date_str):
    try:
        return datetime.strptime(date_str, '%d/%m/%Y %H:%M')
    except ValueError:
        try:
            return datetime.strptime(date_str, '%m/%d/%Y %I:%M:%S %p')
        except ValueError:
            return pd.NaT


net_sales['Finance_Date'] = net_sales['Finance_Date'].apply(parse_date)
net_sales = net_sales.dropna(subset=['Finance_Date'])

# Ensure Quantity column is numeric
net_sales['Quantity'] = pd.to_numeric(net_sales['Quantity'], errors='coerce')

# Filter sales data for TW, LW, and LY
filtered_sales_tw = net_sales[
    (net_sales['Finance_Date'] >= tw_start) &
    (net_sales['Finance_Date'] <= tw_end)
].copy()

filtered_sales_lw = net_sales[
    (net_sales['Finance_Date'] >= lw_start) &
    (net_sales['Finance_Date'] <= lw_end)
].copy()

# Calculate LY sales
random_date_tw = filtered_sales_tw['Finance_Date'].iloc[0]  # Using the first date from TW range
random_date_sort = int(random_date_tw.strftime('%Y%m%d'))

# Step 1: Find the Trading Week for the random_date_tw
tw_week = calendar.loc[calendar['Sort'] == random_date_sort, 'Trading Week'].values[0]

# Step 2: Find the same Trading Week in the previous year (2023)
previous_year = random_date_tw.year - 1
ly_week = calendar.loc[(calendar['Trading Week'] == tw_week) & (calendar['CalYear'] == previous_year)]

# Step 3: Determine all dates for that trading week
ly_dates = ly_week['Sort'].values
ly_dates = [datetime.strptime(str(date), '%Y%m%d') for date in ly_dates]

# Filter sales data for LY dates
filtered_sales_ly = net_sales[
    (net_sales['Finance_Date'].isin(ly_dates))
].copy()

# Group by Brand and calculate the net quantity for TW, LW, and LY
net_quantity_tw = filtered_sales_tw.groupby('Brand')['Quantity'].sum().reset_index(name='TW')
net_quantity_lw = filtered_sales_lw.groupby('Brand')['Quantity'].sum().reset_index(name='LW')
net_quantity_ly = filtered_sales_ly.groupby('Brand')['Quantity'].sum().reset_index(name='LY')

# Merge the results into a single DataFrame
net_quantity = pd.merge(net_quantity_tw, net_quantity_lw, on='Brand', how='outer')
net_quantity = pd.merge(net_quantity, net_quantity_ly, on='Brand', how='outer')

# Fill NaN values with 0
net_quantity = net_quantity.fillna(0)

# Calculate vs LW and vs LY as percentage differences
net_quantity['vs LW'] = ((net_quantity['TW'] - net_quantity['LW']) / net_quantity['LW']) * 100
net_quantity['vs LY'] = ((net_quantity['TW'] - net_quantity['LY']) / net_quantity['LY']) * 100

# Round percentages to 2 decimal places and convert to string with percent sign
net_quantity['vs LW'] = net_quantity['vs LW'].round(2).astype(str) + '%'
net_quantity['vs LY'] = net_quantity['vs LY'].round(2).astype(str) + '%'

# Replace brand codes with brand names using BCodes.csv
net_quantity = pd.merge(net_quantity, bcodes, left_on='Brand', right_on='Code', how='left')
net_quantity = net_quantity.drop(columns=['Brand', 'Code'])
net_quantity = net_quantity.rename(columns={'Name': 'Brand'})

# Reorder columns
net_quantity = net_quantity[['Brand', 'TW', 'LW', 'vs LW', 'LY', 'vs LY']]

# Save the results to the WTD.xlsx file
with pd.ExcelWriter(wtd_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    net_quantity.to_excel(writer, sheet_name='Net Quantity', index=False)

print("Net Quantity sheet has been successfully created.")
