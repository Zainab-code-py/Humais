import pandas as pd
from datetime import datetime, timedelta

# Load the data with specified dtypes to avoid mixed type warning
net_sales = pd.read_csv('NetSales_modified.csv', dtype={'Country Code': str, 'Brand': str}, low_memory=False)
bcodes = pd.read_csv('BCodes.csv')
calendar = pd.read_csv('Calendar.csv')


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

# Filter sales data for TW and LW
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

# Get unique brand codes and sort them
unique_brands = sorted(net_sales['Brand'].unique())

# Prepare a list to store results
results_list = []

# Calculate unique orders for each brand and store in the results list
for brand in unique_brands:
    tw_unique_orders = filtered_sales_tw[filtered_sales_tw['Brand'] == brand]['Order No.'].nunique()
    lw_unique_orders = filtered_sales_lw[filtered_sales_lw['Brand'] == brand]['Order No.'].nunique()
    ly_unique_orders = filtered_sales_ly[filtered_sales_ly['Brand'] == brand]['Order No.'].nunique()

    vs_lw = ((tw_unique_orders - lw_unique_orders) / lw_unique_orders * 100) if lw_unique_orders != 0 else 0
    vs_ly = ((tw_unique_orders - ly_unique_orders) / ly_unique_orders * 100) if ly_unique_orders != 0 else 0

    # Check if brand exists in bcodes
    brand_name_row = bcodes.loc[bcodes['Code'] == brand, 'Name']
    if not brand_name_row.empty:
        brand_name = brand_name_row.values[0]
    else:
        brand_name = brand  # Use the brand code if name is not found

    results_list.append({
        'Brand': brand_name,
        'TW': tw_unique_orders,
        'LW': lw_unique_orders,
        'vs LW': f"{vs_lw:.2f}%",
        'LY': ly_unique_orders,
        'vs LY': f"{vs_ly:.2f}%"
    })

# Convert the results list to a DataFrame
results = pd.DataFrame(results_list)

# Load existing WTD.xlsx file
with pd.ExcelWriter('WTD.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    results.to_excel(writer, sheet_name='Net Transactions', index=False)

print("Data written to WTD.xlsx successfully.")
