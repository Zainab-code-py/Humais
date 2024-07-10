import pandas as pd
from datetime import datetime, timedelta

# Load the data with specified dtypes to avoid mixed type warning
net_sales = pd.read_csv('NetSales_modified.csv', dtype={'Country Code': str, 'Brand': str}, low_memory=False)
bcodes = pd.read_csv('BCodes.csv')
calendar = pd.read_csv('Calendar.csv')

# Load the Net Sales sheet from the WTD.xlsx file
net_sales_sheet = pd.read_excel('WTD.xlsx', sheet_name='Net Sales')


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

# Map brand codes to brand names
brand_code_to_name = dict(zip(bcodes['Code'], bcodes['Name']))

# Get unique brand codes and sort them
unique_brand_codes = sorted(net_sales['Brand'].unique())

# Prepare a list to store results for Net AOV
results_list_aov = []

# Calculate Net AOV for each brand and store in the results list
for brand_code in unique_brand_codes:
    # Map brand code to brand name
    brand_name = brand_code_to_name.get(brand_code, brand_code)

    # Check if the brand name exists in the Net Sales sheet
    if not net_sales_sheet[net_sales_sheet['Brand'] == brand_name].empty:
        tw_orders = filtered_sales_tw[filtered_sales_tw['Brand'] == brand_code]['Order No.'].count()
        lw_orders = filtered_sales_lw[filtered_sales_lw['Brand'] == brand_code]['Order No.'].count()
        ly_orders = filtered_sales_ly[filtered_sales_ly['Brand'] == brand_code]['Order No.'].count()

        print(
            f"Brand: {brand_name}, TW Orders: {tw_orders}, LW Orders: {lw_orders}, LY Orders: {ly_orders}")  # Debug print

        # Get the net amounts from the Net Sales sheet
        tw_net_amount = net_sales_sheet.loc[net_sales_sheet['Brand'] == brand_name, 'TW'].values[0]
        lw_net_amount = net_sales_sheet.loc[net_sales_sheet['Brand'] == brand_name, 'LW'].values[0]
        ly_net_amount = net_sales_sheet.loc[net_sales_sheet['Brand'] == brand_name, 'LY'].values[0]

        print(
            f"TW Net Amount: {tw_net_amount}, LW Net Amount: {lw_net_amount}, LY Net Amount: {ly_net_amount}")  # Debug print

        tw_aov = tw_net_amount / tw_orders if tw_orders != 0 else 0
        lw_aov = lw_net_amount / lw_orders if lw_orders != 0 else 0
        ly_aov = ly_net_amount / ly_orders if ly_orders != 0 else 0

        vs_lw = ((tw_aov - lw_aov) / lw_aov * 100) if lw_aov != 0 else 0
        vs_ly = ((tw_aov - ly_aov) / ly_aov * 100) if ly_aov != 0 else 0

        results_list_aov.append({
            'Brand': brand_name,
            'TW': tw_aov,
            'LW': lw_aov,
            'vs LW': f"{vs_lw:.2f}%",
            'LY': ly_aov,
            'vs LY': f"{vs_ly:.2f}%"
        })

# Convert the results list to a DataFrame
results_aov = pd.DataFrame(results_list_aov)

print(results_aov)  # Debug print

results_aov.to_csv('WTDNETAOV.csv', index=False)
# Load existing WTD.xlsx file and write the Net AOV sheet
with pd.ExcelWriter('WTD.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    results_aov.to_excel(writer, sheet_name='Net AOV', index=False)

print("Net AOV data written to WTD.xlsx successfully.")
