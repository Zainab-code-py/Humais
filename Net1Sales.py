import pandas as pd
from datetime import datetime, timedelta

# Load the data with specified dtypes to avoid mixed type warning
gross_sales = pd.read_csv('NetSales_modified.csv', dtype={'Country Code': str, 'Brand': str}, low_memory=False)
currencies = pd.read_excel('Currencies.xlsx')
bcodes = pd.read_csv('BCodes.csv')
calendar = pd.read_csv('Calendar.csv', dtype={'Sort': int, 'Trading Week': str, 'CalYear': int})

# Load the budget data
budget_data = pd.read_excel('Budget_Summary_TW.xlsx', sheet_name='Budget Summary')

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

gross_sales['Finance_Date'] = gross_sales['Finance_Date'].apply(parse_date)
gross_sales = gross_sales.dropna(subset=['Finance_Date'])

# Filter sales data for TW and LW
filtered_sales_tw = gross_sales[
    (gross_sales['Finance_Date'] >= tw_start) &
    (gross_sales['Finance_Date'] <= tw_end)
    ].copy()
filtered_sales_lw = gross_sales[
    (gross_sales['Finance_Date'] >= lw_start) &
    (gross_sales['Finance_Date'] <= lw_end)
    ].copy()

# Debug prints to check if the filtering is correct
print(f"Filtered Sales Data TW (Head):\n{filtered_sales_tw.head()}")
print(f"Filtered Sales Data TW Dates:\n{filtered_sales_tw['Finance_Date'].unique()}")
print(f"Filtered Sales Data LW (Head):\n{filtered_sales_lw.head()}")
print(f"Filtered Sales Data LW Dates:\n{filtered_sales_lw['Finance_Date'].unique()}")

# Convert Amount and currency rates to appropriate data types
filtered_sales_tw.loc[:, 'Amount Inc VAT'] = pd.to_numeric(filtered_sales_tw['Amount Inc VAT'], errors='coerce')
filtered_sales_lw.loc[:, 'Amount Inc VAT'] = pd.to_numeric(filtered_sales_lw['Amount Inc VAT'], errors='coerce')
currencies['Jun-2024'] = pd.to_numeric(currencies['Jun-2024'], errors='coerce')
currencies['Jun-2023'] = pd.to_numeric(currencies['Jun-2023'], errors='coerce')  # Separate conversion rate for LY

# Prepare the result DataFrame
result = pd.DataFrame(columns=['Brand', 'TW', 'LW', 'vs LW', 'LY', 'vs LY'])

# Get unique brands from TW, LW, and LY sales data
brands_tw = filtered_sales_tw['Brand'].unique()
brands_lw = filtered_sales_lw['Brand'].unique()

# Find the Trading Week for LY
random_date_tw = filtered_sales_tw['Finance_Date'].iloc[0]  # Using the first date from TW range
random_date_sort = int(random_date_tw.strftime('%Y%m%d'))

tw_week = calendar.loc[calendar['Sort'] == random_date_sort, 'Trading Week'].values[0]
ly_week = calendar.loc[(calendar['Trading Week'] == tw_week) & (calendar['CalYear'] == 2023)]
ly_dates = [datetime.strptime(str(date), '%Y%m%d') for date in ly_week['Sort'].values]

filtered_sales_ly = gross_sales[
    (gross_sales['Finance_Date'].isin(ly_dates))
    ].copy()

brands_ly = filtered_sales_ly['Brand'].unique()

# Combine all unique brands from TW, LW, and LY
all_brands = set(brands_tw).union(set(brands_lw)).union(set(brands_ly))

# Dynamic currency conversion function
def convert_to_usd(sales, country_code, currencies):
    total_amount_usd = 0
    sales_by_month = sales.groupby(sales['Finance_Date'].dt.to_period('M'))

    for period, sales in sales_by_month:
        total_revenue = sales['Amount Inc VAT'].sum()
        period_str = period.strftime('%b-%Y')
        conversion_rate = currencies.loc[currencies['Country_Code'] == country_code, period_str].values

        if len(conversion_rate) > 0 and conversion_rate[0] > 0:
            amount_usd = total_revenue / conversion_rate[0]
        else:
            amount_usd = 0

        if country_code == 'AE':
            amount_usd -= amount_usd * 0.05
        elif country_code == 'LB':
            amount_usd -= amount_usd * 0.11

        total_amount_usd += amount_usd

    return total_amount_usd

# Calculate the total amount for each brand for TW, LW, and LY
for brand in all_brands:
    total_amount_usd_tw = 0
    total_amount_usd_lw = 0
    total_amount_usd_ly = 0

    # Calculate TW sales
    if brand in brands_tw:
        brand_sales_tw = filtered_sales_tw[filtered_sales_tw['Brand'] == brand]
        country_codes_tw = brand_sales_tw['Country Code'].unique()

        for country_code in country_codes_tw:
            country_sales_tw = brand_sales_tw[brand_sales_tw['Country Code'] == country_code]
            total_amount_usd_tw += convert_to_usd(country_sales_tw, country_code, currencies)

    # Calculate LW sales
    if brand in brands_lw:
        brand_sales_lw = filtered_sales_lw[filtered_sales_lw['Brand'] == brand]
        country_codes_lw = brand_sales_lw['Country Code'].unique()

        for country_code in country_codes_lw:
            country_sales_lw = brand_sales_lw[brand_sales_lw['Country Code'] == country_code]
            total_amount_usd_lw += convert_to_usd(country_sales_lw, country_code, currencies)

    # Calculate LY sales
    if brand in brands_ly:
        brand_sales_ly = filtered_sales_ly[filtered_sales_ly['Brand'] == brand]
        country_codes_ly = brand_sales_ly['Country Code'].unique()

        for country_code in country_codes_ly:
            country_sales_ly = brand_sales_ly[brand_sales_ly['Country Code'] == country_code]
            total_amount_usd_ly += convert_to_usd(country_sales_ly, country_code, currencies)

    # Calculate vs LW
    if total_amount_usd_lw > 0:
        vs_lw = ((total_amount_usd_tw - total_amount_usd_lw) / total_amount_usd_lw) * 100
    else:
        vs_lw = float('inf') if total_amount_usd_tw > 0 else 0

    # Calculate vs LY
    if total_amount_usd_ly > 0:
        vs_ly = ((total_amount_usd_tw - total_amount_usd_ly) / total_amount_usd_ly) * 100
    else:
        vs_ly = float('inf') if total_amount_usd_tw > 0 else 0

    # Append result
    result = pd.concat([
        result,
        pd.DataFrame({
            'Brand': [brand],
            'TW': [round(total_amount_usd_tw)],  # Round TW to a whole number
            'LW': [round(total_amount_usd_lw)],  # Round LW to a whole number
            'vs LW': [f"{round(vs_lw, 2)}%"],  # Round vs LW to two decimal places
            'LY': [round(total_amount_usd_ly)],  # Round LY to a whole number
            'vs LY': [f"{round(vs_ly, 2)}%"],  # Round vs LY to two decimal places
        })
    ], ignore_index=True)

# Replace brand codes with brand names using the BCodes mapping
bcodes_dict = dict(zip(bcodes['Code'], bcodes['Name']))
result['Brand'] = result['Brand'].replace(bcodes_dict)

# Add Budget and vs BUDGET columns
result = result.merge(budget_data, on='Brand', how='left').fillna({'Budget': 0})
result['vs BUDGET'] = result.apply(lambda row: f"{((row['TW'] - row['Budget']) / row['Budget']) * 100:.2f}%" if row['Budget'] > 0 else "0.00%", axis=1)

# Save the results to the WTD.xlsx file
with pd.ExcelWriter('WTD.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    result.to_excel(writer, sheet_name='Net Sales', index=False)

print("Net Sales calculation and integration with Budget completed successfully.")
print(result.head())
