import pandas as pd
from datetime import datetime, timedelta

# Load the data with specified dtypes to avoid mixed type warning
gross_sales = pd.read_csv('GrossSales_updated.csv', dtype={'Country Code': str, 'Brand': str}, low_memory=False)
currencies = pd.read_excel('Currencies.xlsx')
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

# Convert Order Date to datetime and filter sales data for TW and LW
gross_sales['SFCC Order Date'] = pd.to_datetime(gross_sales['SFCC Order Date'], format='%d/%m/%Y', dayfirst=True)
filtered_sales_tw = gross_sales[
    (gross_sales['SFCC Order Date'] >= tw_start) &
    (gross_sales['SFCC Order Date'] <= tw_end)
    ].copy()
filtered_sales_lw = gross_sales[
    (gross_sales['SFCC Order Date'] >= lw_start) &
    (gross_sales['SFCC Order Date'] <= lw_end)
    ].copy()

# Debug prints to check if the filtering is correct
print(f"Filtered Sales Data TW (Head):\n{filtered_sales_tw.head()}")
print(f"Filtered Sales Data TW Dates:\n{filtered_sales_tw['SFCC Order Date'].unique()}")
print(f"Filtered Sales Data LW (Head):\n{filtered_sales_lw.head()}")
print(f"Filtered Sales Data LW Dates:\n{filtered_sales_lw['SFCC Order Date'].unique()}")

# Convert Amount and currency rates to appropriate data types
filtered_sales_tw.loc[:, 'Amount'] = pd.to_numeric(filtered_sales_tw['Amount'], errors='coerce')
filtered_sales_lw.loc[:, 'Amount'] = pd.to_numeric(filtered_sales_lw['Amount'], errors='coerce')
currency_columns = [col for col in currencies.columns if col not in ['Country_Code', 'Country_Name']]
for col in currency_columns:
    currencies[col] = pd.to_numeric(currencies[col], errors='coerce')

# Remove negative values
filtered_sales_tw = filtered_sales_tw[filtered_sales_tw['Amount'] >= 0]
filtered_sales_lw = filtered_sales_lw[filtered_sales_lw['Amount'] >= 0]

# Prepare the result DataFrame
result = pd.DataFrame(columns=['Brand', 'TW', 'LW', 'vs LW', 'LY', 'vs LY'])

# Get unique brands from TW, LW, and LY sales data
brands_tw = filtered_sales_tw['Brand'].unique()
brands_lw = filtered_sales_lw['Brand'].unique()

# Calculate LY dates
random_date_tw = filtered_sales_tw['SFCC Order Date'].iloc[0]  # Using the first date from TW range
random_date_sort = int(random_date_tw.strftime('%Y%m%d'))

# Step 1: Find the Trading Week for the random_date_tw
tw_week = calendar.loc[calendar['Sort'] == random_date_sort, 'Trading Week'].values[0]

# Step 2: Find the same Trading Week in the previous year (2023)
previous_year = 2023
ly_week = calendar.loc[(calendar['Trading Week'] == tw_week) & (calendar['CalYear'] == previous_year)]

# Step 3: Determine all dates for that trading week
ly_dates = ly_week['Sort'].values
ly_dates = [datetime.strptime(str(date), '%Y%m%d') for date in ly_dates]

# Filter sales data for LY dates
filtered_sales_ly = gross_sales[
    (gross_sales['SFCC Order Date'].isin(ly_dates))
    ]

brands_ly = filtered_sales_ly['Brand'].unique()
all_brands = set(brands_tw).union(set(brands_lw)).union(set(brands_ly))


# Function to get conversion rate for a specific date
def get_conversion_rate(date, country_code):
    month_year = date.strftime('%b-%Y')
    return currencies.loc[currencies['Country_Code'] == country_code, month_year].values[0]


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

            for date in country_sales_tw['SFCC Order Date'].unique():
                daily_sales_tw = country_sales_tw[country_sales_tw['SFCC Order Date'] == date]
                total_amount_tw = daily_sales_tw['Amount'].sum()

                conversion_rate_tw = get_conversion_rate(date, country_code)
                if conversion_rate_tw > 0:
                    amount_usd_tw = total_amount_tw / conversion_rate_tw
                else:
                    amount_usd_tw = 0

                total_amount_usd_tw += amount_usd_tw

    # Calculate LW sales
    if brand in brands_lw:
        brand_sales_lw = filtered_sales_lw[filtered_sales_lw['Brand'] == brand]
        country_codes_lw = brand_sales_lw['Country Code'].unique()

        for country_code in country_codes_lw:
            country_sales_lw = brand_sales_lw[brand_sales_lw['Country Code'] == country_code]

            for date in country_sales_lw['SFCC Order Date'].unique():
                daily_sales_lw = country_sales_lw[country_sales_lw['SFCC Order Date'] == date]
                total_amount_lw = daily_sales_lw['Amount'].sum()

                conversion_rate_lw = get_conversion_rate(date, country_code)
                if conversion_rate_lw > 0:
                    amount_usd_lw = total_amount_lw / conversion_rate_lw
                else:
                    amount_usd_lw = 0

                total_amount_usd_lw += amount_usd_lw

    # Calculate LY sales
    if brand in brands_ly:
        brand_sales_ly = filtered_sales_ly[filtered_sales_ly['Brand'] == brand]
        country_codes_ly = brand_sales_ly['Country Code'].unique()

        for country_code in country_codes_ly:
            country_sales_ly = brand_sales_ly[brand_sales_ly['Country Code'] == country_code]

            for date in country_sales_ly['SFCC Order Date'].unique():
                daily_sales_ly = country_sales_ly[country_sales_ly['SFCC Order Date'] == date]
                total_amount_ly = daily_sales_ly['Amount'].sum()

                conversion_rate_ly = get_conversion_rate(date, country_code)
                if conversion_rate_ly > 0:
                    amount_usd_ly = total_amount_ly / conversion_rate_ly
                else:
                    amount_usd_ly = 0

                total_amount_usd_ly += amount_usd_ly

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
            'vs LW': [f"{round(vs_lw, 2)}%"],  # Round vs LW to 2 decimal places and add %
            'LY': [round(total_amount_usd_ly)],  # Round LY to a whole number
            'vs LY': [f"{round(vs_ly, 2)}%"]  # Round vs LY to 2 decimal places and add %
        })
    ], ignore_index=True)

# Replace brand codes with brand names
result = result.merge(bcodes, left_on='Brand', right_on='Code')
result = result.drop(columns=['Brand', 'Code'])
result = result.rename(columns={'Name': 'Brand'})

# Reorder columns to place Brand first
result = result[['Brand', 'TW', 'LW', 'vs LW', 'LY', 'vs LY']]

# Print the result DataFrame
print(result)

# Save the result to a CSV file and Excel file
result.to_csv('WTD.csv', index=False)
with pd.ExcelWriter('WTD.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    result.to_excel(writer, sheet_name='Gross Sales', index=False)
