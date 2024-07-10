import pandas as pd

# Define the file name
file_name = 'WTD.xlsx'

# Load the existing sheets from the WTD.xlsx file
with pd.ExcelFile(file_name) as xls:
    gross_sales = pd.read_excel(xls, sheet_name='Gross Sales')
    net_sales = pd.read_excel(xls, sheet_name='Net Sales')

# Ensure the brand columns are strings
gross_sales['Brand'] = gross_sales['Brand'].astype(str)
net_sales['Brand'] = net_sales['Brand'].astype(str)

# Find common brands in both sheets
common_brands = set(gross_sales['Brand']).intersection(set(net_sales['Brand']))

# Filter the sheets to only include common brands
gross_sales_filtered = gross_sales[gross_sales['Brand'].isin(common_brands)].set_index('Brand')
net_sales_filtered = net_sales[net_sales['Brand'].isin(common_brands)].set_index('Brand')

# Calculate Gross to Net% for TW, LW, and LY
gross_to_net = pd.DataFrame()
gross_to_net['TW'] = (net_sales_filtered['TW'] / gross_sales_filtered['TW']) * 100
gross_to_net['LW'] = (net_sales_filtered['LW'] / gross_sales_filtered['LW']) * 100
gross_to_net['LY'] = (net_sales_filtered['LY'] / gross_sales_filtered['LY']) * 100

# Replace NaN values resulting from division by zero with 0
# gross_to_net = gross_to_net.fillna(0)

# Round TW, LW, and LY values to two decimal places and add '%'
gross_to_net['TW'] = gross_to_net['TW'].round(2).astype(str) + '%'
gross_to_net['LW'] = gross_to_net['LW'].round(2).astype(str) + '%'
gross_to_net['LY'] = gross_to_net['LY'].round(2).astype(str) + '%'

# Calculate vs LW and vs LY
gross_to_net['vs LW'] = (((net_sales_filtered['TW'] / gross_sales_filtered['TW']) - (net_sales_filtered['LW'] / gross_sales_filtered['LW'])) / (net_sales_filtered['LW'] / gross_sales_filtered['LW'])) * 100
gross_to_net['vs LY'] = (((net_sales_filtered['TW'] / gross_sales_filtered['TW']) - (net_sales_filtered['LY'] / gross_sales_filtered['LY'])) / (net_sales_filtered['LY'] / gross_sales_filtered['LY'])) * 100

# Round vs LW and vs LY to two decimal places and add '%'
gross_to_net['vs LW'] = gross_to_net['vs LW'].round(2).astype(str) + '%'
gross_to_net['vs LY'] = gross_to_net['vs LY'].round(2).astype(str) + '%'

# Reset the index to make Brand a column
gross_to_net = gross_to_net.reset_index()

# Reorder columns to place Brand first
gross_to_net = gross_to_net[['Brand', 'TW', 'LW', 'vs LW', 'LY', 'vs LY']]

# Print the result DataFrame
print(gross_to_net)

gross_to_net.to_csv('WTDG2N.csv', index=False)
# Save the result to the WTD.xlsx file
with pd.ExcelWriter(file_name, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    gross_to_net.to_excel(writer, sheet_name='Gross to Net%', index=False)
