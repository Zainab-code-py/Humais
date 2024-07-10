import pandas as pd

# Load Offline Sales and Net Sales sheets from the WTD.xlsx file
offline_sales_sheet = pd.read_excel('WTD.xlsx', sheet_name='Offline Sales')
net_sales_sheet = pd.read_excel('WTD.xlsx', sheet_name='Net Sales')

# Convert brand names to lowercase to make matching case-insensitive
offline_sales_sheet['Brand'] = offline_sales_sheet['Brand'].str.lower()
net_sales_sheet['Brand'] = net_sales_sheet['Brand'].str.lower()

# Merge Offline Sales and Net Sales sheets on Brand, keeping only common brands
merged_sales = pd.merge(net_sales_sheet, offline_sales_sheet, on='Brand', suffixes=('_net', '_offline'))

# Calculate Net SOB for TW, LW, and LY
merged_sales['TW'] = merged_sales['TW_net'] / (merged_sales['TW_net'] + merged_sales['TW_offline'])
merged_sales['LW'] = merged_sales['LW_net'] / (merged_sales['LW_net'] + merged_sales['LW_offline'])
merged_sales['LY'] = merged_sales['LY_net'] / (merged_sales['LY_net'] + merged_sales['LY_offline'])

# Calculate vs LW and vs LY
merged_sales['vs LW'] = (merged_sales['TW'] - merged_sales['LW']) / merged_sales['LW'] * 100
merged_sales['vs LY'] = (merged_sales['TW'] - merged_sales['LY']) / merged_sales['LY'] * 100

# Round TW, LW, and LY to 4 decimal places
merged_sales['TW'] = merged_sales['TW'].round(4)
merged_sales['LW'] = merged_sales['LW'].round(4)
merged_sales['LY'] = merged_sales['LY'].round(4)

# Round vs LW and vs LY to 2 decimal places and add a percent sign
merged_sales['vs LW'] = merged_sales['vs LW'].round(2).astype(str) + '%'
merged_sales['vs LY'] = merged_sales['vs LY'].round(2).astype(str) + '%'

# Prepare the final DataFrame
net_sob_sheet = merged_sales[['Brand', 'TW', 'LW', 'vs LW', 'LY', 'vs LY']]

# Save the Net SOB sheet to the WTD.xlsx file
with pd.ExcelWriter('WTD.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    net_sob_sheet.to_excel(writer, sheet_name='Net SOB', index=False)

# Debug information
print("Net SOB Sheet:\n", net_sob_sheet)
