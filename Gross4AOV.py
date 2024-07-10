import pandas as pd

# Load the existing sheets from the WTD.xlsx file
with pd.ExcelFile('WTD.xlsx') as xls:
    gross_sales = pd.read_excel(xls, sheet_name='Gross Sales')
    gross_transactions = pd.read_excel(xls, sheet_name='Gross Transactions')

# Ensure the brand columns are strings
gross_sales['Brand'] = gross_sales['Brand'].astype(str)
gross_transactions['Brand'] = gross_transactions['Brand'].astype(str)

# Find common brands in both sheets
common_brands = set(gross_sales['Brand']).intersection(set(gross_transactions['Brand']))

# Filter the sheets to only include common brands
gross_sales_filtered = gross_sales[gross_sales['Brand'].isin(common_brands)].set_index('Brand')
gross_transactions_filtered = gross_transactions[gross_transactions['Brand'].isin(common_brands)].set_index('Brand')

# Calculate AOV for TW, LW, and LY
gross_aov = pd.DataFrame()
gross_aov['TW'] = gross_sales_filtered['TW'] / gross_transactions_filtered['TW']
gross_aov['LW'] = gross_sales_filtered['LW'] / gross_transactions_filtered['LW']
gross_aov['LY'] = gross_sales_filtered['LY'] / gross_transactions_filtered['LY']

# Replace NaN values resulting from division by zero with 0
gross_aov = gross_aov.fillna(0)

# Calculate vs LW and vs LY
gross_aov['vs LW'] = ((gross_aov['TW'] - gross_aov['LW']) / gross_aov['LW']) * 100
gross_aov['vs LY'] = ((gross_aov['TW'] - gross_aov['LY']) / gross_aov['LY']) * 100

# Round vs LW and vs LY to 2 decimal places and add %
gross_aov['vs LW'] = gross_aov['vs LW'].round(2).astype(str) + '%'
gross_aov['vs LY'] = gross_aov['vs LY'].round(2).astype(str) + '%'

# Reset the index to make Brand a column
gross_aov = gross_aov.reset_index()

# Reorder columns to place Brand first
gross_aov = gross_aov[['Brand', 'TW', 'LW', 'vs LW', 'LY', 'vs LY']]

# Print the result DataFrame
print(gross_aov)

# Save the result to the WTD.xlsx file
with pd.ExcelWriter('WTD.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    gross_aov.to_excel(writer, sheet_name='Gross AOV', index=False)
