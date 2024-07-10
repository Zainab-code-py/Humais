import pandas as pd

# Define the file name
file_name = 'WTD.xlsx'

# Load the existing sheets from the WTD.xlsx file
with pd.ExcelFile(file_name) as xls:
    gross_quantity = pd.read_excel(xls, sheet_name='Gross Quantity')
    gross_transactions = pd.read_excel(xls, sheet_name='Gross Transactions')

# Ensure the brand columns are strings
gross_quantity['Brand'] = gross_quantity['Brand'].astype(str)
gross_transactions['Brand'] = gross_transactions['Brand'].astype(str)

# Find common brands in both sheets
common_brands = set(gross_quantity['Brand']).intersection(set(gross_transactions['Brand']))

# Filter the sheets to only include common brands
gross_quantity_filtered = gross_quantity[gross_quantity['Brand'].isin(common_brands)].set_index('Brand')
gross_transactions_filtered = gross_transactions[gross_transactions['Brand'].isin(common_brands)].set_index('Brand')

# Calculate UPT for TW, LW, and LY
gross_upt = pd.DataFrame()
gross_upt['TW'] = gross_quantity_filtered['TW'] / gross_transactions_filtered['TW']
gross_upt['LW'] = gross_quantity_filtered['LW'] / gross_transactions_filtered['LW']
gross_upt['LY'] = gross_quantity_filtered['LY'] / gross_transactions_filtered['LY']

# Replace NaN values resulting from division by zero with 0
gross_upt = gross_upt.fillna(0)

# Round TW, LW, and LY values
gross_upt['TW'] = gross_upt['TW'].round(2)
gross_upt['LW'] = gross_upt['LW'].round(2)
gross_upt['LY'] = gross_upt['LY'].round(2)

# Calculate vs LW and vs LY
gross_upt['vs LW'] = ((gross_upt['TW'] - gross_upt['LW']) / gross_upt['LW']) * 100
gross_upt['vs LY'] = ((gross_upt['TW'] - gross_upt['LY']) / gross_upt['LY']) * 100

# Round vs LW and vs LY to 2 decimal places and add %
gross_upt['vs LW'] = gross_upt['vs LW'].round(2).astype(str) + '%'
gross_upt['vs LY'] = gross_upt['vs LY'].round(2).astype(str) + '%'

# Reset the index to make Brand a column
gross_upt = gross_upt.reset_index()

# Reorder columns to place Brand first
gross_upt = gross_upt[['Brand', 'TW', 'LW', 'vs LW', 'LY', 'vs LY']]

# Print the result DataFrame
print(gross_upt)

# Save the result to the WTD.xlsx file
with pd.ExcelWriter(file_name, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    gross_upt.to_excel(writer, sheet_name='Gross UPT', index=False)
