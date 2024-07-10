import pandas as pd

# Define the file name
file_name = 'WTD.xlsx'

# Load the existing sheets from the WTD.xlsx file
with pd.ExcelFile(file_name) as xls:
    gross_sales = pd.read_excel(xls, sheet_name='Gross Sales')
    gross_quantity = pd.read_excel(xls, sheet_name='Gross Quantity')

# Ensure the brand columns are strings
gross_sales['Brand'] = gross_sales['Brand'].astype(str)
gross_quantity['Brand'] = gross_quantity['Brand'].astype(str)

# Find common brands in both sheets
common_brands = set(gross_sales['Brand']).intersection(set(gross_quantity['Brand']))

# Filter the sheets to only include common brands
gross_sales_filtered = gross_sales[gross_sales['Brand'].isin(common_brands)].set_index('Brand')
gross_quantity_filtered = gross_quantity[gross_quantity['Brand'].isin(common_brands)].set_index('Brand')

# Calculate ASP for TW, LW, and LY
gross_asp = pd.DataFrame()
gross_asp['TW'] = gross_sales_filtered['TW'] / gross_quantity_filtered['TW']
gross_asp['LW'] = gross_sales_filtered['LW'] / gross_quantity_filtered['LW']
gross_asp['LY'] = gross_sales_filtered['LY'] / gross_quantity_filtered['LY']

# Replace NaN values resulting from division by zero with 0
gross_asp = gross_asp.fillna(0)

# Round TW, LW, and LY values to whole numbers
gross_asp['TW'] = gross_asp['TW'].round(0)
gross_asp['LW'] = gross_asp['LW'].round(0)
gross_asp['LY'] = gross_asp['LY'].round(0)

# Calculate vs LW and vs LY
gross_asp['vs LW'] = ((gross_asp['TW'] - gross_asp['LW']) / gross_asp['LW']) * 100
gross_asp['vs LY'] = ((gross_asp['TW'] - gross_asp['LY']) / gross_asp['LY']) * 100

# Round vs LW and vs LY to 2 decimal places and add %
gross_asp['vs LW'] = gross_asp['vs LW'].round(2).astype(str) + '%'
gross_asp['vs LY'] = gross_asp['vs LY'].round(2).astype(str) + '%'

# Reset the index to make Brand a column
gross_asp = gross_asp.reset_index()

# Reorder columns to place Brand first
gross_asp = gross_asp[['Brand', 'TW', 'LW', 'vs LW', 'LY', 'vs LY']]

# Print the result DataFrame
print(gross_asp)

# Save the result to the WTD.xlsx file
with pd.ExcelWriter(file_name, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    gross_asp.to_excel(writer, sheet_name='Gross ASP', index=False)
