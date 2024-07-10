import pandas as pd
from openpyxl import load_workbook

# Load the WTD.xlsx file
wtd_file = 'WTD.xlsx'

# Load the Net Sales and Net Quantity sheets
net_sales = pd.read_excel(wtd_file, sheet_name='Net Sales')
net_quantity = pd.read_excel(wtd_file, sheet_name='Net Quantity')

# Sort the dataframes by the Brand column to ensure matching order
net_sales = net_sales.sort_values(by='Brand').reset_index(drop=True)
net_quantity = net_quantity.sort_values(by='Brand').reset_index(drop=True)

# Calculate Net ASP for each date range with zero handling
net_asp = net_sales[['Brand']].copy()


def safe_divide(numerator, denominator):
    return numerator / denominator if denominator != 0 else 0


net_asp['TW'] = net_sales.apply(lambda row: safe_divide(row['TW'], net_quantity.loc[row.name, 'TW']), axis=1).round()
net_asp['LW'] = net_sales.apply(lambda row: safe_divide(row['LW'], net_quantity.loc[row.name, 'LW']), axis=1).round()
net_asp['LY'] = net_sales.apply(lambda row: safe_divide(row['LY'], net_quantity.loc[row.name, 'LY']), axis=1).round()

# Calculate percentage difference columns with zero handling
net_asp['vs LW'] = net_asp.apply(lambda row: safe_divide(row['TW'] - row['LW'], row['LW']) * 100 if row['LW'] != 0 else 0, axis=1)
net_asp['vs LY'] = net_asp.apply(lambda row: safe_divide(row['TW'] - row['LY'], row['LY']) * 100 if row['LY'] != 0 else 0, axis=1)

# Round percentages to 2 decimal places and convert to string with percent sign
net_asp['vs LW'] = net_asp['vs LW'].round(2).astype(str) + '%'
net_asp['vs LY'] = net_asp['vs LY'].round(2).astype(str) + '%'

# Reorder columns
net_asp = net_asp[['Brand', 'TW', 'LW', 'vs LW', 'LY', 'vs LY']]

# Save the results to the WTD.xlsx file
with pd.ExcelWriter(wtd_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    net_asp.to_excel(writer, sheet_name='Net ASP', index=False)

print("Net ASP sheet has been successfully created.")
