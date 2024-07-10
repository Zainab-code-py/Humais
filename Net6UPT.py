import pandas as pd
from openpyxl import load_workbook

# Load the WTD.xlsx file
wtd_file = 'WTD.xlsx'

# Load the Net Quantity and Net Transactions sheets
net_quantity = pd.read_excel(wtd_file, sheet_name='Net Quantity')
net_transactions = pd.read_excel(wtd_file, sheet_name='Net Transactions')

# Ensure Brand columns are sorted and reset index for both DataFrames
net_quantity = net_quantity.sort_values(by='Brand').reset_index(drop=True)
net_transactions = net_transactions.sort_values(by='Brand').reset_index(drop=True)

# Merge the dataframes to ensure all brands are included
net_upt = pd.merge(net_transactions[['Brand']], net_quantity, on='Brand', how='left', suffixes=('_quantity', ''))
net_upt = pd.merge(net_upt, net_transactions, on='Brand', how='left', suffixes=('_quantity', '_transactions'))


# Calculate Net UPT for each date range with zero handling
def safe_divide(numerator, denominator):
    return numerator / denominator if denominator != 0 else 0


net_upt['TW'] = net_upt.apply(lambda row: safe_divide(row['TW_quantity'], row['TW_transactions']), axis=1).round()
net_upt['LW'] = net_upt.apply(lambda row: safe_divide(row['LW_quantity'], row['LW_transactions']), axis=1).round()
net_upt['LY'] = net_upt.apply(lambda row: safe_divide(row['LY_quantity'], row['LY_transactions']), axis=1).round()

# Calculate percentage difference columns with zero handling
net_upt['vs LW'] = net_upt.apply(lambda row: safe_divide(row['TW'] - row['LW'], row['LW']) * 100 if row['LW'] != 0 else 0, axis=1)
net_upt['vs LY'] = net_upt.apply(lambda row: safe_divide(row['TW'] - row['LY'], row['LY']) * 100 if row['LY'] != 0 else 0, axis=1)

# Round percentages to 2 decimal places and convert to string with percent sign
net_upt['vs LW'] = net_upt['vs LW'].round(2).astype(str) + '%'
net_upt['vs LY'] = net_upt['vs LY'].round(2).astype(str) + '%'

# Reorder columns and handle missing values by filling with zeros
net_upt = net_upt[['Brand', 'TW', 'LW', 'vs LW', 'LY', 'vs LY']].fillna(0)

# Save the results to the WTD.xlsx file
with pd.ExcelWriter(wtd_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    net_upt.to_excel(writer, sheet_name='Net UPT', index=False)

print("Net UPT sheet has been successfully created.")
