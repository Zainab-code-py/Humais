import pandas as pd
from datetime import datetime, timedelta

# Load the budget data from the Net Sales Budget file
budget_data = pd.read_excel('Net Sales Budget file -2024.xlsx', sheet_name='Net_Sales_Budget')

# Parsing the Date column in the budget data
budget_data['Date'] = pd.to_datetime(budget_data['Date'], format='%d/%m/%Y')

# Helper function to get the closest Sunday
def get_closest_sunday(date):
    return (date - timedelta(days=(date.weekday() + 1) % 7)).replace(hour=0, minute=0, second=0, microsecond=0)

# Get the current date and determine the TW date range
today = datetime.today()
closest_sunday = get_closest_sunday(today)

# Calculate TW range
tw_start = closest_sunday - timedelta(days=7)
tw_end = closest_sunday - timedelta(days=1)

# Debug print to check the date ranges
print(f"Closest Sunday: {closest_sunday}")
print(f"TW Range: {tw_start} to {tw_end}")

# Filtering budget data by the TW date range
filtered_budget_tw = budget_data[(budget_data['Date'] >= tw_start) & (budget_data['Date'] <= tw_end)].copy()

# Debug print to check the filtering process
print(f"All Brands in Budget Data: {budget_data['Brand'].unique()}")
print(f"Filtered Brands in TW: {filtered_budget_tw['Brand'].unique()}")

# Ensure Budget values are numeric (handle $ and commas)
filtered_budget_tw['Budget'] = filtered_budget_tw['Budget'].replace('[\$,]', '', regex=True).astype(float)

# Calculate the total budget for each brand
total_budgets = filtered_budget_tw.groupby('Brand')['Budget'].sum().reset_index()

# Save the results to a CSV file
total_budgets.to_csv('Budget_Summary_TW.csv', index=False)

print("Budget calculation completed successfully.")
print(total_budgets)

# Save the results to an Excel file
with pd.ExcelWriter('Budget_Summary_TW.xlsx', mode='w', engine='openpyxl') as writer:
    total_budgets.to_excel(writer, sheet_name='Budget Summary', index=False)
