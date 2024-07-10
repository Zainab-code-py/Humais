import pandas as pd

# Load the GrossSales.csv file
gross_sales = pd.read_csv('GrossSales.csv')

# Define the brand codes to be replaced
brand_codes_to_replace = ['C7', 'D2']

# Replace the brand codes with subbrand codes
gross_sales['Brand'] = gross_sales.apply(
    lambda row: row['SubBrand Code'] if row['Brand'] in brand_codes_to_replace else row['Brand'], axis=1)

# Save the updated DataFrame back to the CSV file
gross_sales.to_csv('GrossSales_updated.csv', index=False)

# Print the first few rows to verify the changes
print(gross_sales.head())
