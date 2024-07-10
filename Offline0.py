import pandas as pd

# Load the original CSV file
file_path = 'Offline.csv'
data = pd.read_csv(file_path)

# Define the countries to keep
valid_countries = ['UAE', 'Lebanon', 'Kuwait', 'Qatar']

# Filter the data to keep only records with the specified countries
filtered_data = data[data['Country'].isin(valid_countries)]

# Save the filtered data to a new CSV file
output_file_path = 'Filtered_OfflineSales.csv'
filtered_data.to_csv(output_file_path, index=False)

print(f"Filtered data saved to {output_file_path}")
