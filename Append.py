import pandas as pd
from azure.storage.blob import BlobServiceClient
import io

# Azure Storage connection details
connection_string = 'DefaultEndpointsProtocol=https;AccountName=navprojectstorage;AccountKey=hBRb0iTZ8w0fA7AXjgF2eeIkjIDBr/41lMBcIysZwrW8kxmGGzDHKY82eGy2cGWZjfYUqZMYbBDLeegpis2a1A==;EndpointSuffix=core.windows.net'
container_name = 'nav'  # Use 'nav' as the container name
file_paths = [
    'Decathlon/DL.csv',
    'Decathlon/DO.csv',
    'Decathlon/KWT.csv',
    'Decathlon/QAT.csv',
    'Decathlon/UAE.csv',
    'Decathlon/BHR.csv'
]

# Mapping of file paths to currency names
currency_map = {
    'Decathlon/DL.csv': 'LBP',
    'Decathlon/DO.csv': 'OMR',
    'Decathlon/KWT.csv': 'KWD',
    'Decathlon/QAT.csv': 'QAR',
    'Decathlon/UAE.csv': 'AED',
    'Decathlon/BHR.csv': 'BHD'
}

# Initialize BlobServiceClient
blob_service_client = BlobServiceClient.from_connection_string(connection_string)
container_client = blob_service_client.get_container_client(container_name)

# Verify if the container exists
try:
    container_client.get_container_properties()
    print(f"Container '{container_name}' exists.")
except Exception as e:
    print(f"Container '{container_name}' does not exist or cannot be accessed: {e}")
    exit(1)

# List to hold the dataframes
dataframes = []

# Process each file
for file_path in file_paths:
    try:
        print(f"Processing file: {file_path}")
        blob_client = container_client.get_blob_client(file_path)
        download_stream = blob_client.download_blob()
        data = download_stream.readall()

        # Load the CSV data into a DataFrame with low_memory=False
        df = pd.read_csv(io.BytesIO(data), low_memory=False)

        # Append the dataframe to the list
        dataframes.append(df)
    except Exception as e:
        print(f"Failed to process {file_path}: {e}")

# Concatenate all dataframes in the list, ensuring columns are aligned
if dataframes:
    # Find the common columns across all dataframes
    common_columns = list(set.intersection(*(set(df.columns) for df in dataframes)))

    # Align dataframes to have the same columns
    aligned_dataframes = [df[common_columns] for df in dataframes]

    # Concatenate the aligned dataframes
    appended_df = pd.concat(aligned_dataframes, ignore_index=True)

    # Add the "Currency Name" column based on the file paths
    currency_column = []
    for file_path, df in zip(file_paths, dataframes):
        currency_column.extend([currency_map[file_path]] * len(df))

    appended_df['Currency Name'] = currency_column

    # Ensure "Currency Name" is at the end
    final_columns = common_columns + ['Currency Name']
    appended_df = appended_df[final_columns]

    # Save the appended dataframe to Azure Blob Storage
    output_file_path = 'Decathlon/appended_file.csv'
    output_blob_client = container_client.get_blob_client(output_file_path)
    output_blob_client.upload_blob(appended_df.to_csv(index=False), overwrite=True)

    print(f'Appended file saved as {output_file_path}')
else:
    print("No dataframes to concatenate.")
