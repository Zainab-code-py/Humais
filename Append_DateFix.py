from azure.storage.blob import BlobServiceClient
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta

# Azure Storage connection string
connection_string = "DefaultEndpointsProtocol=https;AccountName=azadea;AccountKey=a3T4kBCWOZJb3ndoPdMYwfyvyAKlQCkBHkqOB6FEeSAf2w2slI0eDwQUtsIdHRQYS4ig5w4OH6miW+FWSRvA2w==;EndpointSuffix=core.windows.net"
container_name = "collated-new"
blob_service_client = BlobServiceClient.from_connection_string(connection_string)
container_client = blob_service_client.get_container_client(container_name)

# Download the main tally file
blob_client_tally = container_client.get_blob_client("VOCServiceRecovery_tally.xlsx")
tally_stream = BytesIO()
blob_client_tally.download_blob().readinto(tally_stream)
tally_stream.seek(0)
df_tally = pd.read_excel(tally_stream)

# Download the input file
blob_client_input = container_client.get_blob_client("VOCServiceRecovery_tally_input.xlsx")
input_stream = BytesIO()
blob_client_input.download_blob().readinto(input_stream)
input_stream.seek(0)
df_input = pd.read_excel(input_stream)

# Append the data from input to the main tally
df_combined = pd.concat([df_tally, df_input], ignore_index=True)

# Function to convert Excel serial date to datetime
def excel_serial_date_to_datetime(serial):
    if isinstance(serial, datetime):
        return serial
    if pd.isna(serial):
        return pd.NaT
    try:
        return datetime(1900, 1, 1) + timedelta(days=int(serial) - 2)
    except (ValueError, TypeError):
        return pd.NaT

# Convert the date columns
df_combined['Date of Negative Response'] = df_combined['Date of Negative Response'].apply(excel_serial_date_to_datetime)
df_combined['Date of Reaching out'] = df_combined['Date of Reaching out'].apply(lambda x: excel_serial_date_to_datetime(x) if isinstance(x, (int, float)) else pd.NaT)

# Format the dates as mm/dd/yyyy
df_combined['Date of Negative Response'] = df_combined['Date of Negative Response'].dt.strftime('%m/%d/%Y')
df_combined['Date of Reaching out'] = df_combined['Date of Reaching out'].dt.strftime('%m/%d/%Y')

# Upload the updated file back to Azure Blob Storage
output_stream = BytesIO()
df_combined.to_excel(output_stream, index=False)
output_stream.seek(0)
blob_client_tally.upload_blob(output_stream, overwrite=True)

print("Data has been appended, dates have been formatted, and the updated file has been uploaded to Azure Blob Storage.")
