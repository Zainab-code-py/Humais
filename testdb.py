import csv
import pymongo
from azure.storage.blob import BlobServiceClient
from datetime import datetime

def transaction():
    # MongoDB connection details
    mongodb_uri = 'mongodb+srv://mongoAdmin:pR9d%7DyqM%29.z%3A%3F%22E@dbmongosp.mongocluster.cosmos.azure.com/?tls=true&authMechanism=SCRAM-SHA-256&retrywrites=false&maxIdleTimeMS=120000'
    database_name = 'Datacrypt'
    collection_name = 'MERCHANT_TRANSACTIONS'

    # Azure Blob Storage connection details
    azure_connection_string = 'DefaultEndpointsProtocol=https;AccountName=navprojectstorage;AccountKey=hIjjya8JYmtakPTymBNWetZV07I6zNOUalk+SiTaNZY09sFXbc2L7M5TXj2VO9dAhTtYCihAseNObVeGL1CZmg==;EndpointSuffix=core.windows.net'
    container_name = 'safexpay'
    blob_name = 'transaction.csv'  # Custom blob name

    # Fieldnames as per the exact order required
    fieldnames = [
        '_id', 'name', 'code', 'ag_id', 'ag_name', 'paybyface', 'transaction_request._id',
        'order_no', 'paymode_id', 'scheme_id', 'emi_months', 'unique_id', 'ship_days', 
        'address_count', 'item_count', 'item_value', 'bill_address', 'bill_city', 
        'bill_state', 'bill_country', 'bill_zip', 'ship_address', 'ship_city', 
        'ship_state', 'ship_country', 'ship_zip', 'amount', 'country_code', 'currency_code', 
        'udf_1', 'udf_2', 'item_category', 'is_logged_in', 'creation_date', 'pg_ref', 
        'ag_ref', 'status', 'sub_status', 'fee', 'me_igst', 'pg_id', 'description', 
        'me_total_amount', 'me_net_amount', 'operating_mode'
    ]

    try:
        # Connect to MongoDB
        client = pymongo.MongoClient(mongodb_uri)
        db = client[database_name]
        collection = db[collection_name]

        # Fetch data from MongoDB collection
        documents = list(collection.find())  # Convert cursor to list for reuse

        # Check if documents are empty
        if not documents:
            raise ValueError("No documents found in the MongoDB collection.")

        # Write MongoDB data to CSV
        csv_file_path = 'transaction.csv'  # Custom local file name
        with open(csv_file_path, 'w', newline='') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            for document in documents:
                # Extract nested fields
                transaction_request_id = document.get('transaction_request', {}).get('_id', '')
                # Format the creation_date field
                creation_date = document.get('creation_date', '')
                if isinstance(creation_date, datetime):
                    creation_date = creation_date.isoformat() + 'Z'
                elif isinstance(creation_date, str) and creation_date:
                    creation_date = datetime.strptime(creation_date, '%Y-%m-%dT%H:%M:%S.%fZ').isoformat() + 'Z'
                # Create the row dictionary
                row = {field: document.get(field, '') for field in fieldnames}
                row['transaction_request._id'] = transaction_request_id
                row['creation_date'] = creation_date
                writer.writerow(row)

        # Connect to Azure Blob Storage
        blob_service_client = BlobServiceClient.from_connection_string(azure_connection_string)
        blob_client = blob_service_client.get_blob_client(container=container_name, blob=blob_name)

        # Upload the CSV file to Azure Blob Storage
        with open(csv_file_path, 'rb') as data:
            blob_client.upload_blob(data, overwrite=True)

        print(f"CSV file '{csv_file_path}' has been uploaded to Azure Blob Storage as '{blob_name}' successfully.")

    except Exception as e:
        print(f"An error occurred: {e}")

transaction()
