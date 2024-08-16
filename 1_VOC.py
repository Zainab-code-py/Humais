#REMEMBER TO CHANGE DATE

import pandas as pd
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate

# Set the hardcoded date for testing
hardcoded_date = datetime(2024, 7, 15).date()
formatted_date = hardcoded_date.strftime('%d/%m/%Y')

# Read the CSV and Excel files
combined_df = pd.read_csv('combined_new.csv', low_memory=False)
bucode_df = pd.read_excel('BuCode.xlsx', sheet_name=0, header=0)
store_emails_df = pd.read_excel('Azadea Stores.xlsx', sheet_name=0, header=0)

# Ensure the relevant columns are named properly
store_emails_df = store_emails_df.rename(
    columns={'Email1': 'Email', 'BusinessUnitCode': 'BUCode', 'Country': 'Country'}
)

# Convert the 'Submitted at' column to datetime format
combined_df['Submitted at'] = pd.to_datetime(combined_df['Submitted at'])


# Function to process data for a given date and country
def process_date(date, combined_df, bucode_df, country):
    # Get BU Codes for the given country from bucode_df
    bu_codes_for_country = bucode_df[bucode_df['Country'] == country]['BU'].unique()

    # Filter the combined.csv file to the given date and BU Codes for the country
    filtered_df = combined_df[
        (combined_df['Submitted at'].dt.date == date) & (combined_df['BUCode'].isin(bu_codes_for_country))
        ]

    # Get the unique BuCodes from the filtered DataFrame
    filtered_bucodes = filtered_df['BUCode'].unique()

    # Get the unique BuCodes from the BuCode.xlsx file for the specified country
    excel_bucodes = bu_codes_for_country

    # Find BuCodes present in BuCode.xlsx but not in combined.csv (no surveys)
    missing_bucodes = [bucode for bucode in excel_bucodes if bucode not in filtered_bucodes]

    # Identify BuCodes that have any "Completed" status
    completed_bucodes = filtered_df[filtered_df['Completion Status'] == 'Completed']['BUCode'].unique()

    # Identify BuCodes that have "Partially Completed" status and exclude those with "Completed" status
    partially_completed_only_bucodes = filtered_df[
        (filtered_df['Completion Status'] == 'Partially Completed') &
        (~filtered_df['BUCode'].isin(completed_bucodes))
        ]['BUCode'].unique()

    # Combine the missing and partially completed only BuCodes
    combined_bucodes = list(set(missing_bucodes + list(partially_completed_only_bucodes)))

    # Filter the BuCode.xlsx DataFrame to get the relevant rows for the specified country
    filtered_bucode_df = bucode_df[(bucode_df['BU'].isin(combined_bucodes)) & (bucode_df['Country'] == country)].copy()

    # Add the Date column
    filtered_bucode_df.loc[:, 'Date'] = date

    # Reorder columns to match the required output format
    output_df = filtered_bucode_df[['Date', 'BU', 'Brand', 'Location', 'Country']]

    # Rename columns to match the required output format
    output_df.columns = ['Date', 'BUCode', 'Brand', 'Location', 'Country']

    return output_df


# Email configuration
sender_email = 'analytics@datacrypt.ae'
sender_password = 'ykta lwxs degx ihoi'
recipients = ['zainabsiddiqui@datacrypt.ae', 'stephany.mady@azadea.com']  # Main recipients
subject_template = 'Survey Status for {country}'
body_template = '''
Dear All

We hope you are all doing well.

Please note that your store did not submit any VOC survey yesterday. Please remember that you can reach your weekly target by having a minimum of 5 surveys daily.
Let's work together to make gathering VOC feedback a top priority.

Regards,
Azadea VOC Customer Experience Department
'''

# Process the hardcoded date for each country
countries = bucode_df['Country'].unique()
for country in countries:
    country_output_df = process_date(hardcoded_date, combined_df, bucode_df, country)

    # Save the output dataframe to a CSV file
    output_file = f'output_{country}.csv'
    country_output_df.to_csv(output_file, index=False)

    print(f'Saved file for {country}: {output_file}')

    # Get the BU codes with no surveys or only partially completed surveys
    bu_codes_list = country_output_df['BUCode'].tolist()

    # Get the emails for those specific BU codes
    store_emails = store_emails_df[store_emails_df['BUCode'].isin(bu_codes_list)]
    cc_emails = store_emails['Email'].dropna().tolist()
    cc_emails += store_emails['Email4'].dropna().tolist()
    cc_emails += store_emails['Email5'].dropna().tolist()
    cc_emails += store_emails['Email6'].dropna().tolist()

    # Create the email
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ", ".join(recipients)
    msg['Cc'] = ", ".join(cc_emails)  # Add specific store emails in CC
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject_template.format(country=country)

    # Attach the email body with the dynamic date
    msg.attach(MIMEText(body_template.format(date=formatted_date)))

    # Send the email
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:  # Use your SMTP server and port
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipients + cc_emails, msg.as_string())
        print(f'Email sent successfully for {country}.')
    except Exception as e:
        print(f'Failed to send email for {country}: {e}')
