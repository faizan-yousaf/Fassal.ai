import requests
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl
import os
from datetime import datetime

# Function to scrape crop data
def scrape_crop_data(url):
    # Send a request to the website
    response = requests.get(url)
    response.raise_for_status()  # Check for request errors

    # Parse the HTML content
    soup = BeautifulSoup(response.content, 'html.parser')

    # Find the table containing crop data
    # Adjust the selector below to match the actual table ID or class
    table = soup.find('table')  # You may need to specify an ID or class here

    # Extract headers
    headers = [header.text.strip() for header in table.find_all('th')]

    # Extract rows of data
    rows = []
    for row in table.find_all('tr')[1:]:
        cells = [cell.text.strip() for cell in row.find_all('td')]
        rows.append(cells)

    # Create a DataFrame
    df = pd.DataFrame(rows, columns=headers)

    return df

# Function to save data to Excel
def save_to_excel(df, file_name):
    if os.path.exists(file_name):
        # If file exists, append data to existing sheet
        with pd.ExcelWriter(file_name, engine='openpyxl', mode='a') as writer:
            df.to_excel(writer, index=False, sheet_name=datetime.now().strftime('%Y-%m-%d'))
    else:
        # If file doesn't exist, create a new one
        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name=datetime.now().strftime('%Y-%m-%d'))

# URL of the website to scrape
url = 'http://amis.pk/'

# Name of the Excel file to store data
excel_file_name = 'daily_crop_rates.xlsx'

# Scrape the data
try:
    crop_data = scrape_crop_data(url)
    # Save the data to Excel
    save_to_excel(crop_data, excel_file_name)
    print(f"Data saved to {excel_file_name}")
except Exception as e:
    print("An error occurred:", e)
