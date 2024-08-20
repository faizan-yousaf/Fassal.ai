from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import pandas as pd
import openpyxl
import os
from datetime import datetime

# Setup Chrome options
chrome_options = Options()
chrome_options.add_argument("--headless")  # Run in headless mode (no GUI)
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

# Setup the WebDriver
service = Service('/path/to/chromedriver')  # Replace with your chromedriver path
driver = webdriver.Chrome(service=service, options=chrome_options)

# Function to scrape crop data
def scrape_crop_data(url):
    # Load the webpage
    driver.get(url)

    # Allow some time for the page to load completely
    driver.implicitly_wait(10)

    # Locate the table containing crop data
    # Adjust the XPath or CSS selector based on the actual structure
    table = driver.find_element(By.XPATH, '//*[@id="example"]')  # Change XPath as needed

    # Extract headers
    headers = [header.text for header in table.find_elements(By.TAG_NAME, 'th')]

    # Extract rows of data
    rows = []
    for row in table.find_elements(By.TAG_NAME, 'tr')[1:]:
        cells = [cell.text for cell in row.find_elements(By.TAG_NAME, 'td')]
        if cells:  # Ensure the row is not empty
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
finally:
    # Close the WebDriver
    driver.quit()
