import time
import os
import pandas as pd
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup

# Setup Chrome WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Function to scrape data from the current page
def scrape_current_page(driver):
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    data_rows = []

    # Select the main container for each member
    members = soup.select('div[class*="col-sm-7"]')  # Adjust the selector to match partial class names if needed

    for member in members:
        try:
            # Extract Name
            name = member.select_one('div.member_name').get_text(strip=True) if member.select_one('div.member_name') else None
            
            # Extract Work
            work = member.select_one('p.member_work').get_text(strip=True) if member.select_one('p.member_work') else None
            
            # Extract Address
            address_detail = member.find('p', text='Add.		')  # Find the 'Add.' label
            address = address_detail.find_next('span').find_next('span').get_text(strip=True) if address_detail else None
            
            # Extract Executive
            executive_detail = member.find('p', text='Exc.		')  # Find the 'Exc.' label
            executive = executive_detail.find_next('span').find_next('span').get_text(strip=True) if executive_detail else None
            
            # Extract Phone
            phone_detail = member.find('p', text='Ph.		')  # Find the 'Ph.' label
            phone = phone_detail.find_next('span').find_next('span').get_text(strip=True) if phone_detail else None
            
            # Extract Mobile
            mobile_detail = member.find('p', text='Mob.')  # Find the 'Mob.' label
            mobile = mobile_detail.find_next('span').find_next('span').get_text(strip=True) if mobile_detail else None
            
            # Extract Email
            email_detail = member.find('p', text='E-mail.		    ')  # Find the 'E-mail.' label
            email = email_detail.find_next('span').find_next('span').get_text(strip=True) if email_detail else None
            
            # Extract Product
            product_detail = member.find('p',  text=re.compile(r'\bProd\b'))  # Find the 'prod.' label
            product = product_detail.find_next_sibling(text=':').find_next_sibling('span').find_next_sibling('span').get_text(strip=True) if product_detail else None

            # Append the data
            data_rows.append([name, work, address, executive, phone, mobile, email, product])
        except AttributeError:
            # If any element is not found, handle it gracefully
            continue
    
    return data_rows

# List to hold data from all pages
all_data = []

# Open the initial page
driver.get('https://yourwebsite/thespecific/page')  # Replace with the actual URL

while True: 
    # Scrape the data from the current page
    page_data = scrape_current_page(driver)
    all_data.extend(page_data)

    try:
        # Find the 'Next' button and click it
        next_button = driver.find_element(By.LINK_TEXT, 'Next')  # Adjust the method to locate the 'Next' button
        next_button.click()
        time.sleep(2)  # Wait for the next page to load
    except:
        # If there is no 'Next' button, break the loop
        break

# Close the browser
driver.quit()

# Define the column names
columns = ['Name', 'Work', 'Address', 'Executive', 'Phone', 'Mobile', 'Email', 'Product']

# Convert the list of data into a DataFrame
df = pd.DataFrame(all_data, columns=columns)

# Specify the directory where you want to save the file
save_directory = r'D:\New\new\new'  # Replace with your desired path

# Ensure the directory exists
os.makedirs(save_directory, exist_ok=True)

# Full path to the output Excel file
output_file = os.path.join(save_directory, 'scraped_data.xlsx')

# Write the DataFrame to an Excel file at the specified location
df.to_excel(output_file, index=False)

print(f"Data has been scraped and saved to '{output_file}'")
