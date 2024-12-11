import requests
import openpyxl
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
from openpyxl import Workbook




# Function to select currency and verify price changes
def select_currency_and_verify():
    initial_prices = [price.text for price in driver.find_elements(By.CLASS_NAME, 'js-price-value')]

    for currency in currency_options:
       
        print(f"Selecting currency: {currency["symbol"]}")
        
        # Click the currency option
        dropdown.click()
        # Locate the option based on the country
        option = next(
            (opt for opt in options if opt.get_attribute("data-currency-country") == currency["country"]), None
        )


        # Scroll the option into view and click
        driver.execute_script("arguments[0].scrollIntoView();", option)
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(option)).click()


        # Capture the new prices
        new_prices = [price.text for price in driver.find_elements(By.CLASS_NAME, 'js-price-value')]
        print(initial_prices)
        print(new_prices)


        # Check if prices changed
        if new_prices != initial_prices:
            print(f"Currency change detected for {currency["symbol"]}. Prices updated.")
            currency_data.append({"currency_option": {currency["symbol"]}, "price_changed": "YES"})
        else:
            print(f"Currency change for {currency["symbol"]} did not update the prices.")
            currency_data.append({"currency_option": {currency["symbol"]}, "price_changed": "NO"})

        time.sleep(1)


# Initialize WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Define the correct URL of the property
url = "https://www.alojamiento.io/property/habitaci%c3%b3n-familiar/BC-5546696"
driver.get(url)

WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.TAG_NAME, "body"))
)
    
# Scroll to the bottom of the page multiple times to ensure lazy-loaded content is visible
for _ in range(3):
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(2)

dropdown = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "js-currency-sort-footer"))
)


# Find the currency options (assuming the list items represent currencies)
options = dropdown.find_elements(By.CSS_SELECTOR, ".select-ul > li")
dropdown.click()


currency_options = []
for option in options:
    data_country = option.get_attribute("data-currency-country")
    currency_element = option.find_element(By.CSS_SELECTOR, ".option > p")
    currency_symbol = currency_element.text.split(" ")[0].strip()
    currency_options.append({"country": data_country, "symbol": currency_symbol})


# Initialize the list to store data for the DataFrame
currency_data = []

# Call the function to select currencies and verify price changes
# select_currency_and_verify()
initial_prices = [price.text for price in driver.find_elements(By.CLASS_NAME, 'js-price-value')]

for currency in currency_options:
    
    print(f"Selecting currency: {currency["symbol"]}")
    

    # Locate the option based on the country
    option = next(
        (opt for opt in options if opt.get_attribute("data-currency-country") == currency["country"]), None
    )


    # Scroll the option into view and click
    driver.execute_script("arguments[0].scrollIntoView();", option)
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(option)).click()



    # Capture the new prices
    new_prices = [price.text for price in driver.find_elements(By.CLASS_NAME, 'js-price-value')]
    print(initial_prices)
    print(new_prices)


    # Check if prices changed
    if new_prices != initial_prices:
        print(f"Currency change detected for {currency["symbol"]}. Prices updated.")
        currency_data.append({"currency_option": {currency["symbol"]}, "price_changed": "YES"})
    else:
        print(f"Currency change for {currency["symbol"]} did not update the prices.")
        currency_data.append({"currency_option": {currency["symbol"]}, "price_changed": "NO"})

    # Reset to the dropdown for the next selection
    dropdown.click()
    time.sleep(1)

print("function executed")

# Convert the data to a Pandas DataFrame
df = pd.DataFrame(currency_data)

# Save the DataFrame to an Excel file
excel_filename = 'currency_change_report.xlsx'
df.to_excel(excel_filename, index=False)

# Close the WebDriver
driver.quit()

print(f"Currency change report saved to {excel_filename}")
