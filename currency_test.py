import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
import time
import os

# Initialize WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Open the target URL
url = "https://www.alojamiento.io/property/habitaci%c3%b3n-familiar/BC-5546696"
driver.get(url)

# Wait for the currency dropdown to be present
try:
    currency_dropdown = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'js-currency-sort-footer'))
    )
except TimeoutException:
    print("Currency dropdown not found on the page. Please check the element ID.")
    driver.quit()
    exit()

WebDriverWait(driver, 30).until(
    EC.element_to_be_clickable((By.ID, "js-currency-sort-footer'"))
)

# List to store the test results
currency_test_results = []

# Helper function to log test results
def log_currency_test(currency, passed, comments):
    currency_test_results.append({
        "currency": currency,
        "passed/fail": "Passed" if passed else "Failed",
        "comments": comments
    })

# Function to select currency from the dropdown
def select_currency(currency_code):
    try:
        # Scroll the dropdown into view
        driver.execute_script("arguments[0].scrollIntoView(true);", currency_dropdown)
        
        # Click to open the dropdown
        currency_dropdown.click()

        # Wait for the currency options to load and be clickable
        time.sleep(2)  # Added delay to make sure options are fully visible

        # Find the currency option based on the currency code
        currency_option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//li[@data-currency-country='{currency_code}']"))
        )
    
        # Scroll the option into view and ensure it's clickable
        driver.execute_script("arguments[0].scrollIntoView(true);", currency_option)
        
        # Retry the click action in case of element interception
        try:
            driver.execute_script("arguments[0].click();", currency_option)
        except ElementClickInterceptedException:
            print(f"Element click intercepted for {currency_code}. Retrying...")
            time.sleep(1)
            driver.execute_script("arguments[0].click();", currency_option)

        # Wait for the currency to change and reflect on the page
        time.sleep(3)  # Add delay to allow prices to update

        # Ensure the property prices are updated
        property_prices = driver.find_elements(By.CLASS_NAME, "js-price-value")
        return all(currency_code in price.text for price in property_prices)

    except (TimeoutException, ElementClickInterceptedException) as e:
        log_currency_test(currency_code, False, f"Failed to select {currency_code}: {str(e)}")
        return False

# Function to test currency change
def test_currency_change():
    # Define the currency codes and expected symbols
    currencies = [
        ('US', 'USD'),
        ('CA', 'CAD'),
        ('BE', 'EUR'),
        ('IE', 'GBP'),
        ('AU', 'AUD'),
        ('SG', 'SGD'),
        ('AE', 'AED'),
        ('BD', 'BDT')
    ]

    for country_code, currency in currencies:
        success = select_currency(country_code)
        if success:
            log_currency_test(currency, True, f"Currency changed to {currency} and reflected on the property tiles.")
        else:
            log_currency_test(currency, False, f"Currency change to {currency} failed.")

    print("Currency test completed.")

# Run the currency change test
test_currency_change()

# Close the WebDriver
driver.quit()

# Save the results to the Excel report
def save_currency_test_results():
    # Ensure the 'reports' folder exists
    report_dir = 'reports'
    if not os.path.exists(report_dir):
        os.makedirs(report_dir)

    # Path to the Excel file
    excel_filename = os.path.join(report_dir, 'test_report.xlsx')

    # Open the Excel file or create a new one if it doesn't exist
    if not os.path.exists(excel_filename):
        workbook = openpyxl.Workbook()
        workbook.save(excel_filename)

    workbook = openpyxl.load_workbook(excel_filename)

    # Check if the "Currency Test" sheet already exists, otherwise create it
    if 'Currency Test' not in workbook.sheetnames:
        sheet = workbook.create_sheet(title='Currency Test')
        sheet.append(['Currency', 'Passed/Fail', 'Comments'])  # Add headers
    else:
        sheet = workbook['Currency Test']

    # Append the test results
    for result in currency_test_results:
        sheet.append([result["currency"], result["passed/fail"], result["comments"]])

    # Save the updated workbook
    workbook.save(excel_filename)
    print(f"Currency test results saved to {excel_filename}")

# Save the currency test results to the Excel sheet
save_currency_test_results()
