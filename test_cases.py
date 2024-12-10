import requests
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from openpyxl import Workbook
import os

# Initialize WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Define the correct URL of the property
url = "https://www.alojamiento.io/property/habitaci%c3%b3n-familiar/BC-5546696"
driver.get(url)

# List to store the test results
results = []
url_status_results = []

# Helper function to log test results
def log_result(page_url, testcase, passed, comments):
    results.append({
        "page_url": page_url,
        "testcase": testcase,
        "passed/fail": "Passed" if passed else "Failed",
        "comments": comments
    })

# Function to log URL status in the test report
def log_url_status(url, status_code, passed):
    url_status_results.append({
        "url": url,
        "status_code": status_code,
        "passed/fail": "Passed" if passed else "Failed"
    })

# Test 1: H1 tag existence
try:
    h1_tag = driver.find_element(By.TAG_NAME, "h1")
    log_result(url, "H1 tag existence", True, "H1 tag found.")
except:
    log_result(url, "H1 tag existence", False, "H1 tag not found.")

# Test 2: HTML tag sequence (H1-H6)
def test_html_tag_sequence():
    passed = True
    for i in range(1, 7):
        try:
            driver.find_element(By.TAG_NAME, f"h{i}")
        except:
            passed = False
            log_result(url, f"HTML tag H{i} existence", False, f"H{i} tag is missing.")
    if passed:
        log_result(url, "HTML tag sequence test", True, "H1-H6 tags are correctly available.")

test_html_tag_sequence()

# Test 3: Image alt attribute test
images = driver.find_elements(By.TAG_NAME, "img")
for img in images:
    alt_text = img.get_attribute("alt")
    if not alt_text:
        log_result(url, "Image alt attribute test", False, f"Image missing alt attribute: {img.get_attribute('src')}")
    else:
        log_result(url, "Image alt attribute test", True, f"Alt attribute present: {alt_text}")

# Function to test URL status
def test_url_status():
    links = driver.find_elements(By.TAG_NAME, "a")
    for link in links:
        try:
            href = link.get_attribute("href")
            if href:
                # Checking the status code of the URL
                try:
                    response = requests.get(href)
                    status = response.status_code
                    if status == 404:
                        log_url_status(href, status, False)  # Log failed URLs (404)
                    else:
                        log_url_status(href, status, True)   # Passed status (200, etc.)
                except requests.exceptions.RequestException as e:
                    log_url_status(href, "Request Failed", False)
        except Exception as e:
            print(f"Error retrieving href: {str(e)}")
    
    print("All URL statuses have been checked.")

# Run the URL status test
test_url_status()

# Close the WebDriver
driver.quit()

# Create Excel report
# Ensure the 'reports' folder exists
report_dir = 'reports'
if not os.path.exists(report_dir):
    os.makedirs(report_dir)

# Path to save the Excel file in the 'reports' folder
excel_filename = os.path.join(report_dir, 'test_report.xlsx')

wb = Workbook()

# Create sheets
main_sheet = wb.active
main_sheet.title = 'Main Test Cases'
main_sheet.append(['page_url', 'testcase', 'passed/fail', 'comments'])

url_status_sheet = wb.create_sheet(title='URL Status Test')
url_status_sheet.append(['url', 'status_code', 'passed/fail'])

# Save the main test results to the Excel sheet
for result in results:
    main_sheet.append([result["page_url"], result["testcase"], result["passed/fail"], result["comments"]])

# Save the URL status test results to a separate sheet (both passed and failed)
for result in url_status_results:
    url_status_sheet.append([result["url"], result["status_code"], result["passed/fail"]])

# Save the Excel file inside the 'reports' folder
wb.save(excel_filename)
print(f"Test completed. Results saved to {excel_filename}")