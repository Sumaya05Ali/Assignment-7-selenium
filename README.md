# Assignment-7-selenium

# Vacation Rental Home Page Automation Testing

This project automates the testing of a vacation rental details page to validate essential elements and functionality, focusing on SEO impacted test cases and data scraping. Results are saved in an Excel file for any identified issues.

## Features

The tests include:
- **H1 tag existence**: Check if the H1 tag is present on the page.
- **HTML tag sequence test**: Validate the presence and correct sequence of H1-H6 tags.
- **Image alt attribute validation**: Ensure all images have alt attributes.
- **URL status code test**: Verify that no URLs return a 404 status.
- **Currency filter validation**: Ensure property tiles reflect the selected currency.
- **Script data scraping**: Extract data such as CampaignID, SiteName, Browser, CountryCode, and IP, saving it to an Excel file.

## Tools and Libraries

- **Python**: Core programming language used.
- **Selenium**: Web automation tool to interact with the browser.
- **Pandas**: Used for handling and saving data into Excel files.
- **Openpyxl**: Library used for writing Excel files.
- **Requests**: For checking the status of URLs.

## Requirements

- **Browser**: Google Chrome or Firefox with WebDriver.
- **Python 3.x**
- **Required Python Libraries**:
  - selenium
  - pandas
  - requests
  - openpyxl
  - webdriver_manager

Install all dependencies using the following command:

```bash
pip install selenium pandas requests openpyxl webdriver_manager
```

## How to Run the Tests

1. Clone the repository to your local machine
   ```bash
    git clone https://github.com/Sumaya05Ali/Assignment-7-selenium.git
    cd Assignment-7-selenium-main 
   ```
   
2.  Install the required dependencies:
      ```bash
      pip install -r requirements.txt
     ```
      
3. Run the test cases:
   ```bash
      python test_cases.py
     ```
   
4. Run the currency test:
  ```bash
      python currency_test.py
   ```

5.  Run the data scraping script:
     ```bash
        python scrape_data_test.py
    ```
## Output
- An Excel report is generated in the reports folder named test_report.xlsx.
- The main sheet records the results of the SEO-related test cases (H1 tags, alt attributes, etc.), while the second sheet contains URL status checks.
- Data scraped from the page scripts (CampaignID, SiteName, etc.) is added to a separate sheet named Scraped Data and also currency filtering test report is added to another sheet named Currency Test.

## Test Details
- H1 Tag Existence: Checks if the H1 tag is present on the page. A missing H1 tag is reported as a failure.
- HTML Tag Sequence Test: Validates that all heading tags from H1 to H6 are present and follow the correct sequence. Any missing tag is reported.
- Image Alt Attribute Test: Ensures all images on the page have alt attributes. Missing attributes are flagged as failures.
- URL Status Code Test: Checks all URLs on the page for their status. If a URL returns a 404 (Not Found), it is reported as a failure.
- Currency Filter Test: Validates that the currency changes on the property tiles when selecting different currencies.
- Data Scraping from Scripts: Extracts specific data (CampaignID, SiteName, etc.) from the page’s script tags and saves it to the Excel file.

## License
This project is licensed under the MIT License.
This README outlines the purpose of the project, how to set it up, how to run it, and what the test cases do. Make sure to update the URLs and your repository link as needed.

