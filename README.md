# Automated SEMrush CPC Volume Scraper

This Python script automates the process of scraping CPC (Cost Per Click) keyword volume data from the SEMrush platform for a specified list of keywords, dates, and databases. It utilizes Selenium for web scraping, Openpyxl for Excel file manipulation, and interacts with the SEMrush platform to retrieve CPC keyword volume metrics.

## Features:

### User Input:

The script prompts the user to input a list of keywords, select a target database from a predefined list, and provide a list of dates.

### Excel Workbook Creation:

The script generates an Excel workbook with one sheet named after the selected database, containing CPC keyword volume data for each specified date.

### Web Scraping:

Utilizes Selenium WebDriver to log in to the SEMrush platform, input user credentials, and scrape CPC keyword volume data for each keyword and date.

### Data Validation and Handling:

Handles potential errors such as NoSuchElementException during web scraping and retries the process up to three times before moving on.

### Output:

The final output is an Excel file named 'output.xlsx' containing CPC keyword volume data.

## Usage:

1. **Prerequisites:**
   - Ensure you have the required Python packages installed (`pandas`, `openpyxl`, `selenium`), and download the appropriate Selenium WebDriver for Chrome.

2. **Credentials:**
   - Provide your SEMrush login credentials in the script.

3. **Execution:**
   - Run the script, input the necessary information, and follow the instructions. 
   - Retrieve the final output Excel file ('output.xlsx') for CPC keyword volume analysis.

**Note:** This script assumes the availability and accessibility of the SEMrush platform, and any changes to the platform structure may affect its functionality.
