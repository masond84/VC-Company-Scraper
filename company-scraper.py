# Import Modules
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from openpyxl import load_workbook
import time

def format_url(url):
    # Ensure the URL has a proper scheme
    if not url.startswith(('http://', 'https://')):
        url = "http://" + url
    return url

"""
Function to initialize WebDriver
- Setup Selenium WebDriver: Initialize the Selenium WebDriver.
"""
def init_driver():
    # Setup Chrome Options
    options = webdriver.ChromeOptions()
    #options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_argument('log-level=3')

    # Initialize the Webdriver
    driver = webdriver.Chrome(options=options)
    return driver

"""
Function to load Excel File and Extract Data
"""
def load_excel_data(file_path):
    workbook = load_workbook(filename=file_path)
    sheet = workbook["Pre-Seed US"]

    companies_url_dict = {}

    for row in sheet.iter_rows(min_row=6, max_row=sheet.max_row, min_col=3, max_col=7):
        company = row[0].value
        url = row[4].value
        if company  and url: 
            companies_url_dict[company] = url

    return companies_url_dict

def open_company_sites(driver, company_url_dict):
    for company, url in company_url_dict.items():
        try:
            print(f"Attempting To Format: {url}")
            # Format the URl correctly
            formmated_url = format_url(url)

            print(f"Accessing {company}: {formmated_url}")
            driver.get(formmated_url)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'title')))

            # Placeholder: This is where you'd call the function to search for keywords
            # For now, just print the page title as a simple verification step
            print(f"Title: {driver.title}")
        except TimeoutException:
            print(f"Timeout while accessing {company} at {url}")
        except Exception as e:
            print(f"Failed to access {company} at {url}: {e}")

"""
Function to Search Keywords
- Logic to Scrape Websites: Load the page, extract links, and recursively crawl them.
- Search for Keywords: On each page, search for the given list of keywords.
"""
def search_keywords(driver, url, keywords):
    try:
        # Load the webpage
        driver.get(url)
        time.sleep(2)

        # Get the page content
        page_content = driver.page_source.lower() # Convert content to lowercase

        results = {}
        for keyword in keywords:
            if  keyword in keywords:
                results[keyword] = "Yes"
            else:
                results[keyword] = "No"
        
        return results
    
    except Exception as e:
        print(f'Error scraping {url}: {e}')
        return None
    
"""
Main Scraper Function: The main function will load data, initialize the driver, iterate over each company URL, and use the keyword search function.
"""
def main_scraper(file_path):
    # Load the data from the excel file
    company_url_dict = load_excel_data(file_path)

    driver = init_driver()

    # Open each company's website
    open_company_sites(driver, company_url_dict)

    driver.quit()

#### MAIN CODE ####
file_path = r"C:\Users\dmaso\OneDrive\Documents\002 Projects\002 Freelance\1 - Javier Gutierrez\Scraper00\data\Aug 28 - VCs.xlsx"
"""
# Test the load_excel_file function
companies_url_dict = load_excel_data(file_path)

print(f"Companies: {companies_url_dict}")
"""

main_scraper(file_path)