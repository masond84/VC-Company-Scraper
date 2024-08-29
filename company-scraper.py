# Import Modules
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from openpyxl import load_workbook
from urllib.parse import urljoin, urlparse
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

def load_test_data(file_path):
    # Load the workbook and select the correct sheet
    workbook = load_workbook(filename=file_path)
    sheet = workbook.active  # or specify the sheet name if you know it

    # Dictionary to store the company names and URLs
    companies_url_dict = {}

    # Iterate through the rows to extract company names and URLs
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=2):
        company = row[0].value  # Assuming the company name is in the first column (A)
        url = row[1].value  # Assuming the URL is in the second column (B)
        if company and url: 
            companies_url_dict[company] = url

    return companies_url_dict

"""
Function to extract all internal links from a single webpage
"""
def extract_internal_links(driver, base_url):
    internal_links = set()
    elements = driver.find_elements(By.TAG_NAME, 'a')

    for element in elements:
        href = element.get_attribute('href')
        if href:
            normalized_href = urljoin(base_url, href)
            if urlparse(normalized_href).netloc == urlparse(base_url).netloc:
                internal_links.add(normalized_href)
    
    return internal_links

"""
Function to recursively vist all the internal links on a site
"""
def traverse_site(driver, base_url):
    visited_links = set()
    links_to_visit = {base_url}
    search_results = {}

    while links_to_visit:
        current_link = links_to_visit.pop()
        if current_link in visited_links:
            continue # Skip already visited links

        print(f"Visiting: {current_link}")
        driver.get(current_link)
        visited_links.add(current_link)

        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'title')))
            
            # Search for keywords in the current page
            page_search_results = search_keywords_in_elements(driver)
            search_results[current_link] = page_search_results

            # Extract Internal links and update the traversal list
            new_links = extract_internal_links(driver, base_url)
            links_to_visit.update(new_links - visited_links) # Only add unvisited links

        except TimeoutException:
            print(f"Timeout while accessing {current_link}")
        except Exception as e:
            print(f"Error while traversing {current_link}: {e}")
        
    return visited_links, search_results
        

def open_company_sites(driver, company_url_dict):
    for company, url in company_url_dict.items():
        try:
            # Format the URl correctly
            formmated_url = format_url(url)

            print(f"Accessing {company}: {formmated_url}")
            driver.get(formmated_url)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'title')))

            # Call the traverse_site function to visit all internal pages
            visted_pages, search_results = traverse_site(driver, formmated_url)
            print(f"Total pages visited for {company}: {len(visted_pages)}")
            print(f"Search results for {company}: {search_results}")

        except TimeoutException:
            print(f"Timeout while accessing {company} at {url}")
        except Exception as e:
            print(f"Failed to access {company} at {url}: {e}")

"""
Function to Search Keywords
- Logic to Scrape Websites: Load the page, extract links, and recursively crawl them.
- Search for Keywords: On each page, search for the given list of keywords.
"""
def search_keywords(page_content):
    keywords = [
        'Private Equity',
        'Capital Markets',
        'Leverage Finance', 
        'Investment Banking',
        "Investment Firm",
        'b2b saas',
        'pre-seed',
        'Southeast',
        'latin',
        'Hispanic',
        'Florida'
    ]
    results = {}

    for keyword in keywords:
        if keyword.lower() in page_content:
            results[keyword] = "Yes"
        else:
            results[keyword] = "No"

    return results

def search_keywords_in_elements(driver):
    page_content = driver.page_source.lower()
    results = search_keywords(page_content)

    elements_results = {}
    for keyword in results:
        elements = driver.find_elements(By.XPATH, f"//*[contains(text(), '{keyword.lower()}')]")
        if elements:
            elements_results[keyword] = [element.get_attribute('outerHTML') for element in elements]
        else:
            elements_results[keyword] = []
    
    return elements_results
    
"""
Main Scraper Function: The main function will load data, initialize the driver, iterate over each company URL, and use the keyword search function.
"""
def main_scraper(file_path):
    # Load the data from the excel file
    company_url_dict = load_test_data(file_path)

    driver = init_driver()

    # Open each company's website
    open_company_sites(driver, company_url_dict)

    driver.quit()

#### MAIN CODE ####
#file_path = r"C:\Users\dmaso\OneDrive\Documents\002 Projects\002 Freelance\1 - Javier Gutierrez\Scraper00\data\Aug 28 - VCs.xlsx"
file_path = r"C:\Users\dmaso\OneDrive\Documents\002 Projects\002 Freelance\1 - Javier Gutierrez\Scraper00\data\Test-Data.xlsx"
"""
# Test the load_excel_file function
companies_url_dict = load_excel_data(file_path)

print(f"Companies: {companies_url_dict}")
"""

main_scraper(file_path)