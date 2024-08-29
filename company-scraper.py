# Import Modules
from selenium import webdriver
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
import time

"""
Function to load Excel File and Extract Data
"""
def load_excel_data(file_path):
    workbook = load_workbook(filename=file_path)
    sheet = workbook["Pre-Seed US"]

    companies = []
    urls = []

    for row in sheet.iter_rows(min_row=6, max_row=sheet.max_row, min_col=3, max_col=7):
        company = row[0].value
        url = row[4].value
        if company  and url: 
            companies.append(company)
            urls.append(url)

    return companies, urls

# Test the load_excel_file function
"""
file_path = r"C:\Users\dmaso\OneDrive\Documents\002 Projects\002 Freelance\1 - Javier Gutierrez\Scraper00\data\Aug 28 - VCs.xlsx"
companies, urls = load_excel_data(file_path)

print(f"Companies: {companies}")
print(f"URLs: {urls}")
"""

"""
Function to initialize WebDriver
- Setup Selenium WebDriver: Initialize the Selenium WebDriver.
"""

"""
Function to Scrape and Search Keywords
- Logic to Scrape Websites: Load the page, extract links, and recursively crawl them.
- Search for Keywords: On each page, search for the given list of keywords.
"""

"""
Main Scraper Function: The main function will load data, initialize the driver, iterate over each company URL, and use the keyword search function.
"""