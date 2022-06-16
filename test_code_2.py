import openpyxl
import os
import xlsxwriter
import requests
from numpy import append
from bs4 import BeautifulSoup as bs
from typing import final
from time import process_time_ns
from ast import Continue

link = 'https://www.mcxindia.com/en/market-data/get-quote/FUTCOM/CRUDEOIL/17JUN2022'


# Setting up the directory to save the excel file in the same folder.
os.chdir(r'C:\Users\jaypr\Desktop\Tech Stack\VSCodes\Web Scrapping\StockScrapping\Scrapping Screener Website Data\Scraped Data')



# To fetch the html data from the website
html_text = requests.get(link).text

# Parsing the data using lxml Parser and Beautiful Soup Library
soup = bs(html_text, 'lxml')

print(soup)
