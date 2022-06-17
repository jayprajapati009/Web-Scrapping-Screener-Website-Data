from gettext import find
import openpyxl
import os
import xlsxwriter
import requests
from numpy import append
from bs4 import BeautifulSoup as bs
from typing import final
from time import process_time_ns
from ast import Continue




# Setting up the directory to save the excel file in the same folder.
os.chdir(r'C:\Users\jaypr\Desktop\Tech Stack\VSCodes\Web Scrapping\StockScrapping\Scrapping Screener Website Data\Scraped Data')


def find_data(link):
    # To fetch the html data from the website
    html_text = requests.get(link).text

    # Parsing the data using lxml Parser and Beautiful Soup Library
    soup = bs(html_text, 'lxml')
    Company = soup.find(
        'h1', class_='margin-0').text

    print(Company)


if __name__ == '__main__':
    link = 'https://www.screener.in/company/SHAREINDIA/consolidated'
    find_data(link)
