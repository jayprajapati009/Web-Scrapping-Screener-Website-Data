from bs4 import BeautifulSoup as bs
from numpy import append
import requests as re
import os

os.chdir(r'C:\Users\jaypr\Desktop\Tech Stack\VSCodes\Web Scrapping\StockScrapping')


def find_data():

    html_text = re.get(
        'https://www.screener.in/company/540416/consolidated/#balance-sheet').text
    # html_text = re.get(
    #     'https://www.screener.in/company/540455/').text

    soup = bs(html_text, 'lxml')
    borrowings_dates = soup.find(
        'section', id='balance-sheet', class_='card card-large').table.thead.tr.text.strip().split('\n')
    print(borrowings_dates)


find_data()
