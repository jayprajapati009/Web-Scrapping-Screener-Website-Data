from bs4 import BeautifulSoup as bs
from numpy import append
import requests as re
import os
from selenium.webdriver import chrome
from selenium.webdriver import ChromeOptions

os.chdir(r'C:\Users\jaypr\Desktop\Tech Stack\VSCodes\Web Scrapping\StockScrapping')


def find_data():

    option = ChromeOptions()
    option.headless = True
    driver = chrome(
        executable_path=r'C:\Users\jaypr\Downloads\chromedriver_win32/chromedriver.exe')
    driver.get(
        'https://www.screener.in/company/540416/consolidated/#balance-sheet')

    html_text = re.get(
        'https://www.screener.in/company/540416/consolidated/#balance-sheet').text
    # html_text = re.get(
    #     'https://www.screener.in/company/540455/').text

    soup = bs(driver, 'lxml')

    cashEq = soup.find(
        'section', id='balance-sheet', class_='card card-large').children
    print(cashEq)
    # val_i = cashEq.find('Borrowings')
    # othli_i = cashEq.find('Other Liabilities')
    # cashEq_list = cashEq[val_i + 10: othli_i].strip().split('\n')
    driver.quit()


find_data()
