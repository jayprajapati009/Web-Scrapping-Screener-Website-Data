from bs4 import BeautifulSoup as bs
from numpy import append
import requests
import xlsxwriter
import os

# Setting up the directory to save the excel file in the same folder.
os.chdir(r'C:\Users\jaypr\Desktop\Tech Stack\VSCodes\Web Scrapping\StockScrapping')


def find_data():

    # To fetch the html data from the website
    html_text = requests.get(
        'https://www.screener.in/company/540416/consolidated/#balance-sheet').text
    # html_text = re.get(
    #     'https://www.screener.in/company/540455/').text

    # Parsing the data using lxml Parser and Beautiful Soup Library
    soup = bs(html_text, 'lxml')

    ##### Company Name #####
    Company = soup.find(
        'h1', class_='margin-0').text

    ##### Balance Sheet Months #####
    borrowings_dates_list = soup.find(
        'section', id='balance-sheet', class_='card card-large').table.thead.tr.text.strip().split('\n')

    ##### Borrowings or Debts - Balance Sheet #####
    borrowings_values = soup.find(
        'section', id='balance-sheet', class_='card card-large').text
    val_i = borrowings_values.find('Borrowings')
    othli_i = borrowings_values.find('Other Liabilities')
    borrowings_values_list = borrowings_values[val_i +
                                               10: othli_i].strip().split('\n')

    ##### Share Capital - Balance Sheet #####
    shareCap = soup.find(
        'section', id='balance-sheet', class_='card card-large').text
    shcap_i = shareCap.find('Share Capital')
    Res = shareCap.find('Reserves')
    shareCap_list = shareCap[shcap_i+15: Res].strip().split('\n')

    ##### Reserves - Balance Sheet #####
    Reserves = soup.find(
        'section', id='balance-sheet', class_='card card-large').text
    shcap_i = Reserves.find('Reserves')
    Res = Reserves.find('Borrowings')
    Reserves_list = Reserves[shcap_i+15: Res].strip().split('\n')

    ##### Profit and Loss Months #####
    pldates = soup.find(
        'section', id='profit-loss', class_='card card-large').text
    # print(pldates)
    if 'Standalone' in pldates:
        std_i = pldates.find('Standalone')
    else:
        std_i = pldates.find('Crores')
    sales_i = pldates.find('Sales\xa0+')
    pldates_list = pldates[std_i+25: sales_i].strip().split(
        '\n')

    ##### Profit Before Tax (PBT) - Profit and Loss #####
    pbt = soup.find(
        'section', id='profit-loss', class_='card card-large').text
    pbt_i = pbt.find('Profit before tax')
    tax_i = pbt.find('Tax %')
    pbt_list = pbt[pbt_i+25: tax_i].strip().split('\n')

    ##### Net Profit (PAT) - Profit and Loss #####
    pat = soup.find(
        'section', id='profit-loss', class_='card card-large').text
    pat_i = pat.find('Net Profit')
    tax_i = pat.find('EPS')
    pat_list = pat[pat_i+15: tax_i].strip().split('\n')

    ##### Sales - Profit and Loss #####
    sales = soup.find(
        'section', id='profit-loss', class_='card card-large').text
    sales_i = sales.find('Sales\xa0+')
    tax_i = sales.find('Expenses\xa0+')
    sales_list = sales[sales_i+10: tax_i].strip().split('\n')

    ##### Other Income - Profit and Loss #####
    othInc = soup.find(
        'section', id='profit-loss', class_='card card-large').text
    othInc_i = othInc.find('Other')
    tax_i = othInc.find('Interest')
    othInc_list = othInc[othInc_i+15: tax_i].strip().split('\n')

    # Create a workbook and add a worksheet
    workbook = xlsxwriter.Workbook(f'{Company}_Data.xlsx')
    worksheet = workbook.add_worksheet()

    # Borrowings Coloum (B) width
    worksheet.set_column('B:B', 20)

    # Borrowings Coloum (J) width
    worksheet.set_column('J:J', 20)

    # Defined Formats
    bold = workbook.add_format({'bold': True})
    cen = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
    title = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'bold': True})
    merge_format = workbook.add_format({
        'bold': 1,
        'align': 'center',
        'valign': 'vcenter'})

    # Headings Declarations
    worksheet.merge_range('A1:J1', Company, merge_format)
    worksheet.merge_range('A2:E2', "Balance Sheet", merge_format)
    worksheet.merge_range('F2:K2', "Profit and Loss", merge_format)
    worksheet.merge_range('A3:A4', 'Month', merge_format)
    worksheet.merge_range('B3:B4', 'Borrowings/Debt', merge_format)
    worksheet.merge_range('C3:E3', 'Shareholders Fund', merge_format)
    worksheet.merge_range('F3:F4', 'Month', merge_format)
    worksheet.merge_range('G3:G4', 'PBT', merge_format)
    worksheet.merge_range('H3:H4', 'PAT', merge_format)
    worksheet.merge_range('I3:K3', 'Total Revenue', merge_format)

    worksheet.write(3, 2, "Share Capital", title)
    worksheet.write(3, 3, "Reserves", title)
    worksheet.write(3, 4, "Total", title)
    worksheet.write(3, 8, "Sales", title)
    worksheet.write(3, 9, "Other Income", title)
    worksheet.write(3, 10, "Total", title)

    row, col = 4, 0
    for item in borrowings_dates_list:
        worksheet.write(row, col, item, cen)
        row += 1

    row, col = 4, 1
    for item1 in borrowings_values_list:
        worksheet.write(row, col, item1, cen)
        row += 1

    row, col = 4, 2
    for item1 in shareCap_list:
        worksheet.write(row, col, item1, cen)
        row += 1

    row, col = 4, 3
    for item1 in Reserves_list:
        worksheet.write(row, col, item1, cen)
        row += 1

    row, col = 4, 5
    for item1 in pldates_list:
        worksheet.write(row, col, item1, cen)
        row += 1

    row, col = 4, 6
    for item1 in pbt_list:
        worksheet.write(row, col, item1, cen)
        row += 1

    row, col = 4, 7
    for item1 in pat_list:
        worksheet.write(row, col, item1, cen)
        row += 1
    
    row, col = 4, 8
    for item1 in sales_list:
        worksheet.write(row, col, item1, cen)
        row += 1

    row, col = 4, 9
    for item1 in othInc_list:
        worksheet.write(row, col, item1, cen)
        row += 1

    workbook.close()


if __name__ == '__main__':
    find_data()
