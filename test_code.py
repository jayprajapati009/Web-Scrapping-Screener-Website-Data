import openpyxl

# os.chdir(r'C:\Users\jaypr\Desktop\Tech Stack\VSCodes\Web Scrapping\StockScrapping')

location = (r"C:\Users\jaypr\Desktop\Tech Stack\VSCodes\Web Scrapping\StockScrapping\Scrapping Screener Website Data\Stock Company List.xlsx")

wb = openpyxl.load_workbook(location)
sheet = wb.active
link = []
for i in range(1, 28):
    link.append(sheet.cell(row=i, column=1).value)

