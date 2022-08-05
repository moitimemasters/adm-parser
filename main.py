import requests
from bs4 import BeautifulSoup as BS
from openpyxl import Workbook
import sys

URL, filename, *_ = sys.argv[1:]

page = requests.get(URL)
page.encoding = "utf-8"
page = page.text

bs = BS(page, "lxml")

def parse_adm_table(table):
    headers = [th.text for th in table.thead.tr.find_all("th")]
    result = {header:[] for header in headers}
    result["Поданная программа"] = []
    for row in table.tbody.find_all("tr"):
        tds = row.find_all("td")
        for idx, td in enumerate(tds[:-1]):
            result[headers[idx]].append(td.text)
        result[headers[-1]].append([a.text for a in tds[-1].find_all("a", recursive=True)])
        app = tds[-1].find("b", recursive=True)
        app = app.text if app else "-"
        result["Поданная программа"].append(app)
    return result, headers

table = bs.find_all("table")[-1]
dct, headers = parse_adm_table(table)

def export_to_excel(table, filename="adm.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Конкурсный список"
    for idx, header in enumerate(list(table.keys())[:-2]):
        ws.cell(row=1, column=idx + 1, value=header)
    
    headers = list(table.keys())[:-2]

    for row in range(len(table[headers[0]])):
        for column in range(len(headers)):
            ws.cell(row=row + 2, column=column + 1, value=table[headers[column]][row])
    
    others = wb.create_sheet("Другие ОП")
    others.cell(row=1, column=1, value="СНИЛС")
    others.cell(row=1, column=2, value="Программа")
    total = 1
    for idx, snils in enumerate(table["СНИЛС"]):
        for program in table["Другие ОП"][idx]:
            total += 1
            others.cell(row=total, column=1, value=snils)
            others.cell(row=total, column=2, value=program)
    applied = wb.create_sheet("Поданные документы")
    applied.cell(row=1, column=1, value="СНИЛС")
    applied.cell(row=1, column=2, value="Программа")
    for idx, snils in enumerate(table["СНИЛС"]):
        others.cell(row=idx + 2, column=1, value=snils)
        others.cell(row=idx + 2, column=2, value=table["Поданная программа"][idx])
        
    wb.save(filename)


export_to_excel(dct, filename)
