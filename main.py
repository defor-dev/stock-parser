from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup
from sheets import *

CREDENTIALS = './etc/CREDENTIALS.json'
SS_ID = '1HTJeZDdPY_vesSLugBi2dD4qzeOpNLKugjqoXQDcYEU'
sheetsId = open('./etc/sheets-id.txt', 'r').readline().split(',')

cme_ss = Spreadsheet(CREDENTIALS)
cme_ss.setSpreadsheetById(SS_ID)
cme_ss.prepare_deleteSheet(int(sheetsId[0]))
cme_ss.runPrepared()
cmeSheetId = cme_ss.addSheet('cme', 50000, 3)

theice_ss = Spreadsheet(CREDENTIALS)
theice_ss.setSpreadsheetById(SS_ID)
theice_ss.prepare_deleteSheet(int(sheetsId[1]))
theice_ss.runPrepared()
theiceSheetId = theice_ss.addSheet('theice', 50000, 3)

open('./etc/sheets-id.txt', 'w').write(str(cmeSheetId) + ',' + str(theiceSheetId))

driver = webdriver.Chrome('./etc/chromedriver')

# First parser
txt = open('./etc/first-links.txt')
links = txt.read().split('\n')[0:-1]
rowCount = 1
for link in links:
    driver.get(link)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//li")))
    if (driver.find_elements_by_xpath("//li")[51].text != 'Data'): continue
    driver.find_elements_by_xpath("//li")[51].click()

    try:
        WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, "//table")))
    except TimeoutException: continue

    page = driver.page_source
    soup = BeautifulSoup(page, "html.parser")

    if (soup.find('div', attrs={'class':'alert-danger'})):
        print('alert')
        continue

    title = soup.findAll('h1')[1].text
    table = soup.find('table', attrs={'class':'table-bigdata'})
    theice_ss.prepare_mergeCells('A{}:C{}'.format(rowCount, rowCount))
    theice_ss.prepare_setValues('A{}:C{}'.format(rowCount, rowCount), [[title]])
    theice_ss.prepare_setCellsFormats('A{}:C{}'.format(rowCount, rowCount), [[{"textFormat": {"bold": True},
                                        "horizontalAlignment": "CENTER"}]],
                                        fields = "userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment")
    rowCount += 1
    theice_ss.prepare_setValues('A{}'.format(rowCount), [['Month']])
    theice_ss.prepare_setValues('B{}'.format(rowCount), [['Last']])
    theice_ss.prepare_setValues('C{}'.format(rowCount), [['Volume']])
    rowCount += 1

    rows = []
    try:
        rows = table.findAll('tbody')
    except AttributeError:
        continue

    lastLast = 0.0
    for row in rows:
        row = row.find('tr')
        month = row.find('a').text
        last = row.findAll('td')[1].text
        volume = row.findAll('td')[4].text

        if (last.strip() not in ['', 'Last', None] and lastLast not in ['', 'Last', None]):
            last = float(last)
            print(last, lastLast)
            try:
                if (last < lastLast):
                    theice_ss.prepare_setCellsFormat("B{}:B{}".format(rowCount-1, rowCount-1),
                                                                {"backgroundColor": htmlColorToJSON("#92D050")},
                                                                fields = "userEnteredFormat.backgroundColor")
            except TypeError: pass

        theice_ss.prepare_setValues('A{}'.format(rowCount), [[month]])
        theice_ss.prepare_setValues('B{}'.format(rowCount), [[last]])
        theice_ss.prepare_setValues('C{}'.format(rowCount), [[volume]])

        theice_ss.prepare_setCellsFormat("A{}:A{}".format(rowCount, rowCount), {"horizontalAlignment": "LEFT"}, fields='userEnteredFormat.horizontalAlignment')
        theice_ss.prepare_setCellsFormat("B{}:C{}".format(rowCount, rowCount), {"horizontalAlignment": "RIGHT"}, fields='userEnteredFormat.horizontalAlignment')
        
        rowCount += 1
        lastLast = last
    theice_ss.runPrepared()

# Second parser
txt = open('./etc/second-links.txt')
links = txt.read().split('\n')[0:-1]
rowCount = 1
for link in links:
    driver.get(link)
    try:
        WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, "quotesFuturesProductTable1")))
    except TimeoutException: continue
    page = driver.page_source
    soup = BeautifulSoup(page, "html.parser")

    title = soup.find('h1', {'class': 'title'}).text
    table = soup.find('table', {'id': 'quotesFuturesProductTable1'})

    cme_ss.prepare_mergeCells('A{}:C{}'.format(rowCount, rowCount))
    cme_ss.prepare_setValues('A{}:C{}'.format(rowCount, rowCount), [[title]])
    cme_ss.prepare_setCellsFormats('A{}:C{}'.format(rowCount, rowCount), [[{"textFormat": {"bold": True},
                                        "horizontalAlignment": "CENTER"}]],
                                        fields = "userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment")
    rowCount += 1

    cme_ss.prepare_setValues('A{}'.format(rowCount), [['Month']])
    cme_ss.prepare_setValues('B{}'.format(rowCount), [['Last']])
    cme_ss.prepare_setValues('C{}'.format(rowCount), [['Volume']])
    rowCount += 1

    allColumns = table.find('thead').find('tr').findAll('th')
    allRows = table.find('tbody').findAll('tr')

    for index, column in enumerate(allColumns):
        if (column.text.strip() == 'Volume'): break

    lastLast = 0.0
    for row in allRows:
        print(row)
        if (row.findAll('td')[3].text.strip() not in ['-', '', 'Last']):
            print(row.findAll('td')[3].text)
            last = list(row.findAll('td')[3].text)
            # print(last)
            if last.count("'") != 0: last.remove("'")
            if last.count('A') != 0: last.remove('A')
            if last.count('B') != 0: last.remove('B')
            # print(last)
            last = ''.join(last)
            last = float(last)
            if (last < lastLast):
                cme_ss.prepare_setCellsFormat("B{}:B{}".format(rowCount-1, rowCount-1), {"backgroundColor": htmlColorToJSON("#92D050")}, fields = "userEnteredFormat.backgroundColor")
        
        cme_ss.prepare_setValues('A{}'.format(rowCount), [[row.find('th').text]])
        cme_ss.prepare_setValues('B{}'.format(rowCount), [[row.findAll('td')[3].text]])
        cme_ss.prepare_setValues('C{}'.format(rowCount), [[row.findAll('td')[index-1].text]])

        cme_ss.prepare_setCellsFormat("A{}:A{}".format(rowCount, rowCount), {"horizontalAlignment": "LEFT"}, fields='userEnteredFormat.horizontalAlignment')
        cme_ss.prepare_setCellsFormat("B{}:C{}".format(rowCount, rowCount), {"horizontalAlignment": "RIGHT"}, fields='userEnteredFormat.horizontalAlignment')

        lastLast = last
        rowCount += 1
    cme_ss.runPrepared()

driver.quit() 
