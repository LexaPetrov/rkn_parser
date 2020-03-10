from bs4 import BeautifulSoup
import requests as req
import pandas as pd
pd.io.formats.format.header_style = None

user_agent = ('Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:50.0) '
              'Gecko/20100101 Firefox/50.0')

url = 'http://rkn.gov.ru/communication/register/license/'

# TODO: вывод всех 5 таблиц в 1 эксель файле
#       при выводе форматирование ширины ячеек
#       наименование лицензиата должно быть гиперссылкой, нужно вытаскивать href 
#       запись в эксель ячейки лицензиата как гиперссылка с сопоставлением элементов массива hrefs и текста ячеек
#       запаковать скрипт в исполняемый файл, чтобы запускался без питона на компьютере

# Вывести в файл весь документ
#def save__to__file(resp, path):
#    f = open(path, 'w')
#    soup = BeautifulSoup(resp.text, 'lxml')
#    soup = soup.prettify()
#    f.write(soup)
#    f.close()

def excel__writer(table, path):
    writer = pd.ExcelWriter(path, engine='xlsxwriter')
    table.to_excel(writer, sheet_name='Sheet1', index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    worksheet.set_column(0, 0, 18)
    worksheet.set_column(1, 1, 80)
    worksheet.set_column(2, 2, 18)
    worksheet.set_column(3, 3, 25)
    worksheet.set_column(4, 4, 40)

    writer.save()




# Вывести в файл только таблицы
def save__to__file(resp, path):
    f = open(path, 'w')
    soup = BeautifulSoup(resp.text, 'lxml')
    soup = soup.find("table", id="ResList1")
    if soup != None:
        f.write(str(soup))
    else:
        f.close()

# сохранить в excel
def save__to__excel(resp, path):
    soup = BeautifulSoup(resp.text, 'lxml')
    soup = soup.find("table", id="ResList1")
    href = soup.find_all(href=True)
    hrefs = []
    for a in href:
        hrefs.append(a['href'])
    headers = soup.find("thead")
    tr = soup.find_all('tr')
    html = '<table>' + str(headers) + '</table>'
    t = pd.read_html(str(soup))
    table = t[0].drop("Unnamed: 5", axis=1)
    excel__writer(table, path);
    #table.to_excel(path, index=False)
    #print(table)
    return table
 
# Тут все нужные таблицы 
# for i in range (0, 5):
#     save__to__file(
#     req.post(
#         url + 'p' + str(500 * i) + '/?all=1',
#         headers={'User-Agent':user_agent},
#         params={'SERVICE_ID': 12}
#     ),
#     'table__page' + str(i+1) + '.html'
# )


# Тут 5 экселей
for i in range (5):
    save__to__excel(
        req.post(
        url + 'p' + str(500 * i) + '/?all=1',
        headers={'User-Agent':user_agent},
        params={'SERVICE_ID': 12}
    ),
    'table' + str(i + 1) + '.xlsx'
)

def excel__writer(table, path):
    writer = pd.ExcelWriter(path, engine='xlsxwriter')
    for sheetname, df in table.items():
        df.to_excel(writer, sheet_name=sheetname)
        worksheet = writer.sheets[sheetname]  # pull worksheet object
        for idx, col in enumerate(df):  # loop through all columns
            series = df[col]
            max_len = max((series.astype(str).map(len).max(), len(str(series.name)))) + 1  # adding a little extra space
            worksheet.set_column(idx, idx, max_len)  # set column width
    writer.save()




# https://python-scripts.com/beautifulsoup-html-parsing
# https://www.softwaretestinghelp.com/selenium-webdriver-commands-selenium-tutorial-17/
# https://stackoverflow.com/questions/35831241/converting-html-to-excel-in-python
# https://www.geeksforgeeks.org/python-convert-an-html-table-into-excel/
# https://stackoverflow.com/questions/31820069/add-hyperlink-to-excel-sheet-created-by-pandas-dataframe-to-excel-method