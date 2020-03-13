from bs4 import BeautifulSoup
import requests as req
import pandas as pd
from datetime import datetime
import time
pd.io.formats.format.header_style = None

user_agent = ('Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:50.0) '
              'Gecko/20100101 Firefox/50.0')

# user_agent = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.132 Safari/537.36'

url = 'http://rkn.gov.ru/communication/register/license/'

# TODO: 
#       запаковать скрипт в исполняемый файл, чтобы запускался без питона на компьютере

start = datetime.now()
print('start')

# def get__search__results(resp, col):
#     if resp.status_code == 200:
#         soup = BeautifulSoup(resp.text, 'lxml')
#         soup = soup.find_all('div', {'class': 'r'})
#         for div in soup:
#             return (div.find('cite').get_text())
#     return (f'http://google.com/search?q={col}')

def get__search__results(resp, col):
    print(resp.status_code)
    if resp.status_code == 200:
        soup = BeautifulSoup(resp.text, 'lxml')
        for g in soup.find_all('div', class_='r'):
            anchors = g.find_all('a')
            if anchors:
                link = anchors[0]['href']
                return link
    return (f'http://google.com/search?q={col}')


def get__list__org(resp, col):
    print(resp.text)
    soup = BeautifulSoup(resp.text, 'lxml')
    div = soup.find('div', class_='org_list')
    print(div)

def replace__text(text):
    if text.__contains__('Акционерное общество'):
        text = text.replace('Акционерное общество', 'АО')
        return text
    elif text.__contains__('Закрытое акционерное общество'):
        text = text.replace('Закрытое акционерное общество', 'ЗАО')
        return text
    elif text.__contains__('Индивидуальный предприниматель'):
        text = text.replace('Индивидуальный предприниматель', 'ИП')
        return text
    elif text.__contains__('Публичное акционерное общество'):
        text = text.replace('Публичное акционерное общество', 'ПАО')
        return text
    else:
        return text

def excel__writer(table, path):
    writer = pd.ExcelWriter(path, engine='xlsxwriter')
    table.to_excel(writer, sheet_name='Sheet1', index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    align_format = workbook.add_format({'align':'center'})
    for col_idx, col in enumerate(table.columns):
        series = table[col]
        max_len = max((series.astype(str).map(len).max(), len(str(series.name)))) + 5
        align_format = workbook.add_format({'align':'center'})
        if col == 'ИНН лицензиата':
            # специальный формат для столбца 'ИНН лицензиата' (10 цифр)
            cell_format = workbook.add_format({'num_format': '0' * 10, 'align':'center'})
            worksheet.set_column(col_idx, col_idx, max_len, cell_format)
        elif col == 'Наименование лицензиата':
            # специальный формат и ширина для столбца 'Наименование лицензиата'
            cell_format = workbook.get_default_url_format()
            worksheet.set_column(col_idx, col_idx, 80, cell_format)
            for row_idx, (id, val) in enumerate(zip(table['Номер лицензии'].values, table[col].values)):
                val = replace__text(val)
                worksheet.write_url(row_idx + 1, col_idx, url + f'?id={id}&all=1', string=val)
        elif col == 'Поиск в Google':
            cell_format = workbook.get_default_url_format()
            worksheet.set_column(col_idx, col_idx, 20, cell_format)
            for row_idx, (g) in enumerate(table['Поиск в Google']):
                g = g.replace(' ', '+')
                worksheet.write_url(row_idx + 1, col_idx, g, string='Найти')
        elif col == 'Поиск на List-Org':
            cell_format = workbook.get_default_url_format()
            worksheet.set_column(col_idx, col_idx, 20, cell_format)
            for row_idx, (q) in enumerate(table['Поиск на List-Org']):
                worksheet.write_url(row_idx + 1, col_idx, q, string='Найти по ИНН')
        else:
            worksheet.set_column(col_idx, col_idx, max_len, align_format)

    writer.save()
    print('Saved into', path)

# сохранить страницу в dataframe
def read__part__dataframe(resp, start_idx):
    soup = BeautifulSoup(resp.text, 'lxml')
    soup = soup.find("table", id="ResList1")
    t = pd.read_html(str(soup))
    table = t[0].drop("Unnamed: 5", axis=1)
    table.drop(0, axis=0, inplace=True)
    res = []
    res2 = []
    count = 0
    # for col in table['Наименование лицензиата']:
    #     count += 1
    #     print('Запрос номер', start_idx, count)
    #     res.append(get__search__results(
    #             req.get('http://google.com/search?q=' + col, 
    #                 headers={'User-Agent':user_agent}
    #             ),
    #             col
    #     ))

    # for col in table['ИНН лицензиата']:
    #     count += 1
    #     print('Запрос номер', start_idx, count)
    #     res.append(get__list__org(
    #             req.get('https://www.list-org.com/search?type=inn&val=' + str(col), 
    #                 headers={'User-Agent':user_agent}
    #             ),
    #             col
    #     ))

    for col in table['Наименование лицензиата']:
        res.append(f'https://www.google.com/search?q={col}')

    for col in table['ИНН лицензиата']:
        res2.append(f'https://www.list-org.com/search?type=inn&val={col}')
    
    table['Поиск в Google'] = res
    table['Поиск на List-Org'] = res2
    # table['Наименование лицензиата'] = table['Наименование лицензиата'].map({'Акционерное общество': 'АО'})
    index = pd.Index(range(start_idx, start_idx + table.shape[0]))
    table = table.set_index(index)
   
    return table

dfs = []
i = 0
while True:
    response = req.post(
        url + 'p' + str(500 * i) + '/?all=1',
        headers={'User-Agent':user_agent},
        params={'SERVICE_ID': 12},
        timeout=5
    )
    if 'Записей не найдено' in response.text:
        break

    dfs.append(read__part__dataframe(response, 500 * i))
    i += 1

full_df = pd.concat(dfs, axis=0)
excel__writer(full_df, 'table.xlsx')

print('Заняло времени - ', datetime.now() - start)

# https://python-scripts.com/beautifulsoup-html-parsing
# https://www.softwaretestinghelp.com/selenium-webdriver-commands-selenium-tutorial-17/
# https://stackoverflow.com/questions/35831241/converting-html-to-excel-in-python
# https://www.geeksforgeeks.org/python-convert-an-html-table-into-excel/
# https://stackoverflow.com/questions/31820069/add-hyperlink-to-excel-sheet-created-by-pandas-dataframe-to-excel-method
# https://stackoverflow.com/questions/8287628/proxies-with-python-requests-module
# https://www.list-org.com/