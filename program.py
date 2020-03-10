from bs4 import BeautifulSoup
import requests as req
import pandas as pd
pd.io.formats.format.header_style = None

user_agent = ('Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:50.0) '
              'Gecko/20100101 Firefox/50.0')

url = 'http://rkn.gov.ru/communication/register/license/'

# TODO: 
#       запись в эксель ячейки лицензиата как гиперссылка с сопоставлением элементов массива hrefs и текста ячеек
#       запаковать скрипт в исполняемый файл, чтобы запускался без питона на компьютере

def excel__writer(table, path):
    writer = pd.ExcelWriter(path, engine='xlsxwriter')
    table.to_excel(writer, sheet_name='Sheet1', index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    for col_idx, col in enumerate(table.columns):
        series = table[col]
        max_len = max((series.astype(str).map(len).max(), len(str(series.name)))) + 1
        if col == 'ИНН лицензиата':
            # специальный формат для столбца 'ИНН лицензиата' (10 цифр)
            cell_format = workbook.add_format({'num_format': '0' * 10})
            worksheet.set_column(col_idx, col_idx, max_len, cell_format)
        elif col == 'Наименование лицензиата':
            # специальный формат и ширина для столбца 'Наименование лицензиата'
            cell_format = workbook.get_default_url_format()
            worksheet.set_column(col_idx, col_idx, 80, cell_format)
            for row_idx, (id, val) in enumerate(zip(table['Номер лицензии'].values, table[col].values)):
                worksheet.write_url(row_idx + 1, col_idx, url + f'?id={id}&all=1', string=val)
        else:
            worksheet.set_column(col_idx, col_idx, max_len)

    writer.save()
    print('Saved into', path)

# сохранить страницу в dataframe
def read__part__dataframe(resp, start_idx):
    soup = BeautifulSoup(resp.text, 'lxml')
    soup = soup.find("table", id="ResList1")
    t = pd.read_html(str(soup))
    table = t[0].drop("Unnamed: 5", axis=1)
    table.drop(0, axis=0, inplace=True)
    index = pd.Index(range(start_idx, start_idx + table.shape[0]))
    table = table.set_index(index)
    return table

# #формирование гиперссылки
# def make_hyperlink(id, value):
#     url_str = url + f'?id={id}&all=1'
#     val = value.replace("\"", "\"\"")
#     return f'=HYPERLINK("{url_str}", "{val}")'
#
# #добавление гиперссылок
# def add_hyperlinks(table):
#     res = pd.DataFrame(table)
#     res['Наименование лицензиата'] = res[['Наименование лицензиата', 'Номер лицензии']] \
#         .apply(lambda x: make_hyperlink(x['Номер лицензии'], x['Наименование лицензиата']), axis=1)
#     return res

dfs = []
i = 0
while True:
    response = req.post(
        url + 'p' + str(500 * i) + '/?all=1',
        headers={'User-Agent':user_agent},
        params={'SERVICE_ID': 12}
    )
    if 'Записей не найдено' in response.text:
        break

    dfs.append(read__part__dataframe(response, 500 * i))
    i += 1

full_df = pd.concat(dfs, axis=0)
excel__writer(full_df, 'table.xlsx')



test = req.get('http://google.com/search?q=Аквамарин', headers={'User-Agent':user_agent})

f = open('res.html', 'w')
soup = BeautifulSoup(test.text, 'lxml')
soup = soup.find_all('div', {'class': 'r'})
for div in soup:
    cites = []
    cites.append(div.find('cite').get_text())
    f.write(str(cites))
f.close


# https://python-scripts.com/beautifulsoup-html-parsing
# https://www.softwaretestinghelp.com/selenium-webdriver-commands-selenium-tutorial-17/
# https://stackoverflow.com/questions/35831241/converting-html-to-excel-in-python
# https://www.geeksforgeeks.org/python-convert-an-html-table-into-excel/
# https://stackoverflow.com/questions/31820069/add-hyperlink-to-excel-sheet-created-by-pandas-dataframe-to-excel-method
