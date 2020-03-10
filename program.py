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
    worksheet = writer.sheets['Sheet1']

    for idx, col in enumerate(table.columns):
        series = table[col]
        max_len = min(80, max((series.astype(str).map(len).max(), len(str(series.name)))) + 1)
        worksheet.set_column(idx, idx, max_len)

    writer.save()
    print('Saved into', path)


# Вывести в файл только таблицы
def save__to__file(resp, path):
    f = open(path, 'w')
    soup = BeautifulSoup(resp.text, 'lxml')
    soup = soup.find("table", id="ResList1")
    if soup != None:
        f.write(str(soup))
    else:
        f.close()


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

#формирование гиперссылки
def make_hyperlink(id, value):
    url_str = url + f'?id={id}&all=1'
    val = value.replace("\"", "\"\"")
    return f'=HYPERLINK("{url_str}", "{val}")'

#добавление гиперссылок
def add_hyperlinks(table):
    res = pd.DataFrame(table)
    res['Наименование лицензиата'] = res[['Наименование лицензиата', 'Номер лицензии']] \
        .apply(lambda x: make_hyperlink(x['Номер лицензии'], x['Наименование лицензиата']), axis=1)
    return res

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

full_df = add_hyperlinks(pd.concat(dfs, axis=0))
excel__writer(full_df, 'table.xlsx')



test = req.get('http://google.com/search?q=Аквамарин', headers={'User-Agent':user_agent})

f = open('res.html', 'w')
soup = BeautifulSoup(test.text, 'lxml')
soup = soup.find_all('div', {'class': 'r'})
for div in soup:
    cites = []
    cites.append(div.find('cite').get_text())
    f.write(str(cites))
    # f.write(str(div.find('cite').get_text()))
f.close


# https://python-scripts.com/beautifulsoup-html-parsing
# https://www.softwaretestinghelp.com/selenium-webdriver-commands-selenium-tutorial-17/
# https://stackoverflow.com/questions/35831241/converting-html-to-excel-in-python
# https://www.geeksforgeeks.org/python-convert-an-html-table-into-excel/
# https://stackoverflow.com/questions/31820069/add-hyperlink-to-excel-sheet-created-by-pandas-dataframe-to-excel-method