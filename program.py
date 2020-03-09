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

# Вывести в файл весь документ
# def save__to__file(resp, path):
#     f = open(path, 'w')
#     soup = BeautifulSoup(resp.text, 'lxml')
#     # print(soup.find("table", id="ResList1"))
#     soup = soup.prettify()
#     f.write(soup)
#     f.close()


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
    print(hrefs)
    headers = soup.find("thead")
    t = pd.read_html(str(soup))
    table = t[0].drop("Unnamed: 5", axis=1)
    # print(table)
    # table.to_excel(path, index=False)

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




# https://python-scripts.com/beautifulsoup-html-parsing
# https://www.softwaretestinghelp.com/selenium-webdriver-commands-selenium-tutorial-17/
# https://stackoverflow.com/questions/35831241/converting-html-to-excel-in-python
# https://www.geeksforgeeks.org/python-convert-an-html-table-into-excel/