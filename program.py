from bs4 import BeautifulSoup
import requests as req

user_agent = ('Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:50.0) '
              'Gecko/20100101 Firefox/50.0')

url = 'http://rkn.gov.ru/communication/register/license/'

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


# Лишние вызовы функций потом удалю, пока посмотрите что да как работает (Алексей)
# Главная страница
save__to__file(
    req.get(url, headers={'User-Agent':user_agent}),
    'main_license_page.html'
)


# Тут страница где лежит ссылка "Отобразить результат"
save__to__file(
    req.post(
        url, 
        headers={'User-Agent':user_agent},
        params={'SERVICE_ID': 12}
    ),
    'middle_2063_page.html'
)


# Тут первая страница 2063 итемов
save__to__file(
    req.post(
        url + '?all=1',
        headers={'User-Agent':user_agent},
        params={'SERVICE_ID': 12}
    ),
    'items__page.html'
)


# Тут все нужные таблицы (Единственный нужный вызов)
for i in range (0, 5, 1):
    save__to__file(
    req.post(
        url + 'p' + str(500 * i) + '/?all=1',
        headers={'User-Agent':user_agent},
        params={'SERVICE_ID': 12}
    ),
    'table__page' + str(i+1) + '.html'
)



# https://python-scripts.com/beautifulsoup-html-parsing
# https://www.softwaretestinghelp.com/selenium-webdriver-commands-selenium-tutorial-17/