
from bs4 import BeautifulSoup
import requests as req
import pandas as pd
from datetime import datetime
import time
from random_user_agent.user_agent import UserAgent
from random_user_agent.params import SoftwareName, OperatingSystem
import xlrd
pd.io.formats.format.header_style = None

user_agent = ('Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:50.0) '
              'Gecko/20100101 Firefox/50.0')

url = 'http://rkn.gov.ru/communication/register/license/'
list__org__url = 'https://www.list-org.com'

software_names = [SoftwareName.CHROME.value]
operating_systems = [OperatingSystem.WINDOWS.value, OperatingSystem.LINUX.value]   

user_agent_rotator = UserAgent(software_names=software_names, operating_systems=operating_systems, limit=100)

user_agents = user_agent_rotator.get_user_agents()

start = datetime.now()
print('start')

try: 
    web__sites = pd.read_excel('table__full.xlsx')
    web__sites = list(web__sites['Веб-сайт'])
    print('Найден файл table__full.xlsx / Продолжение поиска ссылок...')
except: 
    web__sites = []
    print('Не найден файл table__full.xlsx / Старт поиска ссылок...')
    pass


# Парсинг компаний по ИНН на List-Org
def get__company__contacts(resp, col):
    soup = BeautifulSoup(resp.text, 'lxml')
    soup = soup.find('div', class_='org_list')
    if soup == None:
        return('- ошибка запроса -')
    elif soup != None:
        soup = soup.find('a')
        try:
            link = soup['href']
        except: return('- не найдено -')
        response__contacts = req.get(
                        f'{list__org__url}{link}',
                        headers={'User-Agent':user_agent},
                    )
        site = BeautifulSoup(response__contacts.text, 'lxml')
        site = site.find('div', class_='sites')
        if site == None: return '- ошибка запроса -'
        site__link = '- не найдено -'
        if site != None:
            site__link = site.find('a', class_='site')
            if site__link == None: return '- не найдено -'
            else:
                link__span = site__link.find_all('span')
                for span in link__span:
                    link__span = span.decompose()
                    for a in site__link:
                        a = a.replace('http://', '')
                        a = a.replace('https://', '')
                        a = a.replace('www.', '')
                        if (a.__contains__('www') & ~a.__contains__('http://')):
                            a = a.replace(a, 'http://'+a)
                            return(a)
                        elif (~a.__contains__('www') & ~a.__contains__('http//')):
                            a = a.replace(a, 'http://'+a)
                            return(a)
                        return(a)


# Замена длинных названий на сокращения
def replace__text(text):
    if text.__contains__('Акционерное общество'):
        text = text.replace('Акционерное общество', 'АО')
    elif text.__contains__('Закрытое акционерное общество'):
        text = text.replace('Закрытое акционерное общество', 'ЗАО')
    elif text.__contains__('Индивидуальный предприниматель'):
        text = text.replace('Индивидуальный предприниматель', 'ИП')
    elif text.__contains__('Публичное акционерное общество'):
        text = text.replace('Публичное акционерное общество', 'ПАО')
    elif text.__contains__('Общество с ограниченной ответственностью'):
        text = text.replace('Общество с ограниченной ответственностью', 'ООО')
    elif text.__contains__('Открытое акционерное общество'):
        text = text.replace('Открытое акционерное общество', 'ОАО')
    elif text.__contains__('Федеральное государственное бюджетное научное учреждение'):
        text = text.replace('Федеральное государственное бюджетное научное учреждение', 'ФГБНУ')
    elif text.__contains__('Федеральное государственное унитарное предприятие'):
        text = text.replace('Федеральное государственное унитарное предприятие', 'ФГУП')
    elif text.__contains__('Муниципальное автономное учреждение'):
        text = text.replace('Муниципальное автономное учреждение', 'МАУ')
    elif text.__contains__('Непубличное акционерное общество'):
        text = text.replace('Непубличное акционерное общество', 'НАО')
    elif text.__contains__('Федеральное государственное бюджетное образовательное учреждение высшего образования'):
        text = text.replace('Федеральное государственное бюджетное образовательное учреждение высшего образования',
                            'ФГБО')
    elif text.__contains__('Товарищество собственников жилья'):
        text = text.replace('Товарищество собственников жилья', 'ТСЖ')
    elif text.__contains__('Товарищество на вере'):
        text = text.replace('Товарищество на вере', 'ТНВ')
    elif text.__contains__('Федеральное государственное автономное образовательное учреждение высшего образования'):
        text = text.replace('Федеральное государственное автономное образовательное учреждение высшего образования',
                            'ФГАОУВО')
    elif text.__contains__('Муниципальное предприятие'):
        text = text.replace('Муниципальное предприятие', 'МП')
    elif text.__contains__('Муниципальное унитарное предприятие'):
        text = text.replace('Муниципальное унитарное предприятие', 'МУП')
    elif text.__contains__('Муниципальное учреждение'):
        text = text.replace('Муниципальное учреждение', 'МУ')
    elif text.__contains__('Некоммерческое партнерство'):
        text = text.replace('Некоммерческое партнерство', 'НП')
    elif text.__contains__('Некоммерческое учреждение'):
        text = text.replace('Некоммерческое учреждение', 'НУ')

    quotes = ['“', '”', '»', '«']
    for q in quotes:
        text = text.replace(q, '\"')

    idx = text.find('\"')
    if(idx != -1):
        text = text[idx:] + ", " + text[:idx]

    return text

def groupby__inn(table):
    df = table.sort_values(by='Номер лицензии')
    groupby_cols = ['Наименование лицензиата', 'ИНН лицензиата']
    grouped = df.groupby(groupby_cols)

    res = pd.DataFrame(index=range(len(grouped.indices)), columns=df.columns)
    for idx, col in enumerate(groupby_cols):
        res[col] = list(map(lambda x: x[idx], grouped.indices))

    columns = list(set(df.columns).difference(set(groupby_cols + ['Регион'])))

    for row_idx, (idx, part_df) in enumerate(grouped):
        for col in columns:
            res.loc[row_idx, col] = '\n'.join(list(map(str, part_df.groupby('Номер лицензии')[col].first())))

    res['Регион'] = grouped.apply(lambda x: ', '.join(x['Регион'].astype('str').unique())).values

    return res

def replace__region(t):
    t['Регион'] = t['Регион'].map({
        1:'Республика Адыгея (Адыгея)', 2:'Республика Башкортостан', 3:'Республика Бурятия', 4:'Республика Алтай', 5:'Республика Дагестан', 
        6:'Республика Ингушетия', 7:'Кабардино-Балкарская Республика', 8:'Республика Калмыкия', 9:'Карачаево-Черкесская Республика', 10:'Республика Карелия', 
        11:'Республика Коми', 12:'Республика Марий Эл', 13:'Республика Мордовия', 14:'Республика Саха (Якутия)', 15:'Республика Северная Осетия - Алания', 
        16:'Республика Татарстан (Татарстан)', 17:'Республика Тыва', 18:'Удмуртская Республика', 19:'Республика Хакасия', 20:'Чеченская Республика', 
        21:'Чувашская  Республика - Чувашия', 22:'Алтайский край', 23:'Краснодарский край', 24:'Красноярский край', 25:'Приморский край', 
        26:'Ставропольский край', 27:'Хабаровский край', 28:'Амурская область', 29:'Архангельская область', 30:'Астраханская область', 
        31:'Белгородская область', 32:'Брянская область', 33:'Владимирская область', 34:'Волгоградская область', 35:'Вологодская область', 
        36:'Воронежская область', 37:'Ивановская область', 38:'Иркутская область', 39:'Калининградская область', 40:'Калужская область', 
        41:'Камчатский край', 42:'Кемеровская область', 43:'Кировская область', 44:'Костромская область', 45:'Курганская область', 
        46:'Курская область', 47:'Ленинградская область', 48:'Липецкая область', 49:'Магаданская область', 50:'Московская область', 
        51:'Мурманская область', 52:'Нижегородская область', 53:'Новгородская область', 54:'Новосибирская область', 55:'Омская область', 
        56:'Оренбургская область', 57:'Орловская область', 58:'Пензенская область', 59:'Пермский край', 60:'Псковская область',
        61:'Ростовская область', 62:'Рязанская область', 63:'Самарская область', 64:'Саратовская область', 65:'Сахалинская область', 
        66:'Свердловская область', 67:'Смоленская область', 68:'Тамбовская область', 69:'Тверская область', 70:'Томская область', 
        71:'Тульская область', 72:'Тюменская область', 73:'Ульяновская область', 74:'Челябинская область', 75:'Забайкальский край', 
        76:'Ярославская область', 77:'Москва', 78:'Санкт-Петербург', 79:'Еврейская автономная область', 83:'Ненецкий автономный округ', 
        86:'Ханты-Мансийский автономный округ - Югра', 87:'Чукотский автономный округ', 89:'Ямало-Ненецкий автономный округ', 99:'Байконур', 130:'Республика Крым', 131:'Севастополь'
    })
    return t

#Подготовка таблицы к выводу в exсel
def format__table(df, max_regions_number):
    cols = ['Номер лицензии',
            'Регион',
            'Наименование лицензиата',
            'Поиск в Google',
            'Поиск на List-Org',
            'Веб-сайт',
            'ИНН лицензиата',
            'Срок действия',
            'День начала оказания услуг(не позднее)'
    ]

    res = pd.DataFrame(df, columns=cols)
    res = replace__region(res)
    res = groupby__inn(res)

    res['ИНН лицензиата'] = res['ИНН лицензиата'].astype('str')
    res['ИНН лицензиата'] = res['ИНН лицензиата'].apply(lambda x: x.zfill(10) if len(x) <= 10 else x.zfill(12))

    res['Поиск в Google'] = res['Наименование лицензиата'].\
        apply(lambda x: f'https://www.google.com/search?q={x.replace(" ", "+")}')
    res['Поиск на List-Org'] = res['ИНН лицензиата'].\
        apply(lambda x: f'https://www.list-org.com/search?type=inn&val={x}')

    res['Регион'] = res['Регион'].apply(lambda x: 'РФ' if len(x.split(', ')) == max_regions_number else x)
    res['Наименование лицензиата'] = res['Наименование лицензиата'].apply(replace__text)
    res['Веб-сайт'] = ""

    res = res.sort_values(by='Наименование лицензиата')
    res = res.set_index(pd.Index(range(res.shape[0])))
    return res


# Сохранить в Excel-файл
def excel__writer(table, path):
    writer = pd.ExcelWriter(path, engine='xlsxwriter') # pylint: disable=abstract-class-instantiated
    table.to_excel(writer, sheet_name='Sheet1', index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    base_format_dict = {'align': 'center', 'valign': 'top', 'border': 1, 'border_color': '#808080'}
    row_dict = {**base_format_dict, 'text_wrap': True}
    url_center_dict = {**base_format_dict, 'hyperlink': True}
    url_left_dict = {**base_format_dict, 'hyperlink': True, 'align': 'left'}

    for col_idx, col in enumerate(table.columns):
        series = table[col].astype(str)
        max_len = max((series.apply(lambda x: max(map(len, x.split('\n')))).max(), len(str(series.name)))) + 5
        if col == 'Наименование лицензиата':
            # специальный формат и ширина для столбца 'Наименование лицензиата'
            worksheet.set_column(col_idx, col_idx, 80)
            for row_idx, (id, val) in enumerate(zip(table['Номер лицензии'].values, table[col].values)):
                search_id = id.split('\n')[0]
                cell_format = workbook.add_format({**url_left_dict, 'bg_color': ('#FFFFFF' if row_idx % 2 == 0 else '#CCCCCC')})
                worksheet.write_url(row_idx + 1, col_idx,
                                    url + f'?id={search_id}&all=1', string=val, cell_format=cell_format)
        elif col == 'Поиск в Google':
            worksheet.set_column(col_idx, col_idx, 20)
            for row_idx, (g) in enumerate(table['Поиск в Google']):
                cell_format = workbook.add_format({**url_center_dict, 'bg_color': ('#FFFFFF' if row_idx % 2 == 0 else '#CCCCCC')})
                worksheet.write_url(row_idx + 1, col_idx, g, string='Найти', cell_format=cell_format)
        elif col == 'Поиск на List-Org':
            worksheet.set_column(col_idx, col_idx, 20)
            for row_idx, (q) in enumerate(table['Поиск на List-Org']):
                cell_format = workbook.add_format({**url_center_dict, 'bg_color': ('#FFFFFF' if row_idx % 2 == 0 else '#CCCCCC')})
                worksheet.write_url(row_idx + 1, col_idx, q, string='Найти по ИНН', cell_format=cell_format)
        elif col == 'Регион':
            worksheet.set_column(col_idx, col_idx, 30)
            for row_idx, r in enumerate(table['Регион']):
                cell_format = workbook.add_format({**row_dict, 'bg_color': ('#FFFFFF' if row_idx % 2 == 0 else '#CCCCCC')})
                worksheet.write(row_idx + 1, col_idx, str(r), cell_format)
        else:
            worksheet.set_column(col_idx, col_idx, max_len)
            for row_idx, r in enumerate(table[col]):
                cell_format = workbook.add_format({**row_dict, 'bg_color': ('#FFFFFF' if row_idx % 2 == 0 else '#CCCCCC')})
                worksheet.write(row_idx + 1, col_idx, str(r), cell_format)

    writer.save()
    print('Saved into', path)

# сохранить страницу в dataframe
def read__part__dataframe(resp, start_idx):
    soup = BeautifulSoup(resp.text, 'lxml')
    soup = soup.find("table", id="ResList1")
    try:
        t = pd.read_html(str(soup))
    except: return 'Записей не найдено'
    table = t[0].drop("Unnamed: 5", axis=1)
    index = pd.Index(range(start_idx, start_idx + table.shape[0]))
    table = table.set_index(index)
   
    return table

# Старт программы
regions = [
    1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 
    31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60,
    61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 83, 86, 87, 89, 99, 130, 131
]
max_regions_number = len(regions)
# Для теста
regions__low = regions[0:1]
max_regions_number = len(regions__low)

dfs = []
arr = []
for index, region in enumerate(regions__low):
    i = 0
    while True: 
        time.sleep(2)
        print(f'req # {index + 1} / 86, i = {i}')
        if index % 2 != 0: time.sleep(30)
        response = req.post(
            url + 'p' + str(500 * i) + '/?all=1',
            headers={'User-Agent':user_agent},
            params={'SERVICE_ID': 12, 'REGION_ID': region},
            timeout=5
        )
        if 'Записей не найдено' in response.text:   
           break
        res = read__part__dataframe(response, 500 * i)
        dfs.append(res)
        arr += [region] * len(res)
        i += 1
        if res.shape[0] < 500:
            break

full_df = pd.concat(dfs, axis=0)
full_df['Регион'] = arr
full_df = format__table(full_df, max_regions_number)
excel__writer(full_df, 'table.xlsx')



counter = 0
# for col in full_df['ИНН лицензиата']:
#     time.sleep(1.5)
#     counter += 1
#     user_agent = user_agent_rotator.get_random_user_agent()
#     response = req.get(
#         f'{list__org__url}/search?type=inn',
#             headers={'User-Agent':user_agent},
#             params={'val': col},
#             timeout=10,
#         )
#     link = get__company__contacts(response, col)
#     if link == None:
#         print('sleep 5s')
#         time.sleep(5)

#     if counter == 21:
#         full_df['Веб-сайт'] = pd.Series(web__sites).fillna(' - пока не найдено - ')
#         excel__writer(full_df, 'table__full.xlsx')
#         print(datetime.now(), f'Выполнено {len(web__sites)} / {len(full_df["Регион"])} запросов. Промежуточная таблица сохранена. Остановка программы на 5 минут. \n Необходимо ввести проверочный код на https://www.list-org.com/bot')
#         time.sleep(300)
#         counter = 0
    
#     if len(web__sites) == len(full_df['Регион']):
#         break
#     print('link -', counter + 1, 'статус:', link)
#     web__sites.append(link)

for row, col in enumerate(full_df['ИНН лицензиата']):
    time.sleep(1.5)
    counter += 1
    user_agent = user_agent_rotator.get_random_user_agent()
    try:
        if web__sites[row] == '- не найдено -' or web__sites[row].__contains__('http'):
            link = web__sites[row]
            print(f'Ссылка {row + 1} : /{link}/ уже обрабатывалась, переход к следующей...')
            counter = 0
        else:
            response = req.get(
            f'{list__org__url}/search?type=inn',
                headers={'User-Agent':user_agent},
                params={'val': col},
                timeout=10,
            )
            link = get__company__contacts(response, col)
            if link == None:
                print('sleep 5s')
                time.sleep(5)
            web__sites[row] = link

            if counter % 20 == 0:
                full_df['Веб-сайт'] = pd.Series(web__sites).fillna(' - пока не найдено - ')
                excel__writer(full_df, 'table__full.xlsx')
                print(datetime.now(), f'Выполнено {row} / {len(full_df["Регион"])} запросов. Промежуточная таблица сохранена. Остановка программы на 5 минут. \n Необходимо ввести проверочный код на https://www.list-org.com/bot')
                time.sleep(300)
                counter = 0
            print('link -', counter + 1, 'статус:', link)

            if row == len(full_df['Регион']):
                print('Завершен поиск ссылок')
                break
    except: 
        response = req.get(
            f'{list__org__url}/search?type=inn',
                headers={'User-Agent':user_agent},
                params={'val': col},
                timeout=10,
            )
        link = get__company__contacts(response, col)
        if link == None:
            print('sleep 5s')
            time.sleep(5)

        if counter % 20 == 0:
            full_df['Веб-сайт'] = pd.Series(web__sites).fillna(' - пока не найдено - ')
            excel__writer(full_df, 'table__full.xlsx')
            print(datetime.now(), f'Выполнено {row} / {len(full_df["Регион"])} запросов. Промежуточная таблица сохранена. Остановка программы на 5 минут. \n Необходимо ввести проверочный код на https://www.list-org.com/bot')
            time.sleep(300)
            counter = 0
        print('link -', counter + 1, 'статус:', link)
        web__sites.append(link)
        
        if row == len(full_df['Регион']):
            print('Завершен поиск ссылок')
            break
    
full_df['Веб-сайт'] = pd.Series(web__sites)
excel__writer(full_df, 'table__full.xlsx')


print('Заняло времени - ', datetime.now() - start)

# https://python-scripts.com/beautifulsoup-html-parsing
# https://www.softwaretestinghelp.com/selenium-webdriver-commands-selenium-tutorial-17/
# https://stackoverflow.com/questions/35831241/converting-html-to-excel-in-python
# https://www.geeksforgeeks.org/python-convert-an-html-table-into-excel/
# https://stackoverflow.com/questions/31820069/add-hyperlink-to-excel-sheet-created-by-pandas-dataframe-to-excel-method
# https://stackoverflow.com/questions/8287628/proxies-with-python-requests-module
# https://www.list-org.com/