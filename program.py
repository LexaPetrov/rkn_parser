from bs4 import BeautifulSoup
import requests as req
import pandas as pd
from datetime import datetime
import time
pd.io.formats.format.header_style = None

user_agent = ('Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:50.0) '
              'Gecko/20100101 Firefox/50.0')

url = 'http://rkn.gov.ru/communication/register/license/'
list__org__url = 'https://www.list-org.com'

proxies = {
    "http": 'http://195.158.232.16:8080', 
    "https": 'http://195.158.232.16:8080', 
}

# TODO: запаковать скрипт в исполняемый файл, чтобы запускался без питона на компьютере

start = datetime.now()
print('start')

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
    else:
        return text
    return text

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

# Сохранить в Excel-файл
def excel__writer(table, path):
    writer = pd.ExcelWriter(path, engine='xlsxwriter') # pylint: disable=abstract-class-instantiated
    table.to_excel(writer, sheet_name='Sheet1', index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    for col_idx, col in enumerate(table.columns):
        series = table[col]
        max_len = max((series.astype(str).map(len).max(), len(str(series.name)))) + 5
        if col == 'ИНН лицензиата':
            # специальный формат для столбца 'ИНН лицензиата' (10 цифр)
            worksheet.set_column(col_idx, col_idx, max_len)
            for row_idx, (id, val) in enumerate(zip(table['Номер лицензии'].values, table[col].values)):
                row_format = workbook.add_format({'num_format': '0' * 10, 'align':'center', 'bg_color': '#FFFFFF' if row_idx%2==0 else '#CCCCCC', 'border':1, 'border_color':'#808080'})
                worksheet.write(row_idx+1, col_idx, val, row_format)

        elif col == 'Наименование лицензиата':
            # специальный формат и ширина для столбца 'Наименование лицензиата'
            cell_format = workbook.get_default_url_format()
            worksheet.set_column(col_idx, col_idx, 80, cell_format)
            for row_idx, (id, val) in enumerate(zip(table['Номер лицензии'].values, table[col].values)):
                row_format = workbook.add_format({'font_color':'#007AA6', 'bg_color':'#FFFFFF' if row_idx%2==0 else '#CCCCCC', 'border':1, 'border_color':'#808080'})
                val = replace__text(val)
                worksheet.write_url(row_idx + 1, col_idx, url + f'?id={id}&all=1', string=val, cell_format=row_format)
        elif col == 'Поиск в Google':
            cell_format = workbook.get_default_url_format()
            worksheet.set_column(col_idx, col_idx, 20, cell_format)
            for row_idx, (g) in enumerate(table['Поиск в Google']):
                row_format = workbook.add_format({'align':'center', 'font_color':'#007AA6', 'bg_color':'#FFFFFF' if row_idx%2==0 else '#CCCCCC', 'border':1, 'border_color':'#808080'})
                g = g.replace(' ', '+')
                worksheet.write_url(row_idx + 1, col_idx, g, string='Найти', cell_format=row_format)
        elif col == 'Поиск на List-Org':
            cell_format = workbook.get_default_url_format()
            worksheet.set_column(col_idx, col_idx, 20, cell_format)
            for row_idx, (q) in enumerate(table['Поиск на List-Org']):
                row_format = workbook.add_format({'align':'center', 'font_color':'#007AA6','bg_color':'#FFFFFF' if row_idx%2==0 else '#CCCCCC', 'border':1, 'border_color':'#808080'})
                worksheet.write_url(row_idx + 1, col_idx, q, string='Найти по ИНН', cell_format=row_format)
            
        # elif col == 'Веб-сайт':
        #     cell_format = workbook.get_default_url_format()
        #     worksheet.set_column(col_idx, col_idx, 20, cell_format)
        #     for row_idx, (q) in enumerate(table['Веб-сайт']):
        #         worksheet.write_url(row_idx + 1, col_idx, q, string=q, cell_format=align_format)
        else:
            worksheet.set_column(col_idx, col_idx, max_len)
            for row_idx, (id, val) in enumerate(zip(table['Номер лицензии'].values, table[col].values)):
                row_format = workbook.add_format({'align':'center', 'bg_color': '#FFFFFF' if row_idx%2==0 else '#CCCCCC', 'border':1, 'border_color':'#808080'})
                worksheet.write(row_idx+1, col_idx, str(val), row_format)

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
    table.drop(0, axis=0, inplace=True)
    res = []
    res2 = []

    for col in table['Наименование лицензиата']:
        res.append(f'https://www.google.com/search?q={col}')

    for col in table['ИНН лицензиата']:
        res2.append(f'https://www.list-org.com/search?type=inn&val={col}')
    
    table['Поиск в Google'] = res
    table['Поиск на List-Org'] = res2
    index = pd.Index(range(start_idx, start_idx + table.shape[0]))
    table = table.set_index(index)
   
    return table

# Старт программы
regions = [
    1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 
    31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60,
    61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 83, 86, 87, 89, 99, 130, 131
]
# Для теста
regions__low = [
    130, 131
]
dfs = []
arr = []
for index, region in enumerate(regions__low):
    i = 0
    while True: 
        time.sleep(2)
        print(f'req # {index + 1} / 86, i = {i}')
        if index % 2 == 0: time.sleep(30)
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
        for j in range(len(res)):
            arr.append(region)
            j += 1
        i += 1
    

full_df = pd.concat(dfs, axis=0)
cols = full_df.columns.tolist()
cols = cols[:2]+cols[-2:]+cols[2:-2]
full_df = full_df[cols]
full_df['Регион'] = arr
full_df = replace__region(full_df)
excel__writer(full_df, 'table.xlsx')
web__sites = []
counter = 0
for col in full_df['ИНН лицензиата']:
    time.sleep(1.5)
    counter += 1
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

    if counter == 21:
        full_df['Веб-сайт'] = pd.Series(web__sites).fillna(' - пока не найдено - ')
        excel__writer(full_df, 'table__full.xlsx')
        print(datetime.now(), f'Выполнено {len(web__sites)} / {len(full_df["Регион"])} запросов. Промежуточная таблица сохранена. Остановка программы на 5 минут')
        time.sleep(300)
        counter = 0
    
    if len(web__sites) == len(full_df['Регион']):
        break
    print('link -', counter + 1, 'статус:', link)
    web__sites.append(link)

full_df['Веб-сайт'] = web__sites
excel__writer(full_df, 'table__full.xlsx')


print('Заняло времени - ', datetime.now() - start)

# https://python-scripts.com/beautifulsoup-html-parsing
# https://www.softwaretestinghelp.com/selenium-webdriver-commands-selenium-tutorial-17/
# https://stackoverflow.com/questions/35831241/converting-html-to-excel-in-python
# https://www.geeksforgeeks.org/python-convert-an-html-table-into-excel/
# https://stackoverflow.com/questions/31820069/add-hyperlink-to-excel-sheet-created-by-pandas-dataframe-to-excel-method
# https://stackoverflow.com/questions/8287628/proxies-with-python-requests-module
# https://www.list-org.com/