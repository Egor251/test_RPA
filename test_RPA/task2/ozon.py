from bs4 import BeautifulSoup
import webbrowser
import time

import pyautogui as pg
import pyperclip

import config


def get_source_auto(url):
    webbrowser.open(url, new=1, autoraise=True)
    # Передвижение мыши
    time.sleep(3)
    pg.click(150, 200, button='right')  # Передвигаем к точке относительно экрана
    pg.move(20, 200, duration=0.5)
    pg.click()
    time.sleep(3)
    pg.hotkey('ctrl', 'a')
    pg.hotkey('ctrl', 'c')
    time.sleep(2)
    pg.hotkey('ctrl', 'w')
    pg.hotkey('ctrl', 'w')
    text = pyperclip.paste()
    return text


def parse(html):
    soup = BeautifulSoup(html, 'html.parser')

    category = soup.find_all('a', class_='d3y_10 tsBody500Medium')  # Категории
    items = soup.find_all('div', class_='ba8 b9a ca w3i_23')  # Название товара
    rates = soup.find_all('div', class_="w3i_23 w5i_23 t4 t5 t6 tsBodyMBold")  # Рейтинг и количество отзывов (будут получены далее)
    prices = soup.find_all('div', class_='c306-a0')

    # Категории
    cat = []
    for categ in category:
        cat.append(categ.contents[0])
    ca = f'{cat[0]}/{cat[1]}/{cat[2]}'

    name_list = []
    # Название
    for item in items:
        tmem = soup.find('span', class_='tsBody500Medium')
        name_list.append(item.get_text())

    # Рейтинги и отзывы
    sss = []
    for rate in rates:
        span = rate.find_all('span')  # Распаковываем первый span
        sp = []
        for spa in span:
            sp.append(spa.find_all('span')) # Распаковываем второй span
        ss = []
        for s in sp:
            try:
                ss.append(s[0].get_text())  # вытаскиваем сами значения
            except IndexError:
                pass
        sss.append(ss)

    pri = []
    for price in prices:
        pric = price.find_all('span')
        pri_tmp = pric[0].get_text().split()
        if len(pri_tmp) == 3:
            pr = pri_tmp[0]+pri_tmp[1]
        elif len(pri_tmp) == 2:
            pr = pri_tmp[0]
        else:
            pr = 0
        pri.append(pr)

    output_list = []
    for i, item in enumerate(name_list):
        try:
            r = sss[i][0]
            e = sss[i][1].split()[0]
        except IndexError:
            r = 0
            e = 0
        output_list.append([ca, item, float(pri[i]), r, e])

    return output_list

def prepare_link(req):
    req = req.replace(' ', '+')
    return req

def main_parse(req):
    req = prepare_link(req)
    url = f'https://www.ozon.ru/search/?text={req}&from_global=true'
    html = get_source_auto(url)
    return parse(html)


if __name__ == '__main__':
    #print(parse(get_html('https://www.ozon.ru/search/?text=тренировочный+хирургический+набор+для+шиться+ран&from_global=true')))
    #browser()
    #text = get_source_auto('https://www.ozon.ru/search/?text=тренировочный+хирургический+набор+для+шиться+ран&from_global=true')

    print(main_parse(config.request))

