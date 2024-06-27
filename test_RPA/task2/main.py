import ozon
import bank
import config
import excel


def main():
    ozon_data = ozon.main_parse(config.request)
    for item in ozon_data:
        item.append(round(item[2]/bank.usd(), 2))
    print(ozon_data)
    headers = ['Категория', 'Наименование', 'Цена (руб)', 'Оценка', 'Кол-во отзывов', ' Цена (USD)']
    excel.make_xlsx(ozon_data, headers)



if __name__ == '__main__':
    main()
