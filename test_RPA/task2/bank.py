import requests
import xml.etree.ElementTree as ET




def usd():
    url = 'https://www.cbr-xml-daily.ru/daily_utf8.xml'
    response = requests.get(url)
    usd = float(ET.fromstring(response.text).find("./Valute[CharCode='USD']/Value").text.replace(',', '.'))
    return usd


if __name__ == '__main__':
    usd()
