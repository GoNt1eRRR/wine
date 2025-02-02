import os
from dotenv import load_dotenv
from datetime import datetime
import pandas
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
from collections import defaultdict


def get_wines_and_category(data_path):
    excel_data_df = pandas.read_excel(
        data_path,
        sheet_name="Лист1",
        usecols=[
            "Категория",
            "Название",
            "Сорт",
            "Цена",
            "Картинка",
            "Акция",
        ],
    )

    products = excel_data_df.to_dict(orient='records')
    grouped_products = defaultdict(list)
    for product in products:
        category = product['Категория']
        grouped_products[category].append(product)
    return grouped_products


def get_winery_age():
    winery_start_age = 1920
    current_year = datetime.now().year
    return current_year - winery_start_age


def get_years_word(wine_age):
    if 11 <= wine_age % 100 <= 19:
        return "лет"

    last_digit = wine_age % 10

    if last_digit == 1:
        return "год"
    elif 2 <= last_digit <= 4:
        return "года"
    else:
        return "лет"


if __name__ == '__main__':
    load_dotenv()
    data_path = os.getenv('WINE_DATA_PATH', 'wine.xlsx')

    special_offer_promotion = os.getenv(
        'SPECIAL_OFFER_PROMOTION',
        'Выгодное предложение'
    )

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    categorized_wines = get_wines_and_category(data_path)
    wine_age = get_winery_age()
    word = get_years_word(wine_age)

    rendered_page = template.render(
        age=wine_age,
        word=word,
        wines=categorized_wines,
        special_offer_promotion=special_offer_promotion
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()