import collections
import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape


def main():
    winery_foundation_year = 1920
    wines = pandas.read_excel(
        'wine3.xlsx',
        sheet_name='Лист1',
        dtype={'Цена': int},
        keep_default_na=False
    ).to_dict('records')
    grouped_wines = collections.defaultdict(list)

    for row in wines:
        wine = {
            'name': row['Название'],
            'sort': row['Сорт'],
            'price': row['Цена'],
            'picture': row['Картинка'],
            'sale': row['Акция'],
        }
        grouped_wines[row['Категория']].append(wine)

    sorted_wine_groups = dict()
    for key in sorted(grouped_wines):
        sorted_wine_groups[key] = grouped_wines[key]

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml']),
    )

    template = env.get_template('template.html')

    rendered_page = template.render(
        wines=sorted_wine_groups,
        age=datetime.datetime.today().year - winery_foundation_year,
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
