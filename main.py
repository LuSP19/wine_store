from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
import datetime
import pandas
import collections


def main():
    start = datetime.datetime(year=1920, month=1, day=1)
    excel_data_df = pandas.read_excel(
        'wine3.xlsx',
        sheet_name='Лист1',
        dtype={'Цена': int},
        keep_default_na=False
    )
    wine_from_file = excel_data_df.to_dict('records')
    wine = collections.defaultdict(list)

    for row in wine_from_file:
        category = {
            'name': row['Название'],
            'sort': row['Сорт'],
            'price': row['Цена'],
            'picture': row['Картинка'],
            'sale': row['Акция'],
        }
        wine[row['Категория']].append(category)

    wine_sorted = dict()
    for key in sorted(wine):
        wine_sorted[key] = wine[key]

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml']),
    )

    template = env.get_template('template.html')

    rendered_page = template.render(
        wine=wine_sorted,
        age=(datetime.datetime.today() - start).days // 365,
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
