import collections
import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler
from pathlib import Path

import configargparse
import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape


def main():
    winery_foundation_year = 1920

    parser = configargparse.ArgParser(
        description='Wine Shop command line application',
    )
    parser.add_argument(
        '-w',
        '--wines_path',
        default='wines.xlsx',
        env_var='WINE_SHOP_DATA_PATH',
        help='path to Excel file containing information about wines',
    )
    parser.add_argument(
        '-t',
        '--template_path',
        default='template.html',
        env_var='WINE_SHOP_TEMPLATE_PATH',
        help='path to template HTML file',
    )
    args = parser.parse_args()
    wines_path = args.wines_path
    template_path = args.template_path

    wines = pandas.read_excel(
        wines_path,
        sheet_name='Лист1',
        dtype={'Цена': int},
        keep_default_na=False
    ).to_dict('records')
    grouped_wines = collections.defaultdict(list)
    for row in wines:
        grouped_wines[row['Категория']].append(row)
    sorted_wine_groups = dict()
    for wines_group in sorted(grouped_wines):
        sorted_wine_groups[wines_group] = grouped_wines[wines_group]

    env = Environment(
        loader=FileSystemLoader(Path(template_path).parent),
        autoescape=select_autoescape(['html', 'xml']),
    )
    template = env.get_template(Path(template_path).name)
    rendered_page = template.render(
        wines=sorted_wine_groups,
        winery_age=datetime.datetime.today().year - winery_foundation_year,
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
