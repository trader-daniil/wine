import argparse
from collections import defaultdict
from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape

YEARS_CAPTIONS = {
    '1': 'год',
    '2': 'года',
    '3': 'года',
    '4': 'года',
    '5': 'лет',
    '6': 'лет',
    '7': 'лет',
    '8': 'лет',
    '9': 'лет',
    '0': 'лет',
}


def get_company_age(foundation_year):
    current_year = datetime.now().year
    company_age = current_year - foundation_year
    last_number_of_age = str(company_age)[-1]
    return f'{company_age} {YEARS_CAPTIONS[last_number_of_age]}'


def parse_wines_from_excel(filepath):
    wines = pandas.read_excel(
        filepath,
        keep_default_na=False,
    ).to_dict(orient='records')
    grouped_wines = defaultdict(list)
    for wine in wines:
        category = wine['Категория']
        grouped_wines[category].append(wine)
    return grouped_wines


def main():
    parser = argparse.ArgumentParser(
        description='Программа возвращает страницу с винами',
    )
    parser.add_argument(
        '--filepath',
        help='Название файла с винами',
        default='wines_data/wines_example.xlsx',
    )
    args = parser.parse_args()

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml']),
    )
    template = env.get_template('template.html')

    company_foundation_year = 1920
    company_age = get_company_age(
        foundation_year=company_foundation_year
    )

    wines_categories = parse_wines_from_excel(filepath=args.filepath)
    rendered_page = template.render(
        age=company_age,
        categories=wines_categories,
    )
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
