from datetime import datetime
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape
import dateutil

WINES_FILEPATH = 'vines_data/wine3.xlsx'


YEARS = {
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

def get_company_age(founded_at):
    current_year = datetime.now().year
    company_age = current_year - founded_at
    last_number_of_age = str(company_age)[-1]
    return f'{company_age} {YEARS[last_number_of_age]}'


def parse_vines_from_excel(filepath):
    excel_wines = pandas.read_excel(filepath)
    excel_wines.fillna('', inplace=True)
    wines = excel_wines.to_dict(orient='records')
    grouped_wines = defaultdict(list)
    for wine in wines:
        category = wine['Категория']
        grouped_wines[category].append(wine)
    return grouped_wines


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml']),
    )
    template = env.get_template('template.html')
    wines_company_founded_at = 1920
    company_age = get_company_age(
        founded_at=wines_company_founded_at
    )
    vines_categories = parse_vines_from_excel(filepath=WINES_FILEPATH)
    rendered_page = template.render(
        age=company_age,
        categories=vines_categories,
    )
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
