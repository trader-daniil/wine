import datetime
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape

PATH_FILE_WITH_VINES = 'vines_data/wine3.xlsx'


def get_company_age(current_date):
    foundation_year = datetime.datetime(
        year=1920,
        month=1,
        day=1,
    )
    delta_years = current_date - foundation_year.date()
    company_age = round(delta_years.days / 365.25)
    if str(company_age)[-1] in ('2', '3', '4'):
        return f'{company_age} года'
    elif str(company_age)[-1] == '1':
        return f'{company_age} год'
    return f'{company_age} лет'


def parse_vines_from_excel(file):
    excel_data_df = pandas.read_excel(file)
    excel_data_df.fillna('', inplace=True)
    vines = excel_data_df.to_dict(orient='records')
    parsed_vines = defaultdict(list)
    for vine in vines:
        category = vine['Категория']
        parsed_vines[category].append(vine)
    return parsed_vines


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml']),
    )
    current_date = datetime.date.today()
    template = env.get_template('template.html')
    company_age = get_company_age(current_date=current_date)
    vines_categories = parse_vines_from_excel(file=PATH_FILE_WITH_VINES)
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
