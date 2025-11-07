import pandas as pd
from jinja2 import Environment, FileSystemLoader, TemplateNotFound
from collections import defaultdict
import argparse
import os
from pathlib import Path
from dotenv import load_dotenv
from datetime import datetime


def create_parser():
    parser = argparse.ArgumentParser(description='Генератор сайта винного магазина')
    parser.add_argument('--excel-file', default=os.getenv('WINE_EXCEL_FILE', 'wine_price_list.xlsx'),
                       help=f'Путь к Excel файлу (по умолчанию: {os.getenv("WINE_EXCEL_FILE", "wine_price_list.xlsx")})')
    parser.add_argument('--template', default=os.getenv('WINE_TEMPLATE_FILE', 'template.html'),
                       help=f'Путь к HTML шаблону (по умолчанию: {os.getenv("WINE_TEMPLATE_FILE", "template.html")})')
    parser.add_argument('--output', default=os.getenv('WINE_OUTPUT_FILE', 'index.html'),
                       help=f'Путь для сохранения HTML (по умолчанию: {os.getenv("WINE_OUTPUT_FILE", "index.html")})')
    parser.add_argument('--foundation-year', type=int, default=int(os.getenv('WINE_FOUNDATION_YEAR', '1920')),
                       help='Год основания винодельни (по умолчанию: 1920)')
    return parser


def calculate_winery_age(foundation_year, current_year):
    return current_year - foundation_year


def get_year_word(years):
    if 11 <= years % 100 <= 14:
        return "лет"
    last_digit = years % 10
    return "год" if last_digit == 1 else "года" if 2 <= last_digit <= 4 else "лет"


def read_excel_file(file_path):
    if not Path(file_path).is_file():
        raise FileNotFoundError(f"Файл не найден: {file_path}")
    return pd.read_excel(file_path, na_values=['', ' ', 'N/A', 'NULL'], keep_default_na=False)


def validate_catalog_columns(catalog):
    required_columns = {'Категория', 'Название', 'Цена', 'Картинка'}
    missing = required_columns - set(catalog.columns)
    if missing:
        raise KeyError(f"Отсутствуют обязательные колонки: {', '.join(missing)}")


def extract_wine(row):
    return {
        'name': row['Название'],
        'grape_type': row.get('Сорт', ''),
        'price': row['Цена'],
        'image': row['Картинка'],
        'promotion': row.get('Акция', '')
    }


def group_by_category(catalog):
    grouped = defaultdict(list)
    for _, row in catalog.iterrows():
        wine = extract_wine(row)
        grouped[row['Категория']].append(wine)
    return grouped


def create_template_environment():
    return Environment(loader=FileSystemLoader('.'))


def render_template(env, template_filepath, context):
    return env.get_template(template_filepath).render(**context)


def save_html(content, output_filepath):
    with open(output_filepath, 'w', encoding='utf-8') as file:
        file.write(content)


def render_catalog_page(catalog, years, year_word, template_filepath, output_filepath):
    env = create_template_environment()
    context = {'winery_years': years, 'year_word': year_word, 'wines': catalog}
    html = render_template(env, template_filepath, context)
    save_html(html, output_filepath)


def count_wines(catalog):
    return sum(len(wines) for wines in catalog.values())


def print_category_stats(catalog):
    for category, wines in catalog.items():
        print(f"     - {category}: {len(wines)}")


def show_report(catalog, years, year_word, excel_filepath, template_filepath, output_filepath):
    print(f"\n📊 Отчет:")
    print(f"   • Винодельне: {years} {year_word}")
    print(f"   • Файл: {excel_filepath}")
    print(f"   • Шаблон: {template_filepath}")
    print(f"   • Результат: {output_filepath}")
    print("   • Вина по категориям:")
    print_category_stats(catalog)
    print(f"   • Всего: {count_wines(catalog)}")


def main():
    load_dotenv()

    parser = create_parser()
    args = parser.parse_args()

    current_year = datetime.now().year
    foundation_year = args.foundation_year
    excel_filepath = args.excel_file
    template_filepath = args.template
    output_filepath = args.output

    years = calculate_winery_age(foundation_year, current_year)
    year_word = get_year_word(years)

    print("🚀 Запуск генератора сайта")
    print(f"⚙️  Конфигурация: {{'excel': '{excel_filepath}', 'template': '{template_filepath}', "
          f"'output': '{output_filepath}', 'foundation_year': {foundation_year}, 'current_year': {current_year}}}")

    try:
        catalog_df = read_excel_file(excel_filepath)
        validate_catalog_columns(catalog_df)
        catalog = group_by_category(catalog_df)
        print("✅ Каталог вин загружен и сгруппирован")
    except FileNotFoundError as e:
        print(f"❌ Файл не найден: {e}")
        return
    except pd.errors.EmptyDataError:
        print("❌ Ошибка: Excel-файл пуст")
        return
    except pd.errors.ParserError as e:
        print(f"❌ Ошибка парсинга Excel: {e}")
        return
    except KeyError as e:
        print(f"❌ Ошибка структуры данных: {e}")
        return

    try:
        render_catalog_page(catalog, years, year_word, template_filepath, output_filepath)
        print("✅ HTML-страница успешно сгенерирована")
    except TemplateNotFound as e:
        print(f"❌ Шаблон не найден: {e}")
        return
    except (OSError, IOError) as e:
        print(f"❌ Ошибка ввода-вывода при записи результата: {e}")
        return

    show_report(catalog, years, year_word, excel_filepath, template_filepath, output_filepath)


if __name__ == "__main__":
    main()