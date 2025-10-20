import pandas as pd
from jinja2 import Environment, FileSystemLoader, TemplateNotFound
from collections import defaultdict
import argparse
import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

FOUNDATION_YEAR = int(os.getenv('WINE_FOUNDATION_YEAR', '1920'))
EXCEL_FILE_PATH = os.getenv('WINE_EXCEL_FILE', 'wine_price_list.xlsx')
TEMPLATE_FILE = os.getenv('WINE_TEMPLATE_FILE', 'template.html')
OUTPUT_FILE = os.getenv('WINE_OUTPUT_FILE', 'index.html')
CURRENT_YEAR = int(os.getenv('WINE_CURRENT_YEAR', '2023'))


def create_parser():
    parser = argparse.ArgumentParser(description='Генератор сайта винного магазина')
    parser.add_argument('--excel-file', default=EXCEL_FILE_PATH,
                       help=f'Путь к Excel файлу (по умолчанию: {EXCEL_FILE_PATH})')
    parser.add_argument('--template', default=TEMPLATE_FILE,
                       help=f'Путь к HTML шаблону (по умолчанию: {TEMPLATE_FILE})')
    parser.add_argument('--output', default=OUTPUT_FILE,
                       help=f'Путь для сохранения HTML (по умолчанию: {OUTPUT_FILE})')
    parser.add_argument('--foundation-year', type=int, default=FOUNDATION_YEAR,
                       help=f'Год основания винодельни (по умолчанию: {FOUNDATION_YEAR})')
    parser.add_argument('--current-year', type=int, default=CURRENT_YEAR,
                       help=f'Текущий год (по умолчанию: {CURRENT_YEAR})')
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


def load_catalog(file_path):
    catalog = read_excel_file(file_path)
    validate_catalog_columns(catalog)
    return catalog


def extract_wine(row):
    return {
        'name': row['Название'],
        'grape_type': row['Сорт'],
        'price': row['Цена'],
        'image': row['Картинка'],
        'promotion': row['Акция']
    }


def group_by_category(catalog):
    grouped = defaultdict(list)
    for _, row in catalog.iterrows():
        wine = extract_wine(row)
        grouped[row['Категория']].append(wine)
    return grouped


def create_template_environment():
    return Environment(loader=FileSystemLoader('.'))


def render_template(env, template_path, context):
    return env.get_template(template_path).render(**context)


def save_html(content, output_path):
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(content)


def render_catalog_page(catalog, years, year_word, template_path, output_path):
    env = create_template_environment()
    context = {'winery_years': years, 'year_word': year_word, 'wines': catalog}
    html = render_template(env, template_path, context)
    save_html(html, output_path)


def count_wines(catalog):
    return sum(len(wines) for wines in catalog.values())


def print_category_stats(catalog):
    for category, wines in catalog.items():
        print(f"     - {category}: {len(wines)}")


def show_report(catalog, years, year_word, config):
    print(f"\n📊 Отчет:")
    print(f"   • Винодельне: {years} {year_word}")
    print(f"   • Файл: {config['excel_file']}")
    print(f"   • Шаблон: {config['template']}")
    print(f"   • Результат: {config['output']}")
    print("   • Вина по категориям:")
    print_category_stats(catalog)
    print(f"   • Всего: {count_wines(catalog)}")


def main():
    parser = create_parser()
    args = parser.parse_args()

    config = {
        'excel_file': args.excel_file,
        'template': args.template,
        'output': args.output,
        'foundation_year': args.foundation_year,
        'current_year': args.current_year
    }

    print("🚀 Запуск генератора сайта")
    print(f"⚙️  Конфигурация: {config}")

    try:
        years = calculate_winery_age(config['foundation_year'], config['current_year'])
        year_word = get_year_word(years)

        catalog_df = load_catalog(config['excel_file'])
        catalog = group_by_category(catalog_df)
        print("✅ Каталог вин загружен и сгруппирован")

        render_catalog_page(catalog, years, year_word, config['template'], config['output'])
        print("✅ HTML-страница успешно сгенерирована")

        show_report(catalog, years, year_word, config)

    except FileNotFoundError as e:
        print(f"❌ Файл не найден: {e}")
    except pd.errors.EmptyDataError:
        print("❌ Ошибка: Excel-файл пуст")
    except pd.errors.ParserError as e:
        print(f"❌ Ошибка парсинга Excel: {e}")
    except KeyError as e:
        print(f"❌ Ошибка структуры данных: {e}")
    except TemplateNotFound as e:
        print(f"❌ Шаблон не найден: {e}")
    except (OSError, IOError) as e:
        print(f"❌ Ошибка ввода-вывода: {e}")
    except KeyboardInterrupt:
        print("\n⏹️  Выполнение прервано пользователем")


if __name__ == "__main__":
    main()
