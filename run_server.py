import pandas as pd
from jinja2 import Environment, FileSystemLoader, TemplateNotFound
from collections import defaultdict
import argparse
import os
from pathlib import Path
from dotenv import load_dotenv

def create_parser():
    parser = argparse.ArgumentParser(description='Генератор сайта винного магазина')
    parser.add_argument('--excel-file', default=None,
                       help='Путь к Excel файлу')
    parser.add_argument('--template', default=None,
                       help='Путь к HTML шаблону')
    parser.add_argument('--output', default=None,
                       help='Путь для сохранения HTML')
    parser.add_argument('--foundation-year', type=int, default=None,
                       help='Год основания винодельни')
    parser.add_argument('--current-year', type=int, default=None,
                       help='Текущий год')
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

def count_wines(catalog):
    return sum(len(wines) for wines in catalog.values())

def print_category_stats(catalog):
    for category, wines in catalog.items():
        print(f"     - {category}: {len(wines)}")

def render_catalog_page(catalog, years, year_word, template_path, output_path):
    env = Environment(loader=FileSystemLoader('.'))
    context = {'winery_years': years, 'year_word': year_word, 'wines': catalog}
    html = env.get_template(template_path).render(**context)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

def main():
    load_dotenv()

    foundation_year = int(os.getenv('WINE_FOUNDATION_YEAR', '1920'))
    excel_file = os.getenv('WINE_EXCEL_FILE', 'wine_price_list.xlsx')
    template = os.getenv('WINE_TEMPLATE_FILE', 'template.html')
    output = os.getenv('WINE_OUTPUT_FILE', 'index.html')
    current_year = int(os.getenv('WINE_CURRENT_YEAR', '2023'))

    parser = create_parser()
    args = parser.parse_args()

    excel_file = args.excel_file or excel_file
    template = args.template or template
    output = args.output or output
    foundation_year = args.foundation_year or foundation_year
    current_year = args.current_year or current_year

    print("🚀 Запуск генератора сайта")
    print(f"⚙️  Конфигурация: excel={excel_file}, template={template}, output={output}, "
          f"foundation={foundation_year}, current={current_year}")

    try:
        catalog_df = load_catalog(excel_file)
        catalog = group_by_category(catalog_df)

        years = calculate_winery_age(foundation_year, current_year)
        year_word = get_year_word(years)

        render_catalog_page(catalog, years, year_word, template, output)

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
    except TemplateNotFound as e:
        print(f"❌ Шаблон не найден: {e}")
        return
    except (OSError, IOError) as e:
        print(f"❌ Ошибка ввода-вывода: {e}")
        return
    except KeyboardInterrupt:
        print("\n⏹️  Выполнение прервано пользователем")
        return

    print("✅ Каталог вин загружен и сгруппирован")
    print("✅ HTML-страница успешно сгенерирована")

    print(f"\n📊 Отчет:")
    print(f"   • Винодельне: {years} {year_word}")
    print(f"   • Файл: {excel_file}")
    print(f"   • Шаблон: {template}")
    print(f"   • Результат: {output}")
    print("   • Вина по категориям:")
    print_category_stats(catalog)
    print(f"   • Всего: {count_wines(catalog)}")

if __name__ == "__main__":
    main()
