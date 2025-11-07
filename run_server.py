import pandas as pd
from jinja2 import Environment, FileSystemLoader, TemplateNotFound
from collections import defaultdict
from dotenv import load_dotenv
from datetime import datetime
import argparse
import os


def get_year_word(n: int) -> str:
    if 11 <= n % 100 <= 14:
        return "лет"
    return ["год", "года", "лет"][(n % 10) // 2 if 2 <= n % 10 <= 4 else min(n % 10, 2)]


def main():
    load_dotenv()
    parser = argparse.ArgumentParser(description='Генератор сайта винного магазина')
    parser.add_argument('--excel-file', default=os.getenv('WINE_EXCEL_FILE', 'wine_price_list.xlsx'))
    parser.add_argument('--template', default=os.getenv('WINE_TEMPLATE_FILE', 'template.html'))
    parser.add_argument('--output', default=os.getenv('WINE_OUTPUT_FILE', 'index.html'))
    parser.add_argument('--foundation-year', type=int, default=int(os.getenv('WINE_FOUNDATION_YEAR', '1920')))
    args = parser.parse_args()

    current_year = datetime.now().year
    years = current_year - args.foundation_year
    year_word = get_year_word(years)

    print("🚀 Запуск генератора сайта")
    print(f"⚙️  Конфигурация: {{'excel': '{args.excel_file}', 'template': '{args.template}', "
          f"'output': '{args.output}', 'foundation_year': {args.foundation_year}, 'current_year': {current_year}}}")

    try:
        df = pd.read_excel(
            args.excel_file,
            na_values=['', ' ', 'N/A', 'NULL'],
            keep_default_na=False
        )
    except FileNotFoundError:
        print(f"❌ Файл не найден: {args.excel_file}")
        return
    except pd.errors.EmptyDataError:
        print("❌ Ошибка: Excel-файл пуст")
        return
    except pd.errors.ParserError as e:
        print(f"❌ Ошибка парсинга Excel: {e}")
        return

    required = {'Категория', 'Название', 'Цена', 'Картинка'}
    if missing := required - set(df.columns):
        print(f"❌ Отсутствуют обязательные колонки: {', '.join(missing)}")
        return

    wines = defaultdict(list)
    for _, row in df.iterrows():
        wines[row['Категория']].append({
            'name': row['Название'],
            'grape_type': row.get('Сорт', ''),
            'price': row['Цена'],
            'image': row['Картинка'],
            'promotion': row.get('Акция', '')
        })
    print("✅ Каталог вин загружен и сгруппирован")

    try:
        env = Environment(loader=FileSystemLoader('.'))
        template = env.get_template(args.template)
        html = template.render(winery_years=years, year_word=year_word, wines=wines)
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(html)
        print("✅ HTML-страница успешно сгенерирована")
    except TemplateNotFound as e:
        print(f"❌ Шаблон не найден: {e}")
        return
    except (OSError, IOError) as e:
        print(f"❌ Ошибка записи: {e}")
        return

    total = sum(len(items) for items in wines.values())
    print(f"\n📊 Отчет:")
    print(f"   • Винодельне: {years} {year_word}")
    print(f"   • Файл: {args.excel_file}")
    print(f"   • Шаблон: {args.template}")
    print(f"   • Результат: {args.output}")
    print("   • Вина по категориям:")
    for cat, items in wines.items():
        print(f"     - {cat}: {len(items)}")
    print(f"   • Всего: {total}")


if __name__ == "__main__":
    main()