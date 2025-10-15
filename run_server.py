import pandas as pd
from jinja2 import Environment, FileSystemLoader
from collections import defaultdict


def calculate_winery_age(foundation_year: int, current_year: int = 2023) -> int:
    """Рассчитывает возраст винодельни."""
    return current_year - foundation_year


def get_russian_year_word(number: int) -> str:
    """
    Возвращает правильную форму слова 'год' в зависимости от числа.
    
    Args:
        number: Число для которого нужно определить форму слова
        
    Returns:
        Строка с правильной формой: 'год', 'года' или 'лет'
    """
    if 11 <= number % 100 <= 14:
        return "лет"
    
    last_digit = number % 10
    if last_digit == 1:
        return "год"
    elif 2 <= last_digit <= 4:
        return "года"
    else:
        return "лет"


def load_wine_data(file_path: str) -> dict:
    """
    Загружает данные о винах из Excel файла и группирует по категориям.
    
    Args:
        file_path: Путь к Excel файлу
        
    Returns:
        Словарь с продуктами, сгруппированными по категориям
    """
    try:
        dataframe = pd.read_excel(file_path)
        print("✅ Файл успешно прочитан!")
        
        return group_products_by_category(dataframe)
        
    except FileNotFoundError:
        print(f"⚠️ Файл {file_path} не найден.")
        return {}
    except Exception as error:
        print(f"❌ Ошибка при чтении файла: {error}")
        return {}


def group_products_by_category(dataframe: pd.DataFrame) -> defaultdict:
    """
    Группирует продукты по категориям.
    
    Args:
        dataframe: DataFrame с данными о продуктах
        
    Returns:
        defaultdict с продуктами, сгруппированными по категориям
    """
    products_by_category = defaultdict(list)
    
    for _, row in dataframe.iterrows():
        product = create_product_dict(row)
        products_by_category[row['Категория']].append(product)
    
    return products_by_category


def create_product_dict(row: pd.Series) -> dict:
    """
    Создает словарь с информацией о продукте.
    
    Args:
        row: Строка DataFrame с данными о продукте
        
    Returns:
        Словарь с информацией о продукте
    """
    return {
        'name': row['Название'],
        'grape_type': row['Сорт'] if pd.notna(row['Сорт']) else '',
        'price': row['Цена'],
        'image': row['Картинка'],
        'promotion': row['Акция'] if pd.notna(row['Акция']) else ''
    }


def generate_website(products_data: dict, winery_age: int, year_word: str) -> None:
    """
    Генерирует HTML страницу на основе данных о продуктах.
    
    Args:
        products_data: Данные о продуктах
        winery_age: Возраст винодельни
        year_word: Правильная форма слова 'год'
    """
    environment = Environment(loader=FileSystemLoader('.'))
    template = environment.get_template('template.html')
    
    rendered_page = template.render(
        winery_age=winery_age,
        year_word=year_word,
        products_data=products_data
    )
    
    with open('index.html', 'w', encoding='utf-8') as file:
        file.write(rendered_page)


def print_statistics(products_data: dict, winery_age: int, year_word: str) -> None:
    """
    Выводит статистику по сгенерированному сайту.
    
    Args:
        products_data: Данные о продуктах
        winery_age: Возраст винодельни
        year_word: Правильная форма слова 'год'
    """
    print(f"\n✅ Сайт сгенерирован. Возраст винодельни: {winery_age} {year_word}.")
    print("✅ Добавлено товаров по категориям:")
    
    for category, products in products_data.items():
        print(f"   - {category}: {len(products)} товаров")


def main():
    """Основная функция программы."""
    # Конфигурационные параметры
    FOUNDATION_YEAR = 1920
    EXCEL_FILE_PATH = 'wine3.xlsx'
    
    # Расчет данных
    winery_age = calculate_winery_age(FOUNDATION_YEAR)
    year_word = get_russian_year_word(winery_age)
    
    # Загрузка и обработка данных
    products_data = load_wine_data(EXCEL_FILE_PATH)
    
    # Генерация сайта
    generate_website(products_data, winery_age, year_word)
    
    # Вывод статистики
    print_statistics(products_data, winery_age, year_word)


if __name__ == "__main__":
    main()