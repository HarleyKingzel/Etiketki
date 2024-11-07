import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# Настройки
EXCEL_INPUT_PATH = 'Артикли.xlsx'  # Путь к вашему Excel-файлу с артиклями
EXCEL_OUTPUT_PATH = 'Данныеt.xlsx'  # Путь для сохранения результатов
WEBDRIVER_PATH = 'C://chromedriver'  # Укажите путь к вашему ChromeDriver
WEBSITE_URL = 'https://www.kirpich.ru/shop/'  # Укажите URL сайта

# Селекторы (необходимо заменить на актуальные для вашего сайта)
SEARCH_INPUT_SELECTOR = 'input#autocomplete'  # Селектор поисковой строки
SEARCH_BUTTON_SELECTOR = 'button#use'  # Селектор кнопки поиска (если есть)
RESULT_ITEM_SELECTOR = 'div.card_hover-area-wrapper'
PARAMETER_SELECTORS = {
    'Вид материала': 'single-tabs > div.tab-content.is-active > div > div.chars__column.chars__column--big > ul > li:nth-child(9)',
    'Название': 'single-tabs > div.tab-content.is-active > div > div.chars__column.chars__column--big > ul > li:nth-child(7) > span:nth-child(2)',
    'Пустотность': '#single-tabs > div.tab-content.is-active > div > div.chars__column.chars__column--big > ul > li:nth-child(14) > span:nth-child(2)',
    'Поверхность': '#single-tabs > div.tab-content.is-active > div > div.chars__column.chars__column--big > ul > li:nth-child(12) > span:nth-child(2)',
    'Размер': '#single-tabs > div.tab-content.is-active > div > div.chars__column.chars__column--big > ul > li:nth-child(5) > span:nth-child(2)',
    'Морозостойкость': '#single-tabs > div.tab-content.is-active > div > div.chars__column.chars__column--big > ul > li:nth-child(8) > span:nth-child(2)',
    'Марка': '#single-tabs > div.tab-content.is-active > div > div.chars__column.chars__column--big > ul > li:nth-child(2) > span:nth-child(2)',
    'Загрузка поддон': '#single-tabs > div.tab-content.is-active > div > div:nth-child(2) > ul > li:nth-child(3) > span:nth-child(2)',
    'Загрузка машины': '#single-tabs > div.tab-content.is-active > div > div:nth-child(2) > ul > li:nth-child(4) > span:nth-child(2)'
}

# Функция для инициализации веб-драйвера
def init_driver():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')  # Запуск без UI
    options.add_argument('--disable-gpu')
    driver = webdriver.Chrome(executable_path=WEBDRIVER_PATH, options=options)
    return driver

# Функция для обработки одного артикля
def process_article(driver, article):
    try:
        driver.get(WEBSITE_URL)
        time.sleep(2)  # Ожидание загрузки страницы
        # Ввод артикля в поисковую строку
        search_input = driver.find_element(By.CSS_SELECTOR, SEARCH_INPUT_SELECTOR)
        search_input.clear()
        search_input.send_keys(article)
        # Нажатие кнопки поиска, если требуется
        if SEARCH_BUTTON_SELECTOR:
            search_button = driver.find_element(By.CSS_SELECTOR, SEARCH_BUTTON_SELECTOR)
            search_button.click()
        else:
            search_input.send_keys(Keys.RETURN)
        time.sleep(3)  # Ожидание загрузки результатов поиска
        # Нажатие на первую найденную позицию
        first_result = driver.find_element(By.CSS_SELECTOR, RESULT_ITEM_SELECTOR)
        first_result.click()
        time.sleep(3)  # Ожидание загрузки страницы товара
        # Извлечение параметров
        data = {'Артикул': article}
        for param, selector in PARAMETER_SELECTORS.items():
            try:
                element = driver.find_element(By.CSS_SELECTOR, selector)
                data[param] = element.text.strip()
            except:
                data[param] = 'Не найдено'
        return data
    except Exception as e:
        print(f"Ошибка при обработке артикля {article}: {e}")
        return {'Артикул': article, **{param: 'Ошибка' for param in PARAMETER_SELECTORS}}

def main():
    # Чтение артиклей из Excel
    df = pd.read_excel(EXCEL_INPUT_PATH)
    # Проверка наличия столбца 'Артикул'
    if 'Артикул' not in df.columns:
        print("В Excel нет столбца 'Артикул'. Пожалуйста, проверьте файл.")
        return
    articles = df['Артикул'].dropna().unique()
    # Подготовка списка для результатов
    results = []
    # Инициализация веб-драйвера
    driver = init_driver()
    try:
        for idx, article in enumerate(articles, 1):
            print(f"Обработка {idx}/{len(articles)}: Артикул {article}")
            data = process_article(driver, str(article))
            results.append(data)
    finally:
        driver.quit()
    # Создание DataFrame с результатами
    results_df = pd.DataFrame(results)
    # Объединение исходного DataFrame с результатами по артиклю
    final_df = df.merge(results_df, on='Артикул', how='left')
    # Сохранение результатов в Excel
    final_df.to_excel(EXCEL_OUTPUT_PATH, index=False)
    print(f"Результаты сохранены в {EXCEL_OUTPUT_PATH}")


if __name__ == "__main__":
    main()
