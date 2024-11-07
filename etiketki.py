import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from collections import defaultdict

# Настройки
EXCEL_INPUT_FILE = 'Артикли.xlsx'        # Входной Excel файл
EXCEL_OUTPUT_FILE = 'Результаты.xlsx'    # Выходной Excel файл
SEARCH_URL = 'https://kirpich.ru/shop'

# Чтение данных из Excel
def read_articles(file_path):
    df = pd.read_excel(file_path)
    # Предполагаем, что нужные артикулы находятся в первом столбце
    articles = df.iloc[:, 0].dropna().astype(str).tolist()
    return articles

# Запись данных в Excel
def write_results(data, file_path):
    df = pd.DataFrame(data)
    df.to_excel(file_path, index=False)

# Инициализация веб-драйвера
def init_driver():
    options = webdriver.ChromeOptions()
    #options.add_argument('--headless')  # Запуск без интерфейса браузера
    options.add_argument('--disable-gpu')
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
    return driver

# Основная функция
def main():
    articles = read_articles(EXCEL_INPUT_FILE)
    driver = init_driver()
    results = []
    all_specs = set()  # Для сбора всех возможных спецификаций

    for idx, article in enumerate(articles, 1):
        try:
            print(f"Обработка {idx}/{len(articles)}: {article}")
            driver.get(SEARCH_URL)
            time.sleep(2)  # Ждем загрузки страницы

            # Найти поле поиска
            search_box = driver.find_element(By.NAME, 'q')  # Обновите селектор, если необходимо
            search_box.clear()
            search_box.send_keys(article)
            search_box.send_keys(Keys.RETURN)  # Нажать Enter
            time.sleep(3)  # Ждем результатов поиска

            # Выбрать первый продукт в результатах
            first_product = driver.find_element(By.CSS_SELECTOR, '.product-item a')  # Обновите селектор согласно сайту
            first_product.click()
            time.sleep(3)  # Ждем загрузки страницы продукта

            # Извлечь данные из списка <li>
            li_elements = driver.find_elements(By.CSS_SELECTOR, 'ul.specs-list li')  # Обновите селектор
            specs = {}
            for li in li_elements:
                text = li.text
                if ':' in text:
                    key, value = map(str.strip, text.split(':', 1))
                    specs[key] = value
                    all_specs.add(key)
                else:
                    # Если формат неизвестен, можно сохранить как "Спецификация X"
                    specs[text] = None

            # Добавить артикул и спецификации
            result = {'Артикул': article}
            result.update(specs)
            results.append(result)

        except Exception as e:
            print(f"Ошибка при обработке артикла {article}: {e}")
            result = {'Артикул': article, 'Ошибка': str(e)}
            results.append(result)

    # Закрыть драйвер
    driver.quit()

    # Создать DataFrame с учетом всех спецификаций
    df = pd.DataFrame(results)

    # Переупорядочить столбцы: сначала 'Артикул', затем спецификации в алфавитном порядке или нужном порядке
    cols = ['Артикул'] + sorted(all_specs)
    if 'Ошибка' in df.columns:
        cols.append('Ошибка')
    df = df.reindex(columns=cols)

    # Записать результаты в Excel
    write_results(df, EXCEL_OUTPUT_FILE)
    print(f"Завершено. Результаты сохранены в {EXCEL_OUTPUT_FILE}")

if __name__ == "__main__":
    main()
