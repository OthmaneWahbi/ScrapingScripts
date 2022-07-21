from selenium.webdriver.common.by import By
import undetected_chromedriver as uc
import time
from openpyxl import Workbook

if __name__ == '__main__':
    uc.TARGET_VERSION = 104

    options = uc.ChromeOptions()
    #options.add_argument("--disable-extensions")
    options.binary_location = 'C:\Program Files\Google\Chrome Beta\Application\chrome.exe'

    driver = uc.Chrome(options=options)
    driver.maximize_window()
    driver.implicitly_wait(30)

    final_list = []

    # Part 1

    driver.get('https://www.boats.net/catalog/mercruiser')
    time.sleep(1)


    types_catalog_table = driver.find_element(By.CSS_SELECTOR, "div[class='catalog-table']")
    types = types_catalog_table.find_elements(By.TAG_NAME, 'a')

    types_links = []
    types_names = []

    types_to_scrape = ["Sterndrive","Inboard"]

    for type in types:
        if type.text in types_to_scrape:
            types_links.append(type.get_attribute('href'))
            types_names.append(type.text)

    for type_name,type_link in zip(types_names,types_links):
        driver.get(type_link)
        time.sleep(3)

        do_not_scrape = ["Transoms","Zues Pod Drive","Outdrives","Exhaust & Cooling Kits"]

        category_catalog_table = driver.find_element(By.CSS_SELECTOR, "div[class='catalog-table']")
        categories = category_catalog_table.find_elements(By.TAG_NAME, 'a')

        category_links = []
        category_names = []

        for category in categories:
            if category.text not in do_not_scrape:
                category_links.append(category.get_attribute('href'))
                category_names.append(category.text)

        for category_name, category_link in zip(category_names,category_links):
            horsepower = type_name + ' ' + category_name

            driver.get(category_link)

            sizes_catalog_table = driver.find_element(By.CSS_SELECTOR, "div[class='catalog-table']")
            sizes = sizes_catalog_table.find_elements(By.TAG_NAME, 'a')

            sizes_links = []
            sizes_names = []

            for size in sizes:
                sizes_links.append(size.get_attribute('href'))
                sizes_names.append(size.text)

            for size_name, size_link in zip(sizes_names,sizes_links):
                driver.get(size_link)
                time.sleep(3)
                try:
                    serial_range_catalog_table = driver.find_element(By.CSS_SELECTOR, "div[class='catalog-table']")
                    serial_ranges = serial_range_catalog_table.find_elements(By.TAG_NAME, 'a')

                    for serial_range in serial_ranges:
                        model_text = size_name + ' ' + serial_range.text
                        list_to_append = ['Boat', "Mercruiser", horsepower, model_text]
                        final_list.append(list_to_append)
                        print(list_to_append)
                except:
                    pass

    
    # Part 2

    driver.get("https://www.boats.net/catalog/mercruiser/sterndrive/outdrives")
    time.sleep(3)

    models_catalog_table = driver.find_element(By.CSS_SELECTOR, "div[class='catalog-table']")
    models = models_catalog_table.find_elements(By.TAG_NAME, 'a')

    for model in models:
        model_text = model.text
        list_to_append = ['Boat', "Mercruiser", "Sterndrive Outdrives", model_text]
        final_list.append(list_to_append)
        print(list_to_append)


    # Part 3
    driver.get("https://www.boats.net/catalog/mercury-sportjet/jet-drive")
    time.sleep(3)

    horsepower_catalog_table = driver.find_element(By.CSS_SELECTOR, "div[class='catalog-table']")
    horsepowers = horsepower_catalog_table.find_elements(By.TAG_NAME, 'a')

    horsepower_names = []
    horsepower_links = []

    for horsepower in horsepowers:
        horsepower_links.append(horsepower.get_attribute('href'))
        horsepower_names.append("Jet Drive " + horsepower.text)

    for horsepower_name,horsepower_link in zip(horsepower_names,horsepower_links):
        driver.get(horsepower_link)
        time.sleep(3)

        serial_ranges_catalog_tables = driver.find_elements(By.CSS_SELECTOR, "div[class='catalog-table']")

        for serial_ranges_catalog_table in serial_ranges_catalog_tables:
            serial_ranges = serial_ranges_catalog_table.find_elements(By.TAG_NAME, 'a')
            for serial_range in serial_ranges:
                model_text = serial_range.text
                list_to_append = ['Boat', "Mercury Sportjet", horsepower_name, model_text]
                print(list_to_append)
                final_list.append(list_to_append)


    
    #Part 4

    driver.get("https://www.boats.net/catalog/yamaha/sterndrive")
    time.sleep(3)

    model_catalog = driver.find_element(By.CSS_SELECTOR,"div[class='catalog-table']")
    models = model_catalog.find_elements(By.TAG_NAME,'a')

    for model in models:
        model_text = model.text
        list_to_append = ['Boat', "Yamaha", "Sterndrive", model_text]
        print(list_to_append)
        final_list.append(list_to_append)


    #Part 5

    driver.get("https://www.boats.net/catalog/yamaha/jet-drive")

    horsepower_catalog_table = driver.find_element(By.CSS_SELECTOR, "div[class='catalog-table']")
    horsepowers = horsepower_catalog_table.find_elements(By.TAG_NAME, 'a')

    horsepower_names = []
    horsepower_links = []

    for horsepower in horsepowers:
        horsepower_links.append(horsepower.get_attribute('href'))
        horsepower_names.append("Jet Drive " + horsepower.text)

    for horsepower_name, horsepower_link in zip(horsepower_names, horsepower_links):
        driver.get(horsepower_link)
        time.sleep(3)

        models_catalog_tables = driver.find_elements(By.CSS_SELECTOR, "div[class='catalog-table']")

        for models_catalog_table in models_catalog_tables:
            models = models_catalog_table.find_elements(By.TAG_NAME, 'a')
            for model in models:
                model_text = model.text
                list_to_append = ['Boat', "Yamaha", horsepower_name, model_text]
                print(list_to_append)
                final_list.append(list_to_append)

    workbook_name = 'Final_results_of_scraping.xlsx'
    wb = Workbook()
    page = wb.active
    headers = ['Type', 'Make', 'HorsePower', 'Model']
    page.append(headers)
    for row in final_list:
        page.append(row)
    wb.save(filename='Final_results_of_scraping.xlsx')




