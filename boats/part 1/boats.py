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
    driver.get('https://www.boats.net/catalog')

    final_list = []

    make_links = []
    make_names = []

    makes_catalog_table = driver.find_elements(By.CSS_SELECTOR,"div[class='catalog-table']")[1]
    makers = makes_catalog_table.find_elements(By.TAG_NAME,'a')
    for make in makers:
        if make.text == 'Mercruiser' or make.text == 'Mercury Sportjet' or make.text == 'Motorguide' or make.text == 'OMC':
            pass
        else:
            make_links.append(make.get_attribute('href'))
            make_names.append(make.text)


    for make_link, make_name in zip(make_links,make_names):

        driver.get(make_link)
        time.sleep(3)

        types_catalog_table = driver.find_element(By.CSS_SELECTOR,"div[class='catalog-table']")
        outboard_type = types_catalog_table.find_element(By.TAG_NAME,'a')

        type_text = outboard_type.text
        type_link = outboard_type.get_attribute('href')

        driver.get(type_link)
        time.sleep(3)
        try:
            horse_catalog_table = driver.find_element(By.CSS_SELECTOR, "div[class='catalog-table']")
            horse_powers = horse_catalog_table.find_elements(By.TAG_NAME, 'a')

            horse_values = []
            horse_links = []

            for horse_power in horse_powers:
                horse_values.append(horse_power.text)
                horse_links.append(horse_power.get_attribute('href'))

            for horse_value,horse_link in zip(horse_values,horse_links):

                driver.get(horse_link)
                time.sleep(3)

                try:
                    models_catalog_tables = driver.find_elements(By.CSS_SELECTOR, "div[class='catalog-table']")
                    for models_catalog_table in models_catalog_tables:
                        models = models_catalog_table.find_elements(By.TAG_NAME, 'a')
                        for model in models:
                            model_text = model.text
                            list_to_append = ["Boat",make_name,type_text + ' ' + horse_value,model_text]
                            print(list_to_append)
                            final_list.append(list_to_append)
                except:
                    pass
        except:
            pass


    workbook_name = 'Final_results_of_scraping.xlsx'
    wb = Workbook()
    page = wb.active
    headers = ['Type', 'Make', 'HorsePower', 'Model']
    page.append(headers)
    for row in final_list:
        page.append(row)
    wb.save(filename='Final_results_of_scraping.xlsx')












