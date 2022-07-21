from selenium.webdriver.common.by import By
import undetected_chromedriver as uc
import time
from openpyxl import Workbook

if __name__ == '__main__':
    uc.TARGET_VERSION = 104

    options = uc.ChromeOptions()
    options.add_argument("--disable-extensions")
    options.binary_location = 'C:\Program Files\Google\Chrome Beta\Application\chrome.exe'
    driver = uc.Chrome(options=options)
    driver.maximize_window()
    driver.implicitly_wait(30)

    driver.get('https://www.volvopentastore.com/Engine-Part-Catalog/dm/cart_id.656502229--func.volvo--session_id.495183777--store_id.366')
    time.sleep(2)

    final_list = []

    categories = driver.find_elements(By.CSS_SELECTOR,"div[class='menu_open']")

    category_data_containers = driver.find_elements(By.CSS_SELECTOR,"div[class='category_data']")

    #IF YES
    category_data_containers += driver.find_elements(By.CSS_SELECTOR,"div[class='category_data list flex']")



    for category, category_data in zip(categories,category_data_containers):
        category_title = category.get_attribute('title').split(' (')[0]

        category.click()
        time.sleep(1.5)

        models = category_data.find_elements(By.TAG_NAME, 'a')
        for model in models:
            model_text = model.get_attribute('title')
            #IF YES
            if category_title == '':
                category_title = category.text
                if category_title == "Drives & Transmissions":
                    pass
                else:
                    final_models = model_text.split(',')
                    for final_model in final_models:
                        list_to_append = ["Boat","Volvo Penta",category_title,final_model.strip()]
                        print(list_to_append)
                        final_list.append(list_to_append)
            else:
                if category_title == "Drives & Transmissions":
                    pass
                else:
                    final_models = model_text.split(',')
                    for final_model in final_models:
                        list_to_append = ["Boat", "Volvo Penta", category_title, final_model.strip()]
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



