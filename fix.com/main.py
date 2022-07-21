from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import undetected_chromedriver as uc
from selenium.webdriver.common.keys import Keys
import time
from openpyxl import Workbook

if __name__ == '__main__':
    #uc.TARGET_VERSION = 104
    options = Options()
    options.binary_location ='C:\Program Files\Google\Chrome Beta\Application\chrome.exe'
    driver = uc.Chrome(options=options)
    driver.maximize_window()
    driver.implicitly_wait(30)

    final_list = []

    driver.get("https://www.fix.com/parts/appliance/range-hood/")

    brands_section = driver.find_element(By.CSS_SELECTOR,"ul[class='nf__links']")
    brands = brands_section.find_elements(By.TAG_NAME,'a')

    brands_names = []
    brands_links = []

    for brand in brands:
        brands_links.append(brand.get_attribute('href')+'models/')
        brand_name = brand.text.split(" Range Hood Parts")[0]
        brands_names.append(brand_name)

    for brand_name, brand_link in zip(brands_names,brands_links):
        driver.get(brand_link)
        time.sleep(2)
        try:
            models_section = driver.find_element(By.CSS_SELECTOR, "ul[class='nf__links']")
            models = models_section.find_elements(By.TAG_NAME, 'a')

            for model in models:
                model_text = model.text
                list_to_append = ["Cooker Hood",brand_name,model_text]
                final_list.append(list_to_append)
                print(list_to_append)
        except:
            pass


    time.sleep(2)

    #part 2
    driver.get("https://www.fix.com/parts/appliance/cooktop/")

    brands_section = driver.find_element(By.CSS_SELECTOR, "ul[class='nf__links']")
    brands = brands_section.find_elements(By.TAG_NAME, 'a')

    brands_names = []
    brands_links = []

    for brand in brands:
        brands_links.append(brand.get_attribute('href') + 'models/')
        brand_name = brand.text.split(" Cooktop Parts")[0]
        brands_names.append(brand_name)

    for brand_name, brand_link in zip(brands_names, brands_links):
        driver.get(brand_link)
        time.sleep(2)
        try:
            models_section = driver.find_element(By.CSS_SELECTOR, "ul[class='nf__links']")
            models = models_section.find_elements(By.TAG_NAME, 'a')

            for model in models:
                model_text = model.text
                list_to_append = ["Stove Top", brand_name, model_text]
                final_list.append(list_to_append)
                print(list_to_append)
        except:
            pass





    workbook_name = 'Final_results_of_scraping.xlsx'
    wb = Workbook()
    page = wb.active
    headers = ['Type', 'Make', 'Model']
    page.append(headers)
    for row in final_list:
        page.append(row)
    wb.save(filename='Final_results_of_scraping.xlsx')

























