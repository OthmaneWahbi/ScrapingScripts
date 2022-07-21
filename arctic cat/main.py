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

    final_list = []

    driver.get('https://www.arcticcatpartshouse.com/')
    time.sleep(2)

    types_containers = driver.find_elements(By.CSS_SELECTOR,"div[class='brand-links']")
    types_names = []
    types_links = []

    for types_container in types_containers:
        button = types_container.find_element(By.CSS_SELECTOR,"a[class='button']")
        button_text = button.text
        types_links.append(button.get_attribute('href'))
        if "atv" in button_text.lower():
            types_names.append("ATV")
        elif "utv" in button_text.lower():
            types_names.append("UTV")
        elif "snow" in button_text.lower():
            types_names.append("Snowmobile")

    print(types_names)
    print(types_links)

    for type_name,type_link in zip(types_names,types_links):
        driver.get(type_link)
        time.sleep(3)

        years = driver.find_elements(By.CSS_SELECTOR,"a[class='pjq']")

        years_links = []
        years_values = []

        for year in years:
            years_values.append(year.text.split()[0])
            years_links.append(year.get_attribute('href'))

        for year_value, year_link in zip(years_values,years_links):
            driver.get(year_link)
            time.sleep(1.5)

            models = driver.find_elements(By.CSS_SELECTOR,"a[class='pjq']")
            for model in models:
                model_text = model.text
                if model_text != '':
                    list_to_append = [type_name,"Arctic Cat",year_value,model_text]
                    final_list.append(list_to_append)
                    print(list_to_append)


    workbook_name = 'Final_results_of_scraping.xlsx'
    wb = Workbook()
    page = wb.active
    headers = ['Type', 'Make', 'Year', 'Model']
    page.append(headers)
    for row in final_list:
        page.append(row)
    wb.save(filename='Final_results_of_scraping.xlsx')














