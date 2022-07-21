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
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    driver.implicitly_wait(30)

    final_list = []

    #PART 1
    list_of_makers = ["CFMOTO","Tracker"]
    driver.get('https://www.powersportsid.com/atv/')

    year_button = driver.find_element(By.CSS_SELECTOR, "div[data-placeholder='Year']")
    make_button = driver.find_element(By.CSS_SELECTOR, "div[data-placeholder='Make']")
    model_button = driver.find_element(By.CSS_SELECTOR, "div[data-placeholder='Model']")

    year_button.click()
    time.sleep(2)
    years = driver.find_elements(By.CSS_SELECTOR, "li[class='item ']")


    for year in years:
        year_button.send_keys(Keys.ARROW_DOWN)
        year_value = year_button.find_element(By.CSS_SELECTOR,"small[class='value']").text
        #print(year_value)

        time.sleep(2)
        make_button.click()
        time.sleep(2)
        makers = driver.find_elements(By.CSS_SELECTOR, "li[class='item ']")

        for maker in makers:
            make_button.send_keys(Keys.ARROW_DOWN)
            make_value = make_button.find_element(By.CSS_SELECTOR,"small[class='value']").text
            #print(make_value)

            if make_value in list_of_makers:
                time.sleep(2)
                model_button.click()
                time.sleep(2)
                models = driver.find_elements(By.CSS_SELECTOR, "li[class='item ']")

                for model in models:
                    model_value = model.get_attribute('innerHTML')
                    #print(model_value)
                    #time.sleep(1)
                    list_to_append = ["ATV",make_value,year_value,model_value]
                    final_list.append(list_to_append)
                    print(list_to_append)

    # PART 2
    list_of_makers = ["Kubota", "Tracker","KYMCO","Cub Cadet","Bobcat","CFMOTO","New Holland","John Deere"]
    time.sleep(1)
    driver.get('https://www.powersportsid.com/utv/')

    year_button = driver.find_element(By.CSS_SELECTOR, "div[data-placeholder='Year']")
    make_button = driver.find_element(By.CSS_SELECTOR, "div[data-placeholder='Make']")
    model_button = driver.find_element(By.CSS_SELECTOR, "div[data-placeholder='Model']")

    year_button.click()
    time.sleep(2)
    years = driver.find_elements(By.CSS_SELECTOR, "li[class='item ']")

    for year in years:
        year_button.send_keys(Keys.ARROW_DOWN)
        year_value = year_button.find_element(By.CSS_SELECTOR, "small[class='value']").text
        # print(year_value)

        time.sleep(2)
        make_button.click()
        time.sleep(2)
        makers = driver.find_elements(By.CSS_SELECTOR, "li[class='item ']")

        for maker in makers:
            make_button.send_keys(Keys.ARROW_DOWN)
            make_value = make_button.find_element(By.CSS_SELECTOR, "small[class='value']").text
            # print(make_value)

            if make_value in list_of_makers:
                time.sleep(3)
                model_button.click()
                time.sleep(3)
                models = driver.find_elements(By.CSS_SELECTOR, "li[class='item ']")

                for model in models:
                    model_value = model.get_attribute('innerHTML')
                    # print(model_value)
                    #time.sleep(1)
                    list_to_append = ["UTV", make_value, year_value, model_value]
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





