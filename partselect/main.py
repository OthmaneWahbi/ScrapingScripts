from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import undetected_chromedriver as uc
import time
from openpyxl import Workbook

if __name__ == '__main__':
    options = Options()
    options.binary_location ='C:\Program Files\Google\Chrome Beta\Application\chrome.exe'
    driver = uc.Chrome(options=options)
    driver.maximize_window()
    driver.implicitly_wait(10)

    final_list = []

    #Part 1

    driver.get("https://www.partselect.com/Freezer-Parts.htm")
    time.sleep(2)

    brands_section = driver.find_element(By.CSS_SELECTOR, "ul[class='nf__links']")
    brands = brands_section.find_elements(By.TAG_NAME, 'a')

    brands_names = []
    brands_links = []

    for brand in brands:
        link = "https://www.partselect.com/"+brand.text.replace("Parts","Models").replace(' ','-')+".htm"
        brands_links.append(link)

        brand_name = brand.text.replace(" Freezer Parts","")
        if brand_name == "GE":
            brand_name = "General Electric"
        brands_names.append(brand_name)

    for brand_name, brand_link in zip(brands_names,brands_links):
        driver.get(brand_link)
        time.sleep(1)

        test = True
        while test == True:
            try:
                models_section = driver.find_element(By.CSS_SELECTOR, "ul[class='nf__links']")
                models = models_section.find_elements(By.TAG_NAME, 'a')
                for model in models:
                    model_text = model.text.replace(brand_name + " Freezer",'')
                    if brand_name == "White-Westinghouse":
                        model_text = model.text.replace(" Westinghouse Freezer",'')
                    list_to_append = ["Freezer",brand_name,model_text.strip()]
                    final_list.append(list_to_append)
                    print(list_to_append)
                try:
                    next_button = driver.find_element(By.CSS_SELECTOR,"li[class='next']")
                    next_page = next_button.find_element(By.TAG_NAME,'a').get_attribute('href')
                    driver.get(next_page)
                    time.sleep(1)
                except:
                    test = False
            except:
                test = False


    #Part 2

    driver.get("https://www.partselect.com/Trash-Compactor-Parts.htm")
    time.sleep(2)

    brands_section = driver.find_element(By.CSS_SELECTOR, "ul[class='nf__links']")
    brands = brands_section.find_elements(By.TAG_NAME, 'a')

    brands_names = []
    brands_links = []

    for brand in brands:
        link = "https://www.partselect.com/" + brand.text.replace("Parts", "Models").replace(' ', '-') + ".htm"
        brands_links.append(link)

        brand_name = brand.text.replace(" Trash Compactor Parts", "")
        if brand_name == "GE":
            brand_name = "General Electric"
        brands_names.append(brand_name)

    for brand_name, brand_link in zip(brands_names, brands_links):
        driver.get(brand_link)
        time.sleep(1)

        test = True
        while test == True:
            try:
                models_section = driver.find_element(By.CSS_SELECTOR, "ul[class='nf__links']")
                models = models_section.find_elements(By.TAG_NAME, 'a')
                for model in models:
                    model_text = model.text.replace(brand_name + " Trash Compactor", '')
                    list_to_append = ["Trash Compactor", brand_name, model_text.strip()]
                    final_list.append(list_to_append)
                    print(list_to_append)
                try:
                    next_button = driver.find_element(By.CSS_SELECTOR, "li[class='next']")
                    next_page = next_button.find_element(By.TAG_NAME, 'a').get_attribute('href')
                    driver.get(next_page)
                    time.sleep(1)
                except:
                    test = False
            except:
                test = False

    workbook_name = 'Final_results_of_scraping.xlsx'
    wb = Workbook()
    page = wb.active
    headers = ['Type', 'Make', 'Model']
    page.append(headers)
    for row in final_list:
        page.append(row)
    wb.save(filename='Final_results_of_scraping.xlsx')
