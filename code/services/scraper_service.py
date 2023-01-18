from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from datetime import date, timedelta
from pathlib import Path

from constants import BASE_URL, TABLE_CLASS_NAME, TIME_PERIOD, BASE_DIR


class RateStats:

    def __init__(self, driver):
        self.driver = driver
    
    def get_page_by_url(self, url = BASE_URL):
        self.driver.get(url)
        
    def get_table(self):
        delay = 5 
        try:
            myElem = WebDriverWait(self.driver, delay).until(EC.presence_of_element_located((By.CLASS_NAME,  TABLE_CLASS_NAME)))
            return myElem
        except TimeoutException:
            raise TimeoutException()

    def get_info_from_table(self, table):
        lined_text = table.text.split('\n')    
        return [text.split(' ') for text in lined_text]
        
    def get_table_data(self, table):
        data = []
        for row in table.find_elements(By.TAG_NAME, 'tr'):
            row_data = []
            for cell in row.find_elements(By.TAG_NAME, 'td'):
                row_data.append(cell.text)
            if row_data:
                data.append(row_data)
        return data

    def get_table_head(self, table):
        return [ t.text for t in table.find_elements(By.TAG_NAME, 'th')]
        
    def add_data_to_table_data(self, data, date):
        for d in data:
            d.append(date)
        return data

    def add_title_to_table_head(self, data, title):
        data.append(title)
        return data
    
    def close(self):
        self.driver.quit()

def scrap(period: int = TIME_PERIOD):
    options = Options()
    options.add_argument('--headless')

    driver_path = Path(BASE_DIR / 'drivers' / 'chromedriver.exe')
    print(driver_path)

    with webdriver.Chrome(executable_path=driver_path, options=options) as driver:

        print('INFO:    ', 'Сбор данных')
        scraper = RateStats(driver)
        today = date.today()
        res_head = []
        res_data = []
        for i in range(period):
            past_day = today - timedelta(days=i)
            url_path = ''.join(str(past_day).split('-'))
            scraper.get_page_by_url(BASE_URL+url_path+'/')

            table = scraper.get_table()

            table_data = scraper.get_table_data(table)
            table_data = scraper.add_data_to_table_data(table_data, str(past_day))

            res_data = [*res_data, *table_data]

            if not res_head:
                table_head = scraper.get_table_head(table)
                table_head = scraper.add_title_to_table_head(table_head, 'Дата')

                res_head = table_head
            print('INFO:    ', f'Страница {i+1} ({past_day}) готова')
        
    return res_head, res_data
