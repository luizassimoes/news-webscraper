import os
import time
import logging
from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from RPA.core.webdriver import start


class NewsWebScraper:

    def __init__(self):
        self.driver = None
        self.logger = logging.getLogger(__name__)

    def set_chrome_options(self):
        options = webdriver.ChromeOptions()
        # options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-gpu")
        options.add_argument('--disable-web-security')
        options.add_argument("--start-maximized")
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        return options

    def set_webdriver(self, browser="Chrome"):
        options = self.set_chrome_options()
        try:
            self.driver = start(browser, options=options)
            self.logger.info("WebDriver started successfully.")
        except Exception as e:
            self.logger.error(f"ERR0R set_webdriver() | Failed to start WebDriver: {e}")

    def open_url(self, url: str):
        if self.driver:
            try:
                self.driver.get(url)
                self.logger.info(f"Opened URL: {url}")
            except Exception as e:
                self.logger.error(f"ERROR open_url() | Failed to open URL {url}: {e}")
        else:
            self.logger.error("ERR0R open_url() | WebDriver not initialized. You must call set_webdriver() first.")

    def search(self, search_query: str):
        try:
            search_button = self.driver.find_element(By.XPATH, '/html/body/ps-header/header/div[2]/button')
            search_button.click()
            self.logger.info(f"Search button clicked.")
            search_field = self.driver.find_element(By.XPATH, '/html/body/ps-header/header/div[2]/div[2]/form/label/input')
            search_field.send_keys(search_query)
            search_field.submit()
            self.logger.info(f"Search query submitted.")
        except Exception as e:
            self.logger.error(f"ERROR search() | Could not find element: {e}")

    def sort_by_newest(self):
        wait = WebDriverWait(self.driver, 100)
        try:
            retries = 2
            for _ in range(retries):
                sort_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.select-input')))
                select = Select(sort_element)
                select.select_by_visible_text('Newest')
            self.logger.info(f"Query sorted by Newest.")
        except Exception as e:
            self.logger.error(f"ERROR sort_by_newest() | Error during selecting option: {e}")
        
    def get_element_list(self, element_selector, src=False):
        wait = WebDriverWait(self.driver, 100)
        retries = 3
        get_element = element_selector.split('-')[-1].split(' ')[0].title()
        for attempt in range(retries):
            try: 
                elements = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, element_selector)))
                self.logger.info(f"{get_element} gotten on attempt {attempt+1}.")
                if not src:
                    return [element.text for element in elements]
                else:
                    return [element.get_attribute('src') for element in elements]
            except Exception as e:
                self.logger.warning(f"Could not get {get_element} on attempt {attempt+1}. Retrying...")
                if attempt == retries - 1:
                    self.logger.error(f"ERROR get_element_list(): {get_element} | Error: {e}")

    def get_news_titles(self):
        return self.get_element_list('h3.promo-title a')

    def get_news_dates(self):
        return self.get_element_list('p.promo-timestamp')
    
    def get_news_descriptions(self):
        return self.get_element_list('p.promo-description')
    
    def get_news_pic_filenames(self):
        return self.get_element_list('picture img.image', src=True)

    def count_search_query(self, query, titles, descriptions):
        query_count = []
        for i in range(len(titles)):
            title_count = titles[i].count(query)
            description_count = descriptions[i].count(query)
            query_count.append(title_count + description_count)
        self.logger.info(f"Search query in Title counted.")
        return query_count
    
    def title_contains_money(self, titles):
        money_patterns = [
            r'\$\d{1,3}(,\d{3})*(\.\d{2})?',      # Pattern for $11.1 or $111,111.11
            r'\b\d+(\.\d{2})?\s?(dollars|USD)\b'  # Pattern for 11 dollars or 11 USD
        ]
        money_in_query = []
        for title in titles:
            money = False
            for pattern in money_patterns:
                if re.search(pattern, title, re.IGNORECASE):
                    money = True
                    break
            money_in_query.append(money)
        self.logger.info(f"Contains money analyzed.")
        return money_in_query
            

    def get_news(self):
        titles = self.get_news_titles()
        descriptions = self.get_news_descriptions()
        dates = self.get_news_dates()
        pic_filenames = self.get_news_pic_filenames()

        return [titles, descriptions, dates, pic_filenames]

    def close_all(self):
        if self.driver:
            try:
                self.driver.quit()
                self.logger.info("WebDriver closed successfully.")
            except Exception as e:
                self.logger.error(f"ERROR close_all() | Failed to close WebDriver: {e}")
        else:
            self.logger.error("ERROR close_all() | WebDriver not initialized.")

    def to_excel(self, data):
        wb = Workbook()
        sheet = wb.active

        headers = ['Title', 'Description', 'Date', 'Filename']
        for i_col, header in enumerate(headers):
            sheet.cell(row=1, column=i_col+1).value = header

            if data[i_col] is not None:
                for i_row, val in enumerate(data[i_col]):
                    sheet.cell(row=i_row+2, column=i_col+1).value = val
            else:
                self.logger.error(f"ERROR to_excel() | {header} list is None.")

        sheet.title = 'NEWS'
        return wb

def main():
    url = 'https://www.latimes.com/'

    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    news_scraper = NewsWebScraper()
    news_scraper.set_webdriver()
    news_scraper.open_url(url)
    news_scraper.search('Taylor Swift')
    news_scraper.sort_by_newest()
    
    news_data = news_scraper.get_news()

    wb = news_scraper.to_excel(news_data)
    wb.save('./News.xlsx')


    time.sleep(3)

    news_scraper.close_all()


if __name__ == '__main__':
    main()
