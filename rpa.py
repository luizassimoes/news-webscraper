import time
import logging
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
        wait = WebDriverWait(self.driver, 10)
        try:
            sort_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.select-input')))
            select = Select(sort_element)
            select.select_by_visible_text('Newest')
        except Exception as e:
            self.logger.error(f"ERROR sort_by_newest() | Error during selecting option: {e}")
        

    def close_all(self):
        if self.driver:
            try:
                self.driver.quit()
                self.logger.info("WebDriver closed successfully.")
            except Exception as e:
                self.logger.error(f"ERROR close_all() | Failed to close WebDriver: {e}")
        else:
            self.logger.error("ERROR close_all() | WebDriver not initialized.")


def main():
    url = 'https://www.latimes.com/'

    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    news_scraper = NewsWebScraper()
    news_scraper.set_webdriver()
    news_scraper.open_url(url)
    news_scraper.search('Taylor Swift')
    news_scraper.sort_by_newest()

    time.sleep(3)

    news_scraper.close_all()


if __name__ == '__main__':
    main()
