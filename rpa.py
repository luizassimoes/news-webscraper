import time
import logging
from selenium import webdriver
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
        self.driver = start(browser, options=options)

    def open_url(self, url: str):
        if self.driver:
            self.driver.get(url)
        else:
            self.logger.error("ERR0R open_url(): WebDriver not initialized. You must call set_webdriver() first.")

    def close_all(self):
        if self.driver:
            self.driver.quit()
        else:
            self.logger.error("ERROR close_all(): WebDriver not initialized.")


def main():
    url = 'https://www.latimes.com/'

    logging.basicConfig(level=logging.INFO)

    news_scraper = NewsWebScraper()
    news_scraper.set_webdriver()
    news_scraper.open_url(url)

    time.sleep(4)

    news_scraper.close_all()


if __name__ == '__main__':
    main()
