# Los Angeles Times News Scraper

This Python project is a web scraping tool designed to extract news articles from the Los Angeles Times website based on user input. The script allows users to search for a specific topic, filter the news by the latest N months, and collect key information from the results. The tool uses Selenium via the rpaframework for browser automation.

### Key Features:

- User Input-Driven Search: Users can enter a search phrase, a news section, and the number of months to filter.
- Web Scraping: Automates interaction with the Los Angeles Times website, collecting news articles and saving relevant data.
- Data Extraction: Extracts the following details for each news article:
   - Title
   - Publication date
   - Description
   - Associated image filename

- Excel File Output: Saves the collected data into an Excel file, with:
   - A column for the count of search phrases in the title and description.
   - A column indicating whether any mention of money (in various formats) appears in the title or description.

- No API or Web Requests: The script operates solely through browser automation without relying on APIs or external web requests.
