# news-webscraper

The code from tasks.py aims to access the Los Angeles Times website, search (user entry), and get all the news from the last N (user entry) months from a specific topic (user entry). 

### The main steps:

1. Open the site by following the link: latimes.com
2. Enter a phrase in the search field
3. On the result page, select a news section and order the news by Newest
4. Get the values: title, description, date, and picture
5. Store in an Excel file: title, date, description, and picture filename
6. Add a column for the count of search phrases in the title and description
7. Add a column for True or False, depending on whether the title or description contains any amount of money<br>
   Possible formats: $11.1 | $111,111.11 | 11 dollars | 11 USD

The idea is to get news until it reaches the period before the desired period.  

### Rules:
1. Do NOT use APIs or Web Requests
2. Please leverage pure Python and pure Selenium (via rpaframework) without utilizing Robot Framework
