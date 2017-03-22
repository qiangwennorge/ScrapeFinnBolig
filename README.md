# ScrapeFinnBolig
Scrape finn.no bolig information

This project is trying to scrape house information from www.finn.no, including the price, area, how many bedrooms, address etc. It starts from each page to scrape the url of each house and after collecting all the urls, it goes to the urls one by one to scrape the detailed housing information, and save the collected information in excel.
urllib2 and BeautifulSoup are used.
Python 2.7
