This is an example of a web scraper I built targeting one specific website and scraping over 2000 products. It was one piece of a much larger project consisting of ~50,000 scraped items and 5 target websites.

Purpose

      To scrape detailed information on 2k items stored on one specific website using Selenium, Apache POI, and Java.

General outline

      main (main sets up general variables, starts geckodriver, creates a new webdriver, and opens a while loop to begin the scrape)
      fetch (fetches the webpage we want to scrape from)
      parser (parses the webpage for the correct info to an arrayList)
      findHeaderColumn (scraping tool to format the data in the spreadsheet correctly)
      removeHtmlTags (remove html tags from strings if they are present)
      tagCounter (counts the html tags to help determine headers end on page)

Input

      1 spreadsheet consisting of prescraped item identifier (SKU's, EDP's, etc)

Output

      1 formatted spreadsheet including all 2k+ items, including all the detailed description the website has available on each item

Future improvements:

      automatically format the correct font, font-size, indentation
      include the pre scrape code, and create a main application which calls the prescraper, then calls the web scraper
      create a web scraping tools package for the html tag remover etc at the bottom
      possibly add thread.sleep(1000) in the web driver, so that it respects network resources of the target

More thoughts

      Because this web scraper had to be custom made for this website, its very difficult to use this code as it is. One would have to remove the whole parser, and also edit the firefox geckodriver code at the top that is for my workstation.

