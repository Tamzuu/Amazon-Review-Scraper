# Amazon Review Scraper

This Python script utilizes Selenium to scrape product reviews from Amazon. It is designed to extract valuable information such as star ratings, review titles, sizes, colors, helpful counts, user profiles, and review dates. The extracted data is then saved into an Excel spreadsheet for further analysis.

## Features

- **Dynamic Scrape:** Handles dynamic loading of review pages by automating interactions with the webpage.
- **Robust Error Handling:** Handles exceptions such as missing elements gracefully to ensure continuous scraping.
- **User Authentication:** Allows users to provide their Amazon credentials for accessing restricted content.
- **Excel Export:** Saves scraped data into an Excel file for easy storage and analysis.

## Usage

1. Clone the repository: `git clone https://github.com/Tamzuu/amazon-review-scraper.git`
2. Install dependencies: `pip install -r requirements.txt`
3. Run the script: `python main.py`
4. Follow the prompts to input your Amazon credentials and start the scraping process.

## Dependencies

- Selenium
- Openpyxl

## Note

- Ensure you have the appropriate version of ChromeDriver installed and its path specified correctly in the script.
