import os, time

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import openpyxl as op

class Scraper:
    def __init__(self) -> None:
        self.email = input("Enter your email: ")
        self.password = input("Enter your password: ")
        self.url_list = ["https://www.amazon.com/WROS-Washable-Napping-Orthopedic-Present/product-reviews/B0BMK4VQJ5/ref=cm_cr_getr_d_paging_btm_next_5?ie=UTF8&reviewerType=all_reviews&pageNumber=1",
                         "https://www.amazon.com/Plufl-Original-Machine-Washable-Durable/product-reviews/B0BWSJ426S/ref=cm_cr_getr_d_paging_btm_prev_1?ie=UTF8&reviewerType=all_reviews&pageNumber=1",
                         "https://www.amazon.com/DOGKE-Waterproof-Washable-Present-Blanket/product-reviews/B0BZNMTY7S/ref=cm_cr_arp_d_paging_btm_next_2?ie=UTF8&reviewerType=all_reviews&pageNumber=1",
                         "https://www.amazon.com/Bedsure-Calming-Families-Supportive-Orthopedic/product-reviews/B0C5JH7H1J/ref=cm_cr_arp_d_paging_btm_next_2?ie=UTF8&reviewerType=all_reviews&pageNumber=1"]
        self.current_url = 1
        self.current_page = 1
        self.excel_setup()
        self.driver_setup()


    def excel_setup(self) -> None:
        self.file_name = "review_data.xlsx"
        self.file_location = os.path.expanduser("~/Desktop") + "/" + self.file_name
        self.workbook = op.Workbook()
        self.worksheet = self.workbook.active
        self.current_row = 2
        self.customize_columns()


    def customize_columns(self) -> None:
        self.worksheet["A1"].value = "URL"
        self.worksheet["B1"].value = "Star Rating"
        self.worksheet["C1"].value = "Review Title"
        self.worksheet["D1"].value = "Size"
        self.worksheet["E1"].value = "Colour"
        self.worksheet["F1"].value = "Helpful Count"
        self.worksheet["G1"].value = "Picture of Product"
        self.worksheet["H1"].value = "Username"
        self.worksheet["I1"].value = "Sub Text (location and date)"


    def driver_setup(self) -> None:
        service = Service(executable_path="chromedriver.exe")
        self.driver = webdriver.Chrome(service=service)

        self.open_url(self.url_list[self.current_url - 1])


    def open_url(self, url: str) -> None:
        self.driver.get(url)

        self.scrape_data()


    def insert_into_excel(self, url: str, star_rating: int, review_title: str, size: str, colour: str, helpful_count: str, picture: str, username: str, sub_text: str) -> None:
        self.worksheet["A" + str(self.current_row)].value = url
        self.worksheet["B" + str(self.current_row)].value = star_rating
        self.worksheet["C" + str(self.current_row)].value = review_title
        self.worksheet["D" + str(self.current_row)].value = size
        self.worksheet["E" + str(self.current_row)].value = colour
        self.worksheet["F" + str(self.current_row)].value = helpful_count
        self.worksheet["G" + str(self.current_row)].value = picture
        self.worksheet["H" + str(self.current_row)].value = username
        self.worksheet["I" + str(self.current_row)].value = sub_text

        self.current_row += 1


    def scrape_data(self) -> None:
        try:
            table = self.driver.find_element(By.ID, "cm_cr-review_list")
            reviews = table.find_elements(By.CLASS_NAME, "a-section.celwidget")
        except NoSuchElementException:
            self.log_in()
            time.sleep(1)

            table = self.driver.find_element(By.ID, "cm_cr-review_list")
            reviews = table.find_elements(By.CLASS_NAME, "a-section.celwidget")

        for index, review in enumerate(reviews):
            print("Scraping page", self.current_page, "review", index + 1)
        
            review_title_html = review.find_element(By.CLASS_NAME, "a-size-base.a-link-normal.review-title.a-color-base.review-title-content.a-text-bold")
            url = review_title_html.get_attribute("href")
            star_rating = review_title_html.find_element(By.TAG_NAME, "i").get_dom_attribute("class").replace("-", " ")
            star_rating = [int(s) for s in star_rating.split() if s.isdigit()][0]
            review_title = review_title_html.find_elements(By.TAG_NAME, "span")[-1].text
            try:
                size = review.find_element(By.CLASS_NAME, "a-size-mini.a-link-normal.a-color-secondary").text.split("Color: ")[0].replace("Size: ", "")
                colour = review.find_element(By.CLASS_NAME, "a-size-mini.a-link-normal.a-color-secondary").text.split("Color: ")[1]
            except NoSuchElementException:
                size = None
                colour = None
            try:
                helpful_count = review.find_element(By.CLASS_NAME, "a-size-base.a-color-tertiary.cr-vote-text").text
                if "person" in helpful_count:
                    helpful_count = 1
                else:
                    helpful_count = int(helpful_count.split(" people")[0])
            except NoSuchElementException:
                helpful_count = 0

            try:
                picture = review.find_element(By.CLASS_NAME, "review-image-tile").get_dom_attribute("src")
            except NoSuchElementException:
                picture = None

            username = review.find_element(By.CLASS_NAME, "a-profile-name").text
            sub_text = review.find_element(By.CLASS_NAME, "a-size-base.a-color-secondary.review-date").text

            if "the" in sub_text:
                sub_text = sub_text.split("the ")[1]
            else:
                sub_text = sub_text.split("in ")[1]
            sub_text = sub_text.replace(" on", ",") 

            print(url, star_rating, review_title, size, colour, helpful_count, picture, username, sub_text)

            self.insert_into_excel(url, star_rating, review_title, size, colour, helpful_count, picture, username, sub_text)

        self.next_page()


    def next_page(self) -> None:
        try:
            next_page_button = self.driver.find_element(By.CLASS_NAME, "a-last").find_element(By.TAG_NAME, "a")
            self.next_page_url = self.driver.find_element(By.CLASS_NAME, "a-last").find_element(By.TAG_NAME, "a").get_attribute("href")
            self.driver.execute_script("arguments[0].scrollIntoView();", next_page_button)
            time.sleep(1)
            self.driver.execute_script("arguments[0].click();", next_page_button)

            self.current_page += 1

            time.sleep(1)
            self.save_data()
            self.scrape_data()
        except NoSuchElementException:
            self.save_data()
            if self.current_url != len(self.url_list):
                self.current_url += 1
                self.open_url(self.url_list[self.current_url - 1])
            else:
                print("DONE")


    def save_data(self) -> None:
        self.workbook.save(self.file_location)


    def log_in(self) -> None:
        email_input = self.driver.find_element(By.ID, "ap_email")
        email_input.send_keys(self.email)
        self.driver.find_element(By.ID, "continue").click()
        password_input = self.driver.find_element(By.ID, "ap_password")
        password_input.send_keys(self.password)
        self.driver.find_element(By.ID, "signInSubmit").click()


Scraper()