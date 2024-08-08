from selenium.webdriver.common.keys import Keys
from Cselenium import ExtendedSelenium
from robocorp.tasks import task
from openpyxl import Workbook
import re


@task
def minimal_task():
    """main function of the project"""
    web_driver = ExtendedSelenium()

    initialize_web_driver(web_driver)
    insert_search_info(web_driver)
    
    titles, descriptions, contains_money, image_paths = extract_news_data(web_driver)
    save_to_excel(titles, descriptions, contains_money, image_paths)


def initialize_web_driver(web_driver: ExtendedSelenium):
    """Initialize the web driver by opening the site and setting implicit wait."""
    web_driver.open_site("https://gothamist.com/")
    search_icon = web_driver.find_element("//div[@class='search-button']")
    search_icon.click()
    web_driver.set_browser_implicit_wait(10)


def insert_search_info(web_driver: ExtendedSelenium):
    """Insert search information into the search input field."""
    web_driver.reload_page()
    search_input = web_driver.find_element("//input[@class='search-page-input']")
    search_input.send_keys("politics")
    search_input.send_keys(Keys.ENTER)
    web_driver.set_browser_implicit_wait(10)


def extract_news_data(web_driver: ExtendedSelenium):
    """Extract news data and return lists of titles, descriptions, money status, and image paths."""
    news_items = get_news_items(web_driver)
   
    titles = []
    descriptions = []
    contains_money = []
    image_paths = []
    money_pattern = re.compile(
        r'\$\d{1,3}(?:,\d{3})*(?:\.\d{2})?|\d+(?:,\d{3})*(?:\.\d{2})?\s?(?:dollars|USD)',
        re.IGNORECASE,
    )
 
    for index, news in enumerate(news_items):
        title = news[1]
        description = news[2]
        has_money = bool(money_pattern.search(title)) or bool(money_pattern.search(description))
        image_path = capture_news_image(web_driver, index + 1, index)

        titles.append(title)
        descriptions.append(description)
        contains_money.append(has_money)
        image_paths.append(image_path)

    return titles, descriptions, contains_money, image_paths


def get_news_items(web_driver: ExtendedSelenium):
    """Retrieve and split news items from the web page."""
    news_elements = web_driver.find_elements(
        "//div[@class='v-card gothamist-card mod-horizontal mb-3 lg:mb-5 tag-small']"
    )
    
    news_list = [news.text.split('\n') for news in news_elements]
    return news_list


def capture_news_image(web_driver: ExtendedSelenium, figure_index: int, order: int):
    """Capture and save a screenshot of a news image."""
    screenshot_path = f"news{order}.png"
    web_driver.capture_element_screenshot(
        f"//*[@id='resultList']/div[2]/div[{figure_index}]/div/div[1]/figure[2]/div/div/a/div/img",
        screenshot_path
    )
    return screenshot_path


def save_to_excel(titles, descriptions, contains_money, image_paths):
    """Save extracted news data to an Excel file."""
    workbook = Workbook()
    worksheet = workbook.active

    worksheet.append(['Title', 'Description', 'Contains Money', 'News Image Path'])

    for title, description, has_money, image_path in zip(titles, descriptions, contains_money, image_paths):
        worksheet.append([title, description, has_money, image_path])

    workbook.save('result.xlsx')


if __name__ == '__main__':
    minimal_task()
