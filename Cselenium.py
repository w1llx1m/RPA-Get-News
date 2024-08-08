from RPA.Browser.Selenium import Selenium
from webdriver_manager.chrome import ChromeDriverManager
from SeleniumLibrary.base import keyword

class ExtendedSelenium(Selenium):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        cdm = ChromeDriverManager()
        cdm.install
    
    @keyword
    def looking_at_element(self, locator):
        element = self.get_webelement(locator)
        self.logger.warn(dir(element))

    @keyword
    def open_site(self, url, **kwargs):
        desired_capabilities = {
            "goog:loggingPrefs" : { 'browser':'ALL',  'driver': 'ALL', 'performance': 'ALL' }
        }
        self.open_browser(
            url=url,
            desired_capabilities=desired_capabilities,
            **kwargs
        )

    @keyword
    def print_webdriver_log(self, logtype):
        print(f"\n{logtype.capitalize()} Log")
        return self.driver.get_log(logtype)