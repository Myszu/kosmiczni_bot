from selenium.webdriver.firefox import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement

class UserInterface():
    def __init__(self, browser) -> None:
        self.browser: webdriver.WebDriver = browser
        self.update_interface()
    
    def update_interface(self) -> None:
        self.prepare_top_menu()
        self.prepare_quick_bar()
        
    def prepare_top_menu(self) -> None:
        top_bar = self.browser.find_element(By.ID, 'top_menu')
        self.map: WebElement = top_bar.find_element(By.ID, 'map_link_btn')
        sections = top_bar.find_elements(By.TAG_NAME, 'button')
    
    def prepare_quick_bar(self) -> None:
        quick_bar = self.browser.find_element(By.ID, 'quick_bar')
        options = quick_bar.find_elements(By.TAG_NAME, 'ul')
        for option in options:
            if option.get_attribute('data-option') == 'use_teleport':
                self.teleport: WebElement = option
            elif option.get_attribute('data-option') == 'use_transform':
                self.transform: WebElement = option