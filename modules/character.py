from selenium.webdriver.firefox import webdriver
from selenium.webdriver.common.by import By

class Char():
    def __init__(self, browser: webdriver.WebDriver):
        self.browser = browser
        self.UP = 'w'
        self.DOWN = 's'
        self.LEFT = 'a'
        self.RIGHT = 'd'
        self.UP_LEFT = 'q'
        self.UP_RIGHT = 'e'
        self.DOWN_LEFT = 'z'
        self.DOWN_RIGHT = 'c'
    
    def _walk(self, direction: str) -> None:
        body = self.browser.find_element(By.TAG_NAME, 'body')
        body.send_keys(direction)