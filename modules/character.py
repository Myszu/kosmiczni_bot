import logging
from typing import Literal
from selenium.webdriver.firefox import webdriver
from selenium.webdriver.common.by import By

from modules import config as cfg

# LOGGER
logger = logging.getLogger(__name__)
handler = logging.FileHandler("logs/char.log")
formatter = cfg.FORMATTER
handler.setFormatter(formatter)
logger.addHandler(handler)

class Char():
    def __init__(self, browser: webdriver.WebDriver):
        self.browser = browser
    
    def walk(self, direction: Literal["UP", "DOWN", "LEFT", "RIGHT", "UP_LEFT", "UP_RIGHT", "DOWN_LEFT", "DOWN_RIGHT"]) -> None:
        body = self.browser.find_element(By.TAG_NAME, 'body')
        key = self.__translate_direction__(direction)
        if key:
            body.send_keys(direction)
        
    def __translate_direction__(self, direction: str) -> str:
        directions = {
            "UP": 'w',
            "DOWN": 's',
            "LEFT": 'a',
            "RIGHT": 'd',
            "UP_LEFT": 'q',
            "UP_RIGHT": 'e',
            "DOWN_LEFT": 'z',
            "DOWN_RIGHT": 'c'
        }
        return directions.get(direction)