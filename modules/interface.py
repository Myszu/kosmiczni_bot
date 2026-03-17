from typing import Literal
from pydantic import BaseModel
from selenium.webdriver.firefox import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement

class UserInterface():
    def __init__(self, browser: webdriver.WebDriver) -> None:
        self.browser = browser
        self.update_interface()
    
    def update_interface(self) -> None:
        self.prepare_top_menu()
        self.prepare_quick_bar()
        
    def prepare_top_menu(self) -> None:
        top_bar = self.browser.find_element(By.ID, 'top_menu')
        sections = top_bar.find_elements(By.TAG_NAME, 'button')
        for section in sections:
            if section.get_attribute('data-page') == 'game_klan':
                clan: WebElement = section
            elif section.get_attribute('data-page') == 'game_raps':
                reports: WebElement = section
            elif section.get_attribute('data-page') == 'game_pw':
                messages: WebElement = section
        if not all([top_bar, clan, reports, messages]):
            return
        self.buttons = Buttons(
            world_map=top_bar.find_element(By.ID, 'map_link_btn'),
            clan=clan,
            reports=reports,
            messages=messages
        )
    
    def prepare_quick_bar(self) -> None:
        quick_bar = self.browser.find_element(By.ID, 'quick_bar')
        options = quick_bar.find_elements(By.TAG_NAME, 'div')
        for option in options:
            if option.get_attribute('data-option') == 'use_teleport':
                self.teleport: WebElement = option
            elif option.get_attribute('data-option') == 'use_transform':
                self.transform: WebElement = option
            elif option.get_attribute('data-option') == 'daily_reward':
                self.daily: WebElement = option
            elif option.get_attribute('data-option') == 'game_buffs':
                self.blessings: WebElement = option
            elif option.get_attribute('data-option') == 'quick_use_subs':
                self.potions: WebElement = option
            elif option.get_attribute('data-option') == 'game_empire':
                self.empire: WebElement = option
    
    def goto(self, element: Literal["Teleport", "Daily", "Blessings", "Potions", "Empire"]):
        if element == "Teleport":
            self.teleport.click()
        elif element == "Daily":
            self.daily.click()
        elif element == "Blessings":
            self.blessings.click()
        elif element == 'Potions':
            self.potions.click()
        elif element == 'Empire':
            self.empire.click()
        else:
            l
    
class Buttons(BaseModel):
    world_map: WebElement | None
    clan: WebElement | None
    reports: WebElement | None
    messages: WebElement | None