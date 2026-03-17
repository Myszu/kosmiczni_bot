import logging
from time import sleep
from selenium.webdriver.firefox import webdriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

from modules import config as cfg

# LOGGER
logger = logging.getLogger(__name__)
handler = logging.FileHandler("logs/utils.log")
formatter = cfg.FORMATTER
handler.setFormatter(formatter)
logger.addHandler(handler)

def incremental_wait(browser: webdriver.WebDriver, ec: EC = EC.presence_of_element_located, by: By = By.ID, value: str = None) -> WebElement | list[WebElement]:
    wait = WebDriverWait(browser, cfg.WAIT)
    wait_l = WebDriverWait(browser, cfg.WAIT*2)
    wait_xl = WebDriverWait(browser, cfg.WAIT*3)
    try:
        sleep(cfg.PROCEEDURE_WAIT)
        return wait.until(ec((by, value)))
    except TimeoutException:
        try:
            sleep(cfg.PROCEEDURE_WAIT)
            return wait_l.until(ec((by, value)))
        except TimeoutException:
            try:
                sleep(cfg.PROCEEDURE_WAIT)
                return wait_xl.until(ec((by, value)))
            finally:
                logger.exception(f"Failed all tries to find \"{value}\" element")