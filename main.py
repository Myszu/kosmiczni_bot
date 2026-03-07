import logging

from time import sleep
from selenium.webdriver.firefox import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

from modules import config as cfg
from modules.interface import UserInterface

# LOGGING FORMAT
logging.basicConfig(format=f'%(asctime)s | %(levelname)s - %(message)s', datefmt='%d.%m.%Y %H:%M:%S', level=logging.INFO, filename=r'./logs/main.log', force=True)

class Bot():
    def __init__(self) -> None:
        self.browser = webdriver.WebDriver()
        self.browser.implicitly_wait(1)
        self.browser.get('https://kosmiczni.pl/')
        self.wait = WebDriverWait(self.browser, cfg.WAIT)

    def login(self) -> None:
        login = self.wait.until(EC.presence_of_element_located((By.ID, 'login_login')))
        login.send_keys(cfg.LOGIN)

        password = self.browser.find_element(By.ID, 'login_pass')
        password.send_keys(cfg.PASSWORD)
        
        submit = self.browser.find_element(By.ID, 'cg_login_button1')
        submit.click()
    
    def choose_server(self) -> None:
        server = self.wait.until(EC.presence_of_element_located((By.ID, 'server_choose')))
        server.click()
        servers = server.find_elements(By.TAG_NAME, 'option')
        servers[0].click()
        
        submit = self.browser.find_element(By.ID, 'cg_login_button2')
        submit.click()
        
    def choose_char(self) -> None:
        chars_list = self.wait.until(EC.presence_of_element_located((By.ID, 'char_list_con')))
        chars = chars_list.find_elements(By.TAG_NAME, 'li')
        chars[0].click()
        
    def login_successful(self) -> bool:
        try:
            stats = self.wait.until(EC.presence_of_element_located((By.ID, 'main_char_stats')))
            if stats:
                self.browser.execute_script("war_container.style.display = 'none'")
                self.ui = UserInterface(self.browser)
                return True
        except:
            return False
        
    def is_ssj(self) -> bool:
        try:
            ssj = self.wait.until(EC.presence_of_element_located((By.ID, 'ssj_status')))
            if ssj:
                return True
        except:
            return False
        
    def play_loop(self) -> None:
        if not self.is_ssj():
            pass
        
if __name__ == "__main__":
    try:
        bot = Bot()
        bot.login()
        sleep(cfg.PROCEEDURE_WAIT)
        bot.choose_server()
        sleep(cfg.PROCEEDURE_WAIT)
        bot.choose_char()
        sleep(cfg.PROCEEDURE_WAIT)
        if bot.login_successful():
            if bot.is_ssj():
                sleep(5)
            else:
                bot.ui.transform.click()
        sleep(5)
    except:
        logging.exception('Error in main procedure.')
    finally:
        bot.browser.quit()