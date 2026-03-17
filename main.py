import logging, os

from time import sleep
from selenium.webdriver import Firefox
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver.common.by import By

from modules import config as cfg
from modules.interface import UserInterface
from modules.character import Char
from modules.utils import incremental_wait

# LOGGING FORMAT
logging.basicConfig(format=f'%(asctime)s | %(levelname)s - %(message)s', datefmt='%d.%m.%Y %H:%M:%S', level=logging.INFO, filename='./logs/main.log', force=True)

class Bot():
    def __init__(self) -> None:
        if cfg.DEBUGGING:
            logging.info(f"Initializing bot")
        options = Options()
        options.set_preference("dom.webnotifications.enabled", False)
        options.set_preference("webdriver_click", False)
        self.browser = Firefox(options=options)
        self.browser.implicitly_wait(1)
        self._incremental_wait = incremental_wait
        
    def _debugger(func):
        def myinner(self):
            name = func.__name__
            if cfg.DEBUGGING:
                logging.info(f"Started {name} procedure")
            
            result = func(self) # Wrapped function
            
            if cfg.DEBUGGING:
                logging.info(f"Ended {name} procedure")
            sleep(cfg.PROCEEDURE_WAIT)
            
            if result:
                return result # Returns if the wrapped function is returning
        return myinner

    
    
    def _try_click(self, element: WebElement, aggresive: bool = False) -> None:
        if not aggresive:
            try:
                element.click()
                return
            except ElementClickInterceptedException:
                try:
                    pop_up = self._incremental_wait(self.browser, EC.presence_of_element_located, By.ID, 'close_kom')
                    pop_up.click()
                    
                    sleep(cfg.PROCEEDURE_WAIT)
                    element.click()
                    return
                except ElementClickInterceptedException:
                    self._try_click(element, True)
                    return
        try:
            rect = element.rect
            x = rect['x'] + rect['width'] / 2
            y = rect['y'] + rect['height'] / 2
            self.browser.execute_script("""let el = document.elementFromPoint(arguments[0],arguments[1]); if(el) el.remove();""",x,y)
            logging.warning('Aggresive removal method was used to release obstructed view.')
        finally:
            return
        
    @_debugger
    def login(self) -> None:
        self.browser.get('https://kosmiczni.pl/')
        
        login = self._incremental_wait(self.browser, EC.presence_of_element_located, By.ID, 'login_login')
        login.send_keys(cfg.LOGIN)

        password = self.browser.find_element(By.ID, 'login_pass')
        password.send_keys(cfg.PASSWORD)
        
        submit = self.browser.find_element(By.ID, 'cg_login_button1')
        submit.click()
    
    @_debugger
    def choose_server(self) -> None:
        server_list = self._incremental_wait(self.browser, EC.element_to_be_clickable, By.ID, 'server_choose')
        self._try_click(server_list)
        
        servers = server_list.find_elements(By.TAG_NAME, 'option')
        servers[0].click()
        
        submit = self.browser.find_element(By.ID, 'cg_login_button2')
        submit.click()
    
    @_debugger
    def choose_char(self) -> None:
        chars_list = self._incremental_wait(self.browser, EC.presence_of_element_located, By.ID, 'char_list_con')
        chars = self._incremental_wait(self.browser, EC.presence_of_all_elements_located, By.CSS_SELECTOR, '#char_list_con > li.option')
        chars = chars_list.find_elements(By.TAG_NAME, 'li')
        chars[0].click()
    
    @_debugger
    def login_successful(self) -> bool:
        try:
            stats = self._incremental_wait(self.browser, EC.presence_of_element_located ,By.ID, 'main_char_stats')
            if stats:
                self.browser.execute_script("war_container.style.display = 'none'")
                return True
        except:
            return False
    
    @_debugger
    def is_ssj(self) -> bool:
        try:
            ssj = self._incremental_wait(self.browser, EC.presence_of_element_located, By.ID, 'ssj_status')
            if ssj:
                return True
        except:
            return False
    
    @_debugger
    def play_loop(self) -> None:
        self.ui = UserInterface(self.browser)
        self.character = Char(self.browser)
        self.ui.prepare_quick_bar()
        if not self.is_ssj():
            self.ui.transform.click()
        sleep(cfg.PROCEEDURE_WAIT)
        self.ui.buttons.navigate("Map")
        self.character.walk("DOWN")
        sleep(cfg.PROCEEDURE_WAIT)

if __name__ == "__main__":
    try:
        bot = Bot()
        bot.login()
        bot.choose_server()
        bot.choose_char()
        if logged:= bot.login_successful():
            bot.play_loop()
        sleep(5)
    except:
        logging.exception('Error in main procedure.')
    finally:
        logging.info('Quitting bot.')
        bot.browser.quit()