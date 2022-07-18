import sys
from PyQt5.QtCore import QCoreApplication
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication, QDialog, QListWidgetItem, QStackedWidget
from PyQt5.uic import loadUi
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
import xml.etree.ElementTree as ET
from pynput.keyboard import Controller
from datetime import datetime
import win32gui, win32com.client
from time import sleep
import logging

version = '1.13'
speaker = win32com.client.Dispatch("SAPI.SpVoice")
logging.basicConfig(filename='logs.log', filemode='w', format='%(levelname)s || %(asctime)s - %(message)s', datefmt='%d.%m.%y %H:%M:%S')


class Main(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        loadUi("gui/main.ui", self)
        self.get_values.clicked.connect(self.scrap)
        self.restart_exp.clicked.connect(self.exping)
        self.try_button.clicked.connect(self.tryout)
        
        servers = ['Server 1', 'Server M', 'Server 8', 'Server 9', 'Server 10', 'Server 11', 'Server 12', 'Server 13', 'Server 14', 'Server 15', 'Server 16', 'Server 17', 'Server 18', 'Server Speed']
            
        for server in servers:
            self.server_combo.addItem(server)


# Log in and get to map
    def scrap(self):
        try:
            self.error_label.setText('')
            self.start = datetime.now().timestamp()
            
            delay = 10
            browser = webdriver.Firefox()
            self.action = webdriver.ActionChains(browser)
            browser.get('https://kosmiczni.pl/')
            self.browser = browser
            
            # Accept Cookies 
            self.stoper('Cookies: ')     
            cookie = browser.find_element(By.ID, 'cookie_accept')
            cookie.click()

            # Login
            self.stoper('Login: ')      
            login = browser.find_element(By.ID, 'login_login')
            password = browser.find_element(By.ID, 'login_pass')
            submit = browser.find_element(By.ID, 'cg_login_button1')
        
            login.send_keys(self.login_input.text())
            password.send_keys(self.password_input.text())

            submit.click()
            
            # Choose server
            serv = self.server_combo.currentText()
            self.stoper('Server: ') 
            submit = browser.find_element(By.ID, 'cg_login_button2')

            try:
                combobox = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'server_choose')))
            except TimeoutException:
                logging.error('Timeout. Loading took to much time.')

            select = Select(browser.find_element(By.ID, 'server_choose'))
            select.select_by_visible_text(serv)
            
            submit.click()
            
            # Choose character
            self.stoper('Char: ') 
            try:
                chars = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'li.option')))
            except TimeoutException:
                logging.error('Timeout. Loading took to much time.')
            
            character = browser.find_element(By.CSS_SELECTOR, 'li.option')
            character.click()
            
            # Go to map
            self.stoper('Map: ') 
            play = browser.find_element(By.CSS_SELECTOR, '#map_link_btn')
            play.click()
            
            # Transform to highest SSJ
            self.stoper('Transform: ') 
            transform = browser.find_elements(By.CSS_SELECTOR, '#quick_bar > div.option.qlink')
            transformation = transform[1]
            transformation.click()
            
            # Close possible alert
            try:
                transformed = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.CLASS_NAME, 'close_kom')))
                alert = browser.find_element(By.CLASS_NAME, 'close_kom')
                alert.click()
            except TimeoutException:
                logging.error('There was no alert to close!')
                
        except:
            browser.close()
            self.error_label.setText('Something went wrong while logging in. Try again.')
            

# Exping Function
    def exping(self):
        try:
            self.error_label.setText('')
            
            mobId = self.mobnumber()
            
            moveSpeed = float(self.pause_input.text())
            attackSpeed = float(self.attack_input.text())
            
            mobRank = self.mob_box.currentIndex()
            if mobRank == 0:
                attack = 't'
            elif mobRank == 1:
                attack = 'u'
            elif mobRank == 2:
                attack = 'p'
            elif mobRank == 3:
                attack = 'g'
            
            self.findbrowser()
                
            maps = ET.parse('maps.xml')
            root = maps.getroot()
            gborn = root.find('gborn')
            
            map_exp = (str(self.path_box.currentText()) + '_exp').lower()
            map_exp = map_exp.replace(' ', '_')
            map_return = (str(self.path_box.currentText()) + '_return').lower()
            map_return = map_return.replace(' ', '_')
            
            input_exp = gborn.find(map_exp).text
            input_return = gborn.find(map_return).text

            try:
                moves = int(self.loops_input.text())
                if moves < 1:
                    moves = int(self.browser.find_element(By.ID, 'char_pa').text)
            except Exception as EX:
                logging.critical('Something went wrong while loading the path.')
                self.error_label.setText(str(EX))
                
            try:
                loopBreak = int(self.break_input.text())
                if loopBreak < 1:
                    loopBreak = 60
            except Exception as EX:
                logging.critical('Something went wrong while loading break time.')
                self.error_label.setText(str(EX))
            
            for j in range(moves):
                start_exp = self.browser.find_element(By.ID, 'char_exp').text
                start_exp = int(start_exp.replace(' ', ''))
                start_lvl = int(self.browser.find_element(By.ID, 'char_level').text)
                
                for letter in input_exp:
                    elite = self.browser.find_element(By.ID, 'mob_' + mobId +'_rank_' + str(mobRank) + '_am')
                    elites = elite.text
                    multiAttack = self.browser.find_element(By.CSS_SELECTOR, '#mob_' + mobId +'_mf > button:nth-child(2)')
                    
                    if multiAttack.is_displayed() and self.multi_checkbox.isChecked():
                        self.press_key('r')
                        sleep(attackSpeed)
                    else:
                        if elite.is_displayed():
                            for i in range(int(elites[2])):
                                self.press_key(attack)
                                sleep(attackSpeed)

                    self.legendalarm(mobId)
                    self.press_key(letter)
                    sleep(moveSpeed)
                    
                for letter in input_return:
                    self.legendalarm(mobId)
                    self.press_key(letter)
                    sleep(moveSpeed)
                        
                end_exp = self.browser.find_element(By.ID, 'char_exp').text
                end_exp = int(end_exp.replace(' ', ''))
                end_lvl = int(self.browser.find_element(By.ID, 'char_level').text)

                if (end_exp - start_exp) < 0:
                    item = (self.timestamp() + str(end_lvl - start_lvl) + ' lvl and ' + str(end_exp) + ' exp gained')
                else:
                    item = (self.timestamp() + str(end_exp - start_exp) + ' exp gained')
                    
                add = QListWidgetItem(item)
                self.log_list.addItem(add)
                QCoreApplication.processEvents()
                
                self.usesenzu()
                self.checkstart()
                
                if (j + 1) < moves:
                    sleep(loopBreak)
            
            logging.info('Session ended, all loops done.')
            speaker.Speak('Koniec')
                
        except Exception as EX:
            self.error_label.setText(str(EX))
            logging.critical('Something went wrong while executing loops.')
            speaker.Speak('Błąd')
       
       
# Tryout new function
    def tryout(self):
        self.error_label.setText('')
        
        maps = ET.parse('maps.xml')
        root = maps.getroot()
        gborn = root.find('gborn')
        
        map_exp = (str(self.path_box.currentText()) + '_exp').lower()
        map_exp = map_exp.replace(' ', '_')
        input_exp = gborn.find(map_exp).text
        moveSpeed = float(self.pause_input.text())
        attackSpeed = float(self.attack_input.text())
        
        moves = len(input_exp)
        estimation = int(((moves * attackSpeed * 3) + (moves * moveSpeed))/60)*2
        
        transformationTimer = self.browser.find_element(By.CSS_SELECTOR, '#ssj_status > span:nth-child(1)')
        if transformationTimer.is_displayed():
            minutesLeft = int(transformationTimer.text[3:5])
            if minutesLeft < estimation:
                stopCurrent = self.browser.find_element(By.CSS_SELECTOR, '#ssj_bar > span:nth-child(1)')
                stopCurrent.click()
                transform = self.browser.find_elements(By.CSS_SELECTOR, '#quick_bar > div.option.qlink')
                transformation = transform[1]
                transformation.click()
            else:
                print('more')
        
        
# Get mob ID number
    def mobnumber(self):
        mobs = self.browser.find_element(By.ID, 'creature_list_con')
        mob = mobs.find_elements(By.CSS_SELECTOR, '*')
        mobID = mob[1].get_attribute('id')
        firstUnderscore = int(mobID.find('_'))+1
        secondUnderscore = int(mobID.find('_', firstUnderscore))
        mobNumber = mobID[firstUnderscore:secondUnderscore]
        return mobNumber
        
        
# Check position after loop
    def checkstart(self):
        mapX = int(self.browser.find_element(By.ID, 'map_x').text)
        mapY = int(self.browser.find_element(By.ID, 'map_y').text)
        if mapX != 2 or mapY != 2:
            speaker.Speak('Popraw pozycje')
        

# Use senzu when energy falls below 10%
    def usesenzu(self):
        maxActionPoints = self.browser.find_element(By.ID, 'char_pa_max').text
        maxActionPoints = int(maxActionPoints.replace(' ', ''))
        actionPoints = self.browser.find_element(By.ID, 'char_pa').text
        actionPoints = int(actionPoints.replace(' ', ''))
        actionPointPercent = (actionPoints * 100) / maxActionPoints
        if actionPointPercent < 10 and self.replenish_checkbox.isChecked():
            senzuWindow = self.browser.find_element(By.CSS_SELECTOR, 'div.qlink:nth-child(5)')
            senzuWindow.click()
            sleep(0.5)
            yellowSenzu = self.browser.find_elements(By.CLASS_NAME, 'option.fast_ekw')
            self.action.move_to_element(yellowSenzu[0]).click().perform()
            sleep(0.5)
            try:
                confirm = self.browser.find_element(By.CSS_SELECTOR, '.kom > div:nth-child(2) > div:nth-child(1) > button:nth-child(9)')
                confirm.click()
                alert = self.browser.find_element(By.CLASS_NAME, 'close_kom')
                alert.click()
            except:
                item = (self.timestamp() + " couldn't find senzu")
                        
                add = QListWidgetItem(item)
                self.log_list.addItem(add)
                QCoreApplication.processEvents()
    
    
# Highlight given element with driver= element._parent element=element
    def highlightelement(self, driver, element):
        def apply_style(s):
            driver.execute_script("arguments[0].setAttribute('style', arguments[1]);", element, s)
            
        original_style = element.get_attribute('style')
        apply_style("border: 2px solid red;")
        sleep(3)
        apply_style(original_style)


# Switch window
    def findbrowser(self):
        self.toplist = []
        self.winlist = []
        
        win32gui.EnumWindows(self.enum_callback, self.toplist)
        firefox = [(hwnd, title) for hwnd, title in self.winlist if 'kosmiczni wojownicy' in title.lower()]
        firefox = firefox[0]
        
        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys('%')
            win32gui.SetForegroundWindow(firefox[0])
        except Exception as EX:
            print(EX)
        
        
# Timestamp of loop
    def timestamp(self):
        stamp = str(datetime.fromtimestamp(datetime.now().timestamp()).strftime('%d.%m %H:%M') + ' - ')
        return stamp


# Legend alarm
    def legendalarm(self, mobId):
        eliteMob = self.browser.find_element(By.ID, 'mob_' + mobId +'_rank_3')
        if eliteMob.is_displayed():
            self.press_key('g')
            sleep(0.2)
            speaker.Speak('legenda')
            
            
# Stoper Function
    def stoper(self, yellowSenzu):
        stop = datetime.now().timestamp()
        elapsed = datetime.fromtimestamp(stop - self.start)
        item = elapsed.strftime('%M:%S')
        
        add = QListWidgetItem(yellowSenzu + item)
        self.log_list.addItem(add)
        
    
# Pressing Key Function  
    def press_key(self, key):
        keyboard = Controller()
        keyboard.press(key)
        keyboard.release(key)
        
        
# Searching For Firefox Window with game open 
    def enum_callback(self, hwnd, results):
        self.winlist.append((hwnd, win32gui.GetWindowText(hwnd)))
        
        
def start_app():
    app = QApplication(sys.argv)
    win = Main()
    widget = QStackedWidget()
    widget.addWidget(win)
    widget.setWindowTitle('Auto-Kosmiczni')
    widget.setWindowIcon(QIcon('gui/img/icon.ico'))
    widget.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    start_app()