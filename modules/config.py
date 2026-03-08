import os
from getpass import getpass
from configparser import ConfigParser, ExtendedInterpolation

path = r'./modules/config.cfg'
if not os.path.isfile(path):
    login = input("No config file found, input your ingame login: ")
    password = getpass(prompt="Now enter your ingame password: ")
    
    try:
        with open(path, '+a', encoding='utf8') as f:
            f.write(f"""
                    [Main]
                    debugging = True
                    wait = 5
                    proceedure_wait = 1
                    
                    [Credentials]
                    login = {login}
                    password = {password}
                    """.strip())
        
        print('Config file created. You can change some settings in config.cfg file in modules folder.')
    except:
        print('Error occured while creating a config file.')
        
cfg = ConfigParser(interpolation=ExtendedInterpolation(), allow_no_value=True)
cfg.read(path)
DEBUGGING = cfg.getboolean('Main', 'debugging')
WAIT = cfg.getint('Main', 'wait')
PROCEEDURE_WAIT = cfg.getfloat('Main', 'proceedure_wait')

LOGIN = cfg.get('Credentials', 'login')
PASSWORD = cfg.get('Credentials', 'password')