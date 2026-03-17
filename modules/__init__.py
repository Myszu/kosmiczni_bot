import os
from getpass import getpass

# CHECKING IF THE CONFIG FILE EXISTS AND HELPS TO PREPARE IT IF IT'S NOT
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

# PREPARES LOGS DIRECTORY IF DOESN'T EXIST
log_path = './logs'
if not os.path.exists(log_path):
    os.mkdir(log_path)