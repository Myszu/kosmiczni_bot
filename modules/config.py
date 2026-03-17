import logging
from configparser import ConfigParser, ExtendedInterpolation

path = r'./modules/config.cfg'
cfg = ConfigParser(interpolation=ExtendedInterpolation(), allow_no_value=True)
cfg.read(path)

DEBUGGING = cfg.getboolean('Main', 'debugging')
WAIT = cfg.getint('Main', 'wait')
PROCEEDURE_WAIT = cfg.getfloat('Main', 'proceedure_wait')

FORMATTER = logging.Formatter(fmt=f'%(asctime)s | %(levelname)s - %(message)s', datefmt='%d.%m.%Y %H:%M:%S')

LOGIN = cfg.get('Credentials', 'login')
PASSWORD = cfg.get('Credentials', 'password')