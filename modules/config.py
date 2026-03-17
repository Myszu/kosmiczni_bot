from configparser import ConfigParser, ExtendedInterpolation

path = r'./modules/config.cfg'
cfg = ConfigParser(interpolation=ExtendedInterpolation(), allow_no_value=True)
cfg.read(path)
DEBUGGING = cfg.getboolean('Main', 'debugging')
WAIT = cfg.getint('Main', 'wait')
PROCEEDURE_WAIT = cfg.getfloat('Main', 'proceedure_wait')

LOGIN = cfg.get('Credentials', 'login')
PASSWORD = cfg.get('Credentials', 'password')