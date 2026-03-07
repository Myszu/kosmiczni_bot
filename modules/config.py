from configparser import ConfigParser, ExtendedInterpolation

cfg = ConfigParser(interpolation=ExtendedInterpolation(), allow_no_value=True)
cfg.read(r'./modules/config.cfg')

DEBUGGING = cfg.getboolean('Main', 'debugging')
WAIT = cfg.getint('Main', 'wait')
PROCEEDURE_WAIT = cfg.getfloat('Main', 'proceedure_wait')

LOGIN = cfg.get('Credentials', 'login')
PASSWORD = cfg.get('Credentials', 'password')