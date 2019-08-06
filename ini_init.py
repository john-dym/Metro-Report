#Python 3.7
import configparser
def iniInit():
#If ini file is not found it will create one in the working directory
    config = configparser.ConfigParser(allow_no_value=True)
    config.add_section('Door = Metro No')
    config.set('Door = Metro No', '#Locations are case sensitive! I think...doors are not')
    config.set('Door = Metro No', '#Add as many doors as needed.')
    config.set('Door = Metro No', 'A7', 'MET015')
    config.set('Door = Metro No', 'A8', 'MET021')
    config.set('Door = Metro No', 'A9', 'MET031')
    with open('metro.ini', 'w') as configFile:
        config.write(configFile)