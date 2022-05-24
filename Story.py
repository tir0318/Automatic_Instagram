import pathlib
import os
from Class_Instagram import Directory_Class, Instagram
#import Local_Instagram.Class_Instagram
from win32com.propsys import propsys, pscon
from win32com.shell import shellcon

directory = Directory_Class()
now_directory = os.path.dirname(__file__)
file_name = directory.Directory_Name[1]
path = os.path.join(now_directory, file_name)
data = Instagram(path)

def run():
    directory.Make_Directory()
    data.Login()
    data.Get_Value()
    data.Story_Post()

if __name__ == '__main__':
    app = run()