import pathlib
import os
import os.path
from Class_Instagram import Directory_Class, Instagram
#import Local_Instagram.Class_Instagram
from win32com.propsys import propsys, pscon
from win32com.shell import shellcon

directory = Directory_Class()
now_directory = os.path.dirname(__file__)
present_name = directory.User_name
file_name = directory.Directory_Name[0]
path = os.path.join(now_directory, present_name)
path = os.path.join(path, file_name)
data = Instagram(path)

def run():
    directory.Make_Directory()
    data.Insta()
    data.Get_Value()
    data.Feed_Post()
    directory.Delete_Directory()

if __name__ == '__main__':
    app = run()