from instaloader import Profile
from win32comext.propsys import propsys, pscon
from win32comext.shell import shellcon
import instaloader
import pathlib
import time
import os

class Directory_Class:
    
    def __init__(self):
        self.User_name = '' #保存したいアカウント名
        self.Directory_Name = ['Feed', 'Story', 'Highlight', 'IGTV']

    def Make_Directory(self):
        now_directory = os.path.dirname(__file__)
        present_directory = os.path.join(now_directory, self.User_name)
        try:
            os.makedirs(present_directory)
        except FileExistsError:
            pass

        for i in self.Directory_Name:
            if i != 'Story':
                make_directory = os.path.join(present_directory, i)
                try:
                    os.makedirs(make_directory)
                except FileExistsError:
                    pass
            else:
                make_directory = os.path.join(now_directory, i)
                try:
                    os.makedirs(make_directory)
                except FileExistsError:
                    pass

        return present_directory

    def Delete_Directory(self):
        present_directory = self.Make_Directory()
        for i in self.Directory_Name:
            delete_directory = os.path.join(present_directory, i)
            print(delete_directory)
            if not os.listdir(delete_directory):
                try:
                    os.rmdir(delete_directory)
                except FileNotFoundError:
                    pass

class Instagram:
    
    def __init__(self, save_directory):
        D = Directory_Class()
        self.L = instaloader.Instaloader(save_metadata=False, download_video_thumbnails=False, post_metadata_txt_pattern='')
        self.Login_id = ['', ''] #1つ目にアカウント名、2つ目にパスワード
        self.User_name = D.User_name
        self.MediaidLists = []
        self.Save_Directory = save_directory

    def Insta(self):
        profile = Profile.from_username(self.L.context, self.User_name)

        return profile

    def Login(self):
        self.L.dirname_pattern = self.Save_Directory
        
        return self.L.login(user=self.Login_id[0], passwd=self.Login_id[1])
    
    def Get_Value(self):
        directory = os.listdir(self.Save_Directory)
        if directory:
            comparison = os.path.join(self.Save_Directory, directory[0])
            properties = propsys.SHGetPropertyStoreFromParsingName(comparison, None, shellcon.GPS_READWRITE)
            oldValue = properties.GetValue(pscon.PKEY_Title).GetValue()
        else:
            oldValue = '0'

        return oldValue

    def Mediaid_List(self):
        mediaid_list = self.MediaidLists
        directory = os.listdir(self.Save_Directory)
        if directory:
            for d in directory:
                title = os.path.join(self.Save_Directory, d)
                properties = propsys.SHGetPropertyStoreFromParsingName(title, None, shellcon.GPS_READWRITE)
                oldValue = properties.GetValue(pscon.PKEY_Title).GetValue()
                mediaid_list.append(oldValue)

        return mediaid_list

    def Feed_Post(self):
        oldValue = self.Get_Value()
        for posts_date in self.Insta().get_posts():
            date_utc = posts_date.date_utc.strftime('【%Y_%m_%d】') #フィードの日付
            caption = posts_date.caption #フィードの投稿テキスト
            mediaid = posts_date.mediaid #メディアID

            if mediaid != int(oldValue):
                file_name = self.User_name + date_utc
                self.L.filename_pattern = file_name
                self.L.dirname_pattern = self.Save_Directory
                file_list = os.listdir(self.Save_Directory)
                file_in = [s for s in file_list if file_name in s]
                self.L.download_post(posts_date, target='')
                time.sleep(0.5)

                file_list2 = os.listdir(self.Save_Directory)
                file_in2 = [s for s in file_list2 if file_name in s]
                file_sym_diff = set(file_in) ^ set(file_in2)
                i = range(len(file_in) +1, len(file_in2) + len(file_in) +1)
                b = 0
                
                for j in list(file_sym_diff):
                    if i[b] < 10:
                        file_num1 = i[b]
                        file_num10 = 0
                        num_str1 = str(file_num1)
                        num_str10 = str(file_num10)
                        num_str = num_str10 + num_str1

                    else:
                        num_str = str(i[b])
                    
                    old_rename_file = os.path.join(self.Save_Directory, j)
                    name = j.split('】')[0]
                    name = name + '】'
                    
                    file_ext = j.split('.')[-1] #拡張子
                    file_ext = '.'+file_ext

                    new_rename_file = os.path.join(self.Save_Directory, name + num_str + file_ext)

                    os.rename(old_rename_file, new_rename_file)
                    properties = propsys.SHGetPropertyStoreFromParsingName(new_rename_file, None, shellcon.GPS_READWRITE)
                    newValue = propsys.PROPVARIANTType(caption) #コメント
                    properties.SetValue(pscon.PKEY_Comment,newValue)
                    if i[b] == len(file_in)+1:
                        newValue = propsys.PROPVARIANTType(mediaid) #タイトル
                        properties.SetValue(pscon.PKEY_Title,newValue)

                    properties.Commit()
                    b += 1
            else:
                print("同じ")
                break
    
    def Story_Post(self):
        mediaid_list = self.Mediaid_List()
        for story in self.L.get_stories():
            for item in story.get_items():
                date_utc = item.date_utc.strftime('【%Y_%m_%d】') #ストーリーの日付
                owner_username = item.owner_username #アカウント名
                mediaid = item.mediaid #メディアID
                typename = item.typename
                confirmation = [s for s in mediaid_list if str(mediaid) in s]
                
                if not confirmation:
                    file_name = owner_username + date_utc
                    file_list = os.listdir(self.Save_Directory)
                    file_in = [s for s in file_list if file_name in s]

                    file_number = len(file_in) + 1
                    if file_number < 10:
                        file_num1 = file_number
                        file_num10 = 0
                        num_str1 = str(file_num1)
                        num_str10 = str(file_num10)
                        num_str = num_str10 + num_str1
                    else:
                        num_str = str(file_number)
                    
                    self.L.filename_pattern = file_name + num_str
                    self.L.download_storyitem(item, '')
                    time.sleep(0.5)

                    if typename == "GraphStoryImage":
                        file_ext = '.jpg'
                    else:
                        file_ext = '.mp4'

                    property_file = os.path.join(self.Save_Directory, file_name + num_str + file_ext)
                    properties = propsys.SHGetPropertyStoreFromParsingName(property_file, None, shellcon.GPS_READWRITE)
                    newValue = propsys.PROPVARIANTType(mediaid) #タイトル
                    properties.SetValue(pscon.PKEY_Title,newValue)
                    properties.Commit()

        print("終わり")

    def Highlight_Post(self):
        for highlight in self.L.get_highlights(self.Insta()):
            mediaid_list = self.MediaidLists
            make_directory  = os.path.join(self.Save_Directory, highlight.title)
            try:
                os.makedirs(make_directory)
            except FileExistsError:
                pass
            directory = os.listdir(make_directory)
            for d in directory:
                title_jpg = os.path.join(make_directory, d)
                properties = propsys.SHGetPropertyStoreFromParsingName(title_jpg, None, shellcon.GPS_READWRITE)
                oldValue = properties.GetValue(pscon.PKEY_Title).GetValue()
                mediaid_list.append(oldValue)
                
            for item in highlight.get_items():
                date_utc = item.date_utc.strftime('【%Y_%m_%d】') #ストーリーの日付
                owner_username = item.owner_username #アカウント名
                mediaid = item.mediaid #メディアID
                typename = item.typename
                confirmation = [s for s in mediaid_list if str(mediaid) in s]

                if not confirmation:
                    file_name = owner_username + date_utc
                    file_list = os.listdir(make_directory)
                    file_in = [s for s in file_list if file_name in s]
                    file_number = len(file_in) + 1
                    
                    if file_number < 10:
                        file_num1 = file_number
                        file_num10 = 0
                        num_str1 = str(file_num1)
                        num_str10 = str(file_num10)
                        num_str = num_str10 + num_str1
                    else:
                        num_str = str(file_number)
                    
                    self.L.dirname_pattern = make_directory
                    self.L.filename_pattern = file_name + num_str
                    self.L.download_storyitem(item, '')
                    time.sleep(0.5)

                    if typename == "GraphStoryImage":
                        file_ext = '.jpg'
                    else:
                        file_ext = '.mp4'

                    property_file = os.path.join(make_directory, file_name + num_str + file_ext)
                    properties = propsys.SHGetPropertyStoreFromParsingName(property_file, None, shellcon.GPS_READWRITE)
                    newValue = propsys.PROPVARIANTType(mediaid) #タイトル
                    properties.SetValue(pscon.PKEY_Title,newValue)
                    properties.Commit()


    def IGTV_Post(self):
        oldValue = self.Get_Value()
        for posts_date in self.Insta().get_igtv_posts():
            date_utc = posts_date.date_utc.strftime('【%Y_%m_%d】') #フィードの日付
            caption = posts_date.caption #フィードの投稿テキスト
            mediaid = posts_date.mediaid #メディアID
            typename = posts_date.typename #メディアタイプ
            
            if mediaid != int(oldValue):
                file_name = self.User_name + date_utc
                self.L.filename_pattern = file_name
                self.L.dirname_pattern = self.Save_Directory
                file_list = os.listdir(self.Save_Directory)
                file_in = [s for s in file_list if file_name in s]
                self.L.download_post(posts_date, target='')
                time.sleep(0.5)

                file_list2 = os.listdir(self.Save_Directory)
                file_in2 = [s for s in file_list2 if file_name in s]

                file_sym_diff = set(file_in) ^ set(file_in2)

                i = range(len(file_in) +1, len(file_in2) + len(file_in) +1)
                b = 0
                
                for j in list(file_sym_diff):
                    if i[b] < 10:
                        file_num1 = i[b]
                        file_num10 = 0
                        num_str1 = str(file_num1)
                        num_str10 = str(file_num10)
                        num_str = num_str10 + num_str1

                    else:
                        num_str = str(i[b])
                    
                    old_rename_file = os.path.join(self.Save_Directory, j)
                    name = j.split('】')[0]
                    name = name + '】'
                    
                    if typename == "GraphStoryImage":
                        file_ext = '.jpg'
                    else:
                        file_ext = '.mp4'

                    new_rename_file = os.path.join(self.Save_Directory, name + num_str + file_ext)

                    os.rename(old_rename_file, new_rename_file)
                    properties = propsys.SHGetPropertyStoreFromParsingName(new_rename_file, None, shellcon.GPS_READWRITE)
                    newValue = propsys.PROPVARIANTType(caption) #コメント
                    properties.SetValue(pscon.PKEY_Comment,newValue)
                    newValue = propsys.PROPVARIANTType(mediaid) #タイトル
                    properties.SetValue(pscon.PKEY_Title,newValue)
                        
                    properties.Commit()
                    b += 1

            else:
                print("同じ")
                break
