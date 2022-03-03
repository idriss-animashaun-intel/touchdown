import os
import urllib.request
import zipfile
import shutil
import time
from subprocess import Popen


touchdown_master_directory = os.getcwd()
touchdown_directory = touchdown_master_directory+"\\touchdown-updates"
touchdown_file = touchdown_directory+"\\main\\main.exe"
touchdown_rev = touchdown_directory+"\\Rev.txt"


def installation():
    print("*** Downloading new version ***")
    urllib.request.urlretrieve("https://gitlab.devtools.intel.com/ianimash/touchdown/-/archive/updates/touchdown-updates.zip", touchdown_master_directory+"\\touchdown_new.zip")
    print("*** Extracting new version ***")
    zip_ref = zipfile.ZipFile(touchdown_master_directory+"\\touchdown_new.zip", 'r')
    zip_ref.extractall(touchdown_master_directory)
    zip_ref.close()
    os.remove(touchdown_master_directory+"\\touchdown_new.zip")
    time.sleep(5)
    
def upgrade():    
    print("*** Removing old files ***")
    shutil.rmtree(touchdown_directory)
    time.sleep(10)
    installation()


### Is touchdown already installed? If yes get file size to compare for upgrade
if os.path.isfile(touchdown_file):
    local_file_size = int(os.path.getsize(touchdown_rev))
    # print(local_file_size)
    ### Check if update needed:
    f = urllib.request.urlopen("https://gitlab.devtools.intel.com/ianimash/touchdown/-/raw/updates/Rev.txt") # points to the exe file for size
    i = f.info()
    web_file_size = int(i["Content-Length"])
    # print(web_file_size)


    if local_file_size != web_file_size:# upgrade available
        updt = input("*** New upgrade available! enter <y> to upgrade now, other key to skip upgrade *** ")
        if updt == "y": # proceed to upgrade
            upgrade()
        elif updt == "Y":
            upgrade()

### touchdown wasn't installed, so we download and install it here                
else:
    install = input("Welcome to touchdown! If you enter <y> Touchdown will be downloaded in the same folder where this file is.\nAfter the installation, this same file you are running now (\"touchdown.exe\") will the one to use to open touchdown :)\nEnter any other key to skip the download\n -->")
    if install == "y":
        installation()
    elif install == "Y":
        installation()

print('Ready')


### We open the real application:
try:
    Popen(touchdown_file)
    print("*** Opening Touchdown ***")
    time.sleep(20)
except:
    print('Failed to open application, Please open manually in subfolder')
    pass