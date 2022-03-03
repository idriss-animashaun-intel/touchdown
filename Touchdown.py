import os
import urllib.request
import zipfile
import shutil
import time
from subprocess import Popen


slice_tracker_master_directory = os.getcwd()
slice_tracker_directory = slice_tracker_master_directory+"\slicetrackpuller-updates"
slice_tracker_file = slice_tracker_directory+"\\main\\main.exe"
slice_tracker_rev = slice_tracker_directory+"\\Rev.txt"


def installation():
    print("*** Downloading new version ***")
    urllib.request.urlretrieve("https://gitlab.devtools.intel.com/ianimash/slicetrackpuller/-/archive/updates/slicetrackpuller-updates.zip", slice_tracker_master_directory+"\\slicetrackpuller_new.zip")
    print("*** Extracting new version ***")
    zip_ref = zipfile.ZipFile(slice_tracker_master_directory+"\slicetrackpuller_new.zip", 'r')
    zip_ref.extractall(slice_tracker_master_directory)
    zip_ref.close()
    os.remove(slice_tracker_master_directory+"\slicetrackpuller_new.zip")
    time.sleep(5)
    
def upgrade():    
    print("*** Removing old files ***")
    shutil.rmtree(slice_tracker_directory)
    time.sleep(10)
    installation()


### Is slice_tracker already installed? If yes get file size to compare for upgrade
if os.path.isfile(slice_tracker_file):
    local_file_size = int(os.path.getsize(slice_tracker_rev))
    # print(local_file_size)
    ### Check if update needed:
    f = urllib.request.urlopen("https://gitlab.devtools.intel.com/ianimash/slicetrackpuller/-/raw/updates/Rev.txt") # points to the exe file for size
    i = f.info()
    web_file_size = int(i["Content-Length"])
    # print(web_file_size)


    if local_file_size != web_file_size:# upgrade available
        updt = input("*** New upgrade available! enter <y> to upgrade now, other key to skip upgrade *** ")
        if updt == "y": # proceed to upgrade
            upgrade()
        elif updt == "Y":
            upgrade()

### slice_tracker wasn't installed, so we download and install it here                
else:
    install = input("Welcome to slice_tracker! If you enter <y> SliceTrackPuller will be downloaded in the same folder where this file is.\nAfter the installation, this same file you are running now (\"slice_tracker.exe\") will the one to use to open slice_tracker :)\nEnter any other key to skip the download\n -->")
    if install == "y":
        installation()
    elif install == "Y":
        installation()

print('Ready')


### We open the real application:
try:
    Popen(slice_tracker_file)
    print("*** Opening SliceTrackPuller ***")
    time.sleep(20)
except:
    print('Failed to open application, Please open manually in subfolder')
    pass