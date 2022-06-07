import os
import urllib.request
import zipfile
import shutil
import time


touchdown_master_directory = os.getcwd()
touchdown_file = touchdown_master_directory+"\\Touchdown.exe"
Old_touchdown_directory = touchdown_master_directory+"\\touchdown_exe-master"

proxy_handler = urllib.request.ProxyHandler({'https': 'http://proxy-dmz.intel.com:912'})
opener = urllib.request.build_opener(proxy_handler)
urllib.request.install_opener(opener)


def installation():
    urllib.request.urlretrieve("https://github.com/idriss-animashaun-intel/touchdown/archive/refs/heads/master.zip", touchdown_master_directory+"\\touchdown_luancher_new.zip")
    print("*** Updating Launcher Please Wait ***")
    zip_ref = zipfile.ZipFile(touchdown_master_directory+"\\touchdown_luancher_new.zip", 'r')
    zip_ref.extractall(touchdown_master_directory)
    zip_ref.close()
    os.remove(touchdown_master_directory+"\\touchdown_luancher_new.zip")

    src_dir = touchdown_master_directory + "\\touchdown-master"
    dest_dir = touchdown_master_directory
    fn = os.path.join(src_dir, "Touchdown.exe")
    shutil.copy(fn, dest_dir)

    shutil.rmtree(touchdown_master_directory+"\\touchdown-master")

    time.sleep(5)
    
def upgrade():
    print("*** Updating Launcher Please Wait ***")    
    print("*** Removing old files ***")
    time.sleep(20)
    os.remove(touchdown_file)
    time.sleep(10)
    installation()


### Is touchdown already installed? If yes get file size to compare for upgrade
if os.path.isfile(touchdown_file):
    local_file_size = int(os.path.getsize(touchdown_file))
    # print(local_file_size)

    url = 'https://github.com/idriss-animashaun-intel/touchdown/raw/master/Touchdown.exe'
    f = urllib.request.urlopen(url)

    i = f.info()
    web_file_size = int(i["Content-Length"])
    # print(web_file_size)

    if local_file_size != web_file_size:# upgrade available
        upgrade()

### touchdown wasn't installed, so we download and install it here                
else:
    installation()

if os.path.isdir(Old_touchdown_directory):
        print('removing touchdown_exe-master')
        time.sleep(5)
        shutil.rmtree(Old_touchdown_directory)

print('Launcher up to date')