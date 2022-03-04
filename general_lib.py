import re
import subprocess
import os.path
from pathlib import Path

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    path = Path(__file__).parent / relative_path
    return str(path)

def linereturn(file,match):
        fileopen = open(file,'r')
        linearray = []
        for line in fileopen:
            if re.match(".*%s.*" % match, line):
                linearray.append(line)
        fileopen.close()
        return(linearray)

def cbilocator():
    if not os.path.isfile("cbcli_loc.txt"):
        print('locator file not present. Generating locator file')
        subprocess.call("cbcli_locator.bat", shell=True)
        print('locater file generated')
    cbloc = linereturn('cbcli_loc.txt','Production')[0]
    print("CB location: ", cbloc)
    return cbloc