import os
from tkinter import *
from tkinter import Tk
from tkinter import Button
from tkinter import ttk
from tkinter import Label
from tkinter import W
import webbrowser
from subprocess import call
from re import sub
from os import remove
import datetime
import win32com.client as win32
import pandas as pd
import json
from general_lib import *

location = os.getcwd()

dir = os.path.join("automated_outputs")
if not os.path.exists(dir):
    os.mkdir(dir)

def cbsql_basic():
    global output_file
    global file_name
    global td_file_lines

    filename = r'automated_requests\test.txt'

    with open(filename, 'r') as fh:
        td_file_lines = [str(line) for line in fh]
    
    today = datetime.datetime.now()
    date_time = today.strftime("%d-%m-%Y_%H-%M-%S")

    file_name = td_file_lines[0].strip('Script Name: ')

    file_name = sub(r"\s", "_", str(file_name)) + date_time
    output_file = str(location) + '\\automated_outputs\\' + file_name
    scriptpath =  str(location) + '\\' + sub(r"\s", "_", str(file_name)) + "_cb_script.acs"
    print("Path for script: %s" % scriptpath)
    print("Output file: %s" % output_file)

    td_file_lines[7] = "                OUTPUT='" + str(output_file) + "'\n"
    fo = open(scriptpath,'w')
    for i in range(3,len(td_file_lines)):
        fo.write(td_file_lines[i])
    
    fo.close()

    cbicall = sub(r"\n", "", str(cblocation) + " tool=runscript script=" + '"' + str(scriptpath) + '"')
    print("CBI call: %s" % cbicall)
    call(cbicall)
    remove(scriptpath)
    reformat()
    
def reformat():
    global output_file
    global output_file_csv
    found_data = 0
    try:
        file = pd.read_csv(output_file, delim_whitespace=True)
        file.to_csv(output_file + '.csv', encoding='utf-8', index=False)
        remove(output_file)
        output_file_csv = output_file + '.csv'
        found_data = 1
    except:
        print('Query Returned Empty, No Data Found')
    
    if found_data == 1:
        try:
            run_jrp()
        except:
            print('your version of JMP is not supported. Analysis completed check automated_outputs folder for data')

def run_jrp():
    global output_file_csv

    run_wfr = 0

    jsl_path = resource_path("Inputs\\touchdown.jsl")

    user_script = "automated_outputs\\" + file_name + ".jrp"
    # if td_file_lines[2] == "    Wafer/s:  [''] ":
    #     jsl_path = resource_path("Inputs\\touchdown.jsl")
    # else:
    #     run_wfr = 1
    #     jsl_path = resource_path("Inputs\\touchdown_wfr.jsl")
    #     wfr_IDs_list = td_file_lines[2].strip("    Wafer/s:  ['")
    #     print('wfr_IDs: ', wfr_IDs_list)


    reading_file = open(jsl_path, "r")

    new_file_content = ""
    for line in reading_file:
        stripped_line = line.strip()
        new_line = stripped_line.replace("C:\Scripts", file_name + '.csv')
        new_file_content += new_line +"\n"
    reading_file.close()
    writing_file = open(user_script, "w")
    writing_file.write(new_file_content)
    writing_file.close()

    # if run_wfr == 1:
    #     reading_file = open(user_script, "r")

    #     new_file_content = ""
    #     for line in reading_file:
    #         stripped_line = line.strip()
    #         new_line = stripped_line.replace("wfrs", f'{wfr_IDs_list}'.strip("[]"))
    #         new_file_content += new_line +"\n"
    #     reading_file.close()
    #     writing_file = open(user_script, "w")
    #     writing_file.write(new_file_content)
    #     writing_file.close()


cblocation = cbilocator()

cbsql_basic()