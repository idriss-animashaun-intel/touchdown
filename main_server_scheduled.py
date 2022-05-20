import os
from subprocess import call
from re import sub
from os import remove
import datetime
import win32com.client as win32
import pandas as pd
from general_lib import *
import schedule
import datetime
import time
from tkinter import *
from tkinter import Tk
from tkinter import Button
from tkinter import ttk
from tkinter import Label
from tkinter import W

location = os.getcwd()

dir = os.path.join("automated_outputs")
if not os.path.exists(dir):
    os.mkdir(dir)

def interate():

    # assign directory
    dir = r'automated_outputs'
    for f in os.listdir(dir):
        os.remove(os.path.join(dir, f))
    
    directory = r'automated_requests'
    # iterate over files in
    # that directory
    for filename in os.listdir(directory):
        f = os.path.join(directory, filename)
        # checking if it is a file
        if os.path.isfile(f):
            print(f)
            cbsql_basic(f)

def cbsql_basic(filename):
    global output_file
    global file_name
    global td_file_lines
    global file_name

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
        run_jrp()

def run_jrp():
    global output_file_csv
    global file_name

    run_wfr = 0

    user_script = "automated_outputs\\" + file_name + ".jrp"

    if "Wafer/s:  ['']" in td_file_lines[2]:
        print('plotting all wafers')
        jsl_path = resource_path("Inputs\\touchdown.jsl")
    else:
        print('plotting selected wafers')
        run_wfr = 1
        jsl_path = resource_path("Inputs\\touchdown_wfr.jsl")
        wfr_IDs_list = td_file_lines[2].strip(" ")
        print(wfr_IDs_list)
        wfr_IDs_list = td_file_lines[2].strip('    Wafer/s:  ').replace(" ", "").strip('\n').replace("'", '"')

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

    if run_wfr == 1:
        reading_file = open(user_script, "r")

        new_file_content = ""
        for line in reading_file:
            stripped_line = line.strip()
            new_line = stripped_line.replace("wfrs", f'{wfr_IDs_list}'.strip("[]"))
            new_file_content += new_line +"\n"
        reading_file.close()
        writing_file = open(user_script, "w")
        writing_file.write(new_file_content)
        writing_file.close()
    send_mail()

def send_mail():
    csv_file = output_file + '.csv'
    jrp_file = output_file + ".jrp"

    to_mail = td_file_lines[1].strip('Send Email To').strip(":")
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = to_mail
    mail.Subject = 'Touchdown Report Out ' + file_name
    mail.Body = 'Report Attached\n\n ###Note: The .csv is referenced within the .jrp.\n \tPlease ensure that the .csv is in the same file location as the .jrp before opening the .jrp\n\n Reply "STOP" to cancel automated reports to this distribution list'
    mail.Attachments.Add(csv_file)
    mail.Attachments.Add(jrp_file)
    mail.Send()

def job():
    interate()
    nowtime = str(datetime.datetime.now())
    print('Completed run @' + nowtime)

def schedule_runs():
    global run_sched
    run_sched = True
    if variable.get() == "Monday":
        schedule.every().monday.at(prod_code.get()).do(job)
    if variable.get() == "Tuesday":
        schedule.every().tuesday.at(prod_code.get()).do(job)
    if variable.get() == "Wednesday":
        schedule.every().wednesday.at(prod_code.get()).do(job)
    if variable.get() == "Thursday":
        schedule.every().thursday.at(prod_code.get()).do(job)
    if variable.get() == "Friday":
        schedule.every().friday.at(prod_code.get()).do(job)

    while run_sched == True:
        schedule.run_pending()
        time.sleep(30)

def stop_schedule():
    global run_sched
    run_sched = False
cblocation = cbilocator()

### Main Root
root = Tk()
root.title('Touchdown Server v1.00')


mainframe = ttk.Frame(root, padding="60 50 60 50")
mainframe.grid(column=0, row=0, sticky=('news'))
mainframe.columnconfigure(0, weight=3)
mainframe.rowconfigure(0, weight=3)

label_2 = Label(mainframe, text = 'Select Day: ', bg  ='black', fg = 'white')
label_2.grid(row = 0, column = 0, sticky=E)
variable = StringVar(mainframe)
variable.set("Monday") # default value

sel_day = OptionMenu(mainframe, variable, "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")

sel_day.grid(row = 0, column = 1, sticky=W)

label_0 = Label(mainframe, text = 'Enter Time: ', bg  ='black', fg = 'white')
label_0.grid(row = 1, sticky=E)
prod_code = Entry(mainframe, width=40, relief = FLAT)
prod_code.insert(4,"08:00")
prod_code.grid(row = 1, column = 1, sticky=W)

button_0 = Button(mainframe, text="Start", height = 1, width = 20, command = schedule_runs, bg = 'green', fg = 'white', font = '-family "SF Espresso Shack" -size 12')
button_0.grid(row = 2, column = 0, sticky=E)

button_1 = Button(mainframe, text="Stop", height = 1, width = 20, command = stop_schedule, bg = 'red', fg = 'white', font = '-family "SF Espresso Shack" -size 12')
button_1.grid(row = 2, column = 1, sticky=W)

### Main loop
root.mainloop()