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

dir = os.path.join("Output_touchdown")
if not os.path.exists(dir):
    os.mkdir(dir)

def cbsql_basic():
    global output_file
    global file_name

    store_val()
    
    today = datetime.datetime.now()
    date_time = today.strftime("%d-%m-%Y_%H-%M-%S")

    site = variable.get();
    print('site: ', site)

    wafer_list = wfr.get().replace(" ", "").split(",")
    print('Wafer/s: ', wafer_list)

    product_list = prod_code.get().replace(" ", "").split(",")
    print('Product/s: ', product_list)

    multi = 0
    if len(product_list) == 1:
        name = product_list[0]
    elif len(product_list) > 1:
        multi = 1
        name = product_list[0]
        for i in range(1,len(product_list)):
            name = name + '-' + product_list[i]

    file_name = sub(r"\s", "_", str(name)) + '_' + date_time
    output_file = str(location) + '\\Output_touchdown\\' + file_name
    scriptpath =  str(location) + '\\' + sub(r"\s", "_", str(name)) + "_cb_script.acs"
    print("Path for script: %s" % scriptpath)
    print("Output file: %s" % output_file)

    fo = open(scriptpath,'w')
    fo.write("<analysis app=cb  >\n")
    fo.write("TOOL=RUNSQL\n")
    fo.write("SCHEMA=ARIES\n")
    fo.write("SITE="+ str(site)+"\n")
    fo.write("OUTPUT=" + '"' + str(output_file) + '"' +"\n")
    fo.write("/ASMERLIN\n")
    fo.write("<SQL >\n")
    fo.write("select ats.devrevstep, ats.WAFER_ID, SUM(ats.WAFER_ID/ats.WAFER_ID) as Touchdowns\n")
    fo.write("from a_testing_session ats\n")
    fo.write("where ats.LOT like '3%'\n")
    
    if multi == 0:
        if '%' in product_list[0]:
            fo.write("and ats.devrevstep like " + "'" + product_list[0] + "'" + "\n")
        else:
            fo.write("and ats.devrevstep = " + "'" + product_list[0] + "'" + "\n")
    else:
        if '%' in product_list[0]:
            fo.write("and ats.devrevstep like " + "'" + product_list[0] + "'" + "\n")
        else:
            fo.write("and ats.devrevstep = " + "'" + product_list[0] + "'" + "\n")
        for j in range(1,len(product_list)):
            if '%' in product_list[0]:
                fo.write("or ats.devrevstep like " + "'" + product_list[j] + "'" + "\n")
            else:
                fo.write("or ats.devrevstep = " + "'" + product_list[j] + "'" + "\n")

    fo.write("and ats.latest_flag = 'Y'\n")
    fo.write("group by ats.WAFER_ID, ats.devrevstep\n")
    fo.write("order by ats.devrevstep, ats.WAFER_ID\n")
    fo.write("</SQL >\n")
    fo.write("</analysis>\n")
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
            print('your version of JMP is not supported. Analysis completed check Output_touchdown folder for data')

def automate():
    store_val()

    print('Summary')

    site = variable.get();
    print('site: ', site)

    remove_zero()
    print('Wafer/s: ', split_wfr_list)

    product_list = prod_code.get().replace(" ", "").split(",")
    print('Product/s: ', product_list)

    report_email = input_email.get()

    multi = 0
    if len(product_list) == 1:
        name = product_list[0]
    elif len(product_list) > 1:
        multi = 1
        name = product_list[0]
        for i in range(1,len(product_list)):
            name = name + '-' + product_list[i]

    script_name =  sub(r"\s", "_", str(name)) + "_weekly_report"

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'idriss.animashaun@intel.com'
    mail.Subject = 'Please add report to automated scripts'
    mail.Body = f'''    Script Name: {script_name}
    Send Email To: {report_email}\n
    Wafer/s:  {split_wfr_list}\n
    <analysis app=cb  >
    TOOL=RUNSQL
    SCHEMA=ARIES
    SITE="{str(site)}
    OUTPUT=
    /ASMERLIN
    <SQL >
    select ats.devrevstep, ats.WAFER_ID, SUM(ats.WAFER_ID/ats.WAFER_ID) as Touchdowns
    from a_testing_session ats
    where ats.LOT like '3%' '''
    if multi == 0:
        mail.Body = mail.Body + "   and ats.devrevstep = " + "'" + product_list[0] + "'" + "\n"
    else:
        mail.Body = mail.Body + "   and ats.devrevstep = " + "'" + product_list[0] + "'" + "\n"
        for j in range(1,len(product_list)):
            mail.Body = mail.Body + "   or ats.devrevstep = " + "'" + product_list[j] + "'" + "\n"
    mail.Body = mail.Body + '''     and ats.latest_flag = 'Y'
    group by ats.WAFER_ID, ats.devrevstep
    order by ats.devrevstep, ats.WAFER_ID
    </SQL >
    </analysis>'''

    mail.Send()

    print('Request Submitted')

def run_jrp():
    global output_file_csv

    user_script = "Output_touchdown\\" + file_name + ".jrp"
    jsl_path = resource_path("Inputs\\touchdown_wfr.jsl")
    reading_file = open(jsl_path, "r")

    new_file_content = ""
    for line in reading_file:
        stripped_line = line.strip()
        new_line = stripped_line
        new_file_content += new_line +"\n"
    reading_file.close()
    writing_file = open(user_script, "w")
    writing_file.write(new_file_content)
    writing_file.close()


    
    os.system(user_script)
    print('Analysis completed report sent to')

cblocation = cbilocator()

hist_path = resource_path("Inputs\\INFO.tmp");
read_hist_file = open(hist_path, "r")

prod = read_hist_file.readline().strip('\n')
wafr = read_hist_file.readline().strip('\n')
e_mail = read_hist_file.readline().strip('\n')