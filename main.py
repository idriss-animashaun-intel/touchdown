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


    if multi == 0:
        devrevstep = "and ats.devrevstep like '" + product_list[0] + "'"
    else:
        devrevstep = "and (ats.devrevstep like '" + product_list[0] + "'"
        for j in range(1,len(product_list)):
            devrevstep+= "or ats.devrevstep like '" + product_list[j] + "'"
        devrevstep+= ")"

    
    fo = open(scriptpath,'w')

    fo.write("""<analysis app=cb
                TOOL=RUNSQL
                SCHEMA=ARIES
                SITE="""+ str(site) + """
                OUTPUT='""" + str(output_file) + """'
                /ASMERLIN
                <SQL >
                select alltable.devrevstep, alltable.WAFER_ID, alltable.Touchdowns, alltable.SORT_LOT
                from
                    (
                    select table.devrevstep, table.WAFER_ID, table.Touchdowns, table.SORT_LOT
                    from (
                        select  ats.devrevstep, ats.WAFER_ID, afs.Touchdowns, CASE WHEN (ISNUMERIC(SUBSTRING(ats.LOT, 1, 1))!=1) THEN ats.LOT ELSE 'LOT_NOT_FOUND' END as SORT_LOT, ats.WAFER_SCRIBE
                        from (
                            select ats.WAFER_SCRIBE, ats.WAFER_ID, ats.devrevstep, idriss.Touchdowns
                            from a_testing_session   ats,
                                (
                                    select ats.WAFER_ID, ats.devrevstep, COUNT(ats.WAFER_ID) as Touchdowns
                                    from a_testing_session   ats
                                    where ISNUMERIC(SUBSTRING(ats.LOT, 1, 1)) = 1 
                                    """ + devrevstep + """
                                    and   ats.latest_flag = 'Y'
                                    group by ats.WAFER_ID, ats.devrevstep
                                    order by ats.devrevstep, ats.WAFER_ID
                                ) idriss

                            where
                            ISNUMERIC(SUBSTRING(ats.LOT, 1, 1)) = 1 
                            and ats.devrevstep = idriss.devrevstep
                            and ats.WAFER_ID = idriss.WAFER_ID
                            and ats.latest_flag = 'Y'
                            and ats.WAFER_SCRIBE IS NOT NULL
                            group by ats.WAFER_ID, ats.devrevstep, ats.WAFER_SCRIBE, idriss.Touchdowns
                            order by ats.devrevstep, ats.WAFER_ID
                            ) afs,
                            a_testing_session   ats
                        where
                            ats.WAFER_SCRIBE = afs.WAFER_SCRIBE
                            """ + devrevstep + """
                            and ats.latest_flag   = 'Y'
                            order by ats.devrevstep, ats.WAFER_ID
                        ) table
                    group by table.devrevstep, table.WAFER_ID, table.Touchdowns, table.SORT_LOT
                    order by table.devrevstep, table.WAFER_ID
                    ) alltable
                where alltable.WAFER_ID not in
                    (
                    select NAtable.WAFER_ID
                    from
                        (
                        select table.devrevstep, table.WAFER_ID, table.Touchdowns, table.SORT_LOT
                        from (
                            select  ats.devrevstep, ats.WAFER_ID, afs.Touchdowns, CASE WHEN (ISNUMERIC(SUBSTRING(ats.LOT, 1, 1))!=1) THEN ats.LOT ELSE 'LOT_NOT_FOUND' END as SORT_LOT, ats.WAFER_SCRIBE
                            from (
                                select ats.WAFER_SCRIBE, ats.WAFER_ID, ats.devrevstep, idriss.Touchdowns
                                from a_testing_session   ats,
                                    (
                                        select ats.WAFER_ID, ats.devrevstep, COUNT(ats.WAFER_ID) as Touchdowns
                                        from a_testing_session   ats
                                        where ISNUMERIC(SUBSTRING(ats.LOT, 1, 1)) = 1 
                                        """ + devrevstep + """
                                        and   ats.latest_flag = 'Y'
                                        group by ats.WAFER_ID, ats.devrevstep
                                        order by ats.devrevstep, ats.WAFER_ID
                                    ) idriss

                                where
                                ISNUMERIC(SUBSTRING(ats.LOT, 1, 1)) = 1 
                                and ats.devrevstep = idriss.devrevstep
                                and ats.WAFER_ID = idriss.WAFER_ID
                                and ats.latest_flag = 'Y'
                                and ats.WAFER_SCRIBE IS NOT NULL
                                group by ats.WAFER_ID, ats.devrevstep, ats.WAFER_SCRIBE, idriss.Touchdowns
                                order by ats.devrevstep, ats.WAFER_ID
                                ) afs,
                                
                                a_testing_session   ats
                            where
                            ats.WAFER_SCRIBE = afs.WAFER_SCRIBE
                            """ + devrevstep + """
                            and ats.latest_flag   = 'Y'
                            order by ats.devrevstep, ats.WAFER_ID
                            ) table
                        where table.SORT_LOT != 'LOT_NOT_FOUND'
                        group by table.devrevstep, table.WAFER_ID, table.Touchdowns, table.SORT_LOT
                        order by table.devrevstep, table.WAFER_ID
                        )	Dtable,
                        (
                        select table.devrevstep, table.WAFER_ID, table.Touchdowns, table.SORT_LOT
                        from (
                            select  ats.devrevstep, ats.WAFER_ID, afs.Touchdowns, CASE WHEN (ISNUMERIC(SUBSTRING(ats.LOT, 1, 1))!=1) THEN ats.LOT ELSE 'LOT_NOT_FOUND' END as SORT_LOT, ats.WAFER_SCRIBE
                            from (
                                select ats.WAFER_SCRIBE, ats.WAFER_ID, ats.devrevstep, idriss.Touchdowns
                                from a_testing_session   ats,
                                    (
                                        select ats.WAFER_ID, ats.devrevstep, COUNT(ats.WAFER_ID) as Touchdowns
                                        from a_testing_session   ats
                                        where ISNUMERIC(SUBSTRING(ats.LOT, 1, 1)) = 1 
                                        """ + devrevstep + """
                                        and   ats.latest_flag = 'Y'
                                        group by ats.WAFER_ID, ats.devrevstep
                                        order by ats.devrevstep, ats.WAFER_ID
                                    ) idriss

                                where
                                ISNUMERIC(SUBSTRING(ats.LOT, 1, 1)) = 1 
                                and ats.devrevstep = idriss.devrevstep
                                and ats.WAFER_ID = idriss.WAFER_ID
                                and ats.latest_flag = 'Y'
                                and ats.WAFER_SCRIBE IS NOT NULL
                                group by ats.WAFER_ID, ats.devrevstep, ats.WAFER_SCRIBE, idriss.Touchdowns
                                order by ats.devrevstep, ats.WAFER_ID
                                ) afs,
                                
                                a_testing_session   ats
                            where
                            ats.WAFER_SCRIBE = afs.WAFER_SCRIBE
                            """ + devrevstep + """
                            and ats.latest_flag   = 'Y'
                            order by ats.devrevstep, ats.WAFER_ID
                            ) table
                        where table.SORT_LOT = 'LOT_NOT_FOUND'
                        group by table.devrevstep, table.WAFER_ID, table.Touchdowns, table.SORT_LOT
                        order by table.devrevstep, table.WAFER_ID
                        ) NAtable
                    where NAtable.WAFER_ID in Dtable.WAFER_ID
                    group by NAtable.devrevstep, NAtable.WAFER_ID, NAtable.Touchdowns, NAtable.SORT_LOT
                    order by NAtable.devrevstep, NAtable.WAFER_ID
                    )
                or alltable.SORT_LOT != 'LOT_NOT_FOUND'
                group by alltable.devrevstep, alltable.WAFER_ID, alltable.Touchdowns, alltable.SORT_LOT
                order by alltable.devrevstep, alltable.WAFER_ID

                </SQL >
                </analysis>""")

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

    if multi == 0:
        devrevstep = "and ats.devrevstep like '" + product_list[0] + "'"
    else:
        devrevstep = "and (ats.devrevstep like '" + product_list[0] + "'"
        for j in range(1,len(product_list)):
            devrevstep+= "or ats.devrevstep like '" + product_list[j] + "'"
        devrevstep+= ")"

    script_name =  sub(r"\s", "_", str(name))

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'idriss.animashaun@intel.com'
    mail.Subject = 'Please add report to automated scripts'
    mail.Body = f"""    Script Name: {script_name}
    Send Email To: {report_email}\n
    Wafer/s:  {split_wfr_list}\n
    <analysis app=cb
                TOOL=RUNSQL
                SCHEMA=ARIES
                SITE="""+ str(site) + """
                OUTPUT=
                /ASMERLIN
                <SQL >
                select alltable.devrevstep, alltable.WAFER_ID, alltable.Touchdowns, alltable.SORT_LOT
                from
                    (
                    select table.devrevstep, table.WAFER_ID, table.Touchdowns, table.SORT_LOT
                    from (
                        select  ats.devrevstep, ats.WAFER_ID, afs.Touchdowns, CASE WHEN (ISNUMERIC(SUBSTRING(ats.LOT, 1, 1))!=1) THEN ats.LOT ELSE 'LOT_NOT_FOUND' END as SORT_LOT, ats.WAFER_SCRIBE
                        from (
                            select ats.WAFER_SCRIBE, ats.WAFER_ID, ats.devrevstep, idriss.Touchdowns
                            from a_testing_session   ats,
                                (
                                    select ats.WAFER_ID, ats.devrevstep, COUNT(ats.WAFER_ID) as Touchdowns
                                    from a_testing_session   ats
                                    where ISNUMERIC(SUBSTRING(ats.LOT, 1, 1)) = 1 
                                    """ + devrevstep + """
                                    and   ats.latest_flag = 'Y'
                                    group by ats.WAFER_ID, ats.devrevstep
                                    order by ats.devrevstep, ats.WAFER_ID
                                ) idriss

                            where
                            ISNUMERIC(SUBSTRING(ats.LOT, 1, 1)) = 1 
                            and ats.devrevstep = idriss.devrevstep
                            and ats.WAFER_ID = idriss.WAFER_ID
                            and ats.latest_flag = 'Y'
                            and ats.WAFER_SCRIBE IS NOT NULL
                            group by ats.WAFER_ID, ats.devrevstep, ats.WAFER_SCRIBE, idriss.Touchdowns
                            order by ats.devrevstep, ats.WAFER_ID
                            ) afs,
                            a_testing_session   ats
                        where
                            ats.WAFER_SCRIBE = afs.WAFER_SCRIBE
                            """ + devrevstep + """
                            and ats.latest_flag   = 'Y'
                            order by ats.devrevstep, ats.WAFER_ID
                        ) table
                    group by table.devrevstep, table.WAFER_ID, table.Touchdowns, table.SORT_LOT
                    order by table.devrevstep, table.WAFER_ID
                    ) alltable
                where alltable.WAFER_ID not in
                    (
                    select NAtable.WAFER_ID
                    from
                        (
                        select table.devrevstep, table.WAFER_ID, table.Touchdowns, table.SORT_LOT
                        from (
                            select  ats.devrevstep, ats.WAFER_ID, afs.Touchdowns, CASE WHEN (ISNUMERIC(SUBSTRING(ats.LOT, 1, 1))!=1) THEN ats.LOT ELSE 'LOT_NOT_FOUND' END as SORT_LOT, ats.WAFER_SCRIBE
                            from (
                                select ats.WAFER_SCRIBE, ats.WAFER_ID, ats.devrevstep, idriss.Touchdowns
                                from a_testing_session   ats,
                                    (
                                        select ats.WAFER_ID, ats.devrevstep, COUNT(ats.WAFER_ID) as Touchdowns
                                        from a_testing_session   ats
                                        where ISNUMERIC(SUBSTRING(ats.LOT, 1, 1)) = 1 
                                        """ + devrevstep + """
                                        and   ats.latest_flag = 'Y'
                                        group by ats.WAFER_ID, ats.devrevstep
                                        order by ats.devrevstep, ats.WAFER_ID
                                    ) idriss

                                where
                                ISNUMERIC(SUBSTRING(ats.LOT, 1, 1)) = 1 
                                and ats.devrevstep = idriss.devrevstep
                                and ats.WAFER_ID = idriss.WAFER_ID
                                and ats.latest_flag = 'Y'
                                and ats.WAFER_SCRIBE IS NOT NULL
                                group by ats.WAFER_ID, ats.devrevstep, ats.WAFER_SCRIBE, idriss.Touchdowns
                                order by ats.devrevstep, ats.WAFER_ID
                                ) afs,
                                
                                a_testing_session   ats
                            where
                            ats.WAFER_SCRIBE = afs.WAFER_SCRIBE
                            """ + devrevstep + """
                            and ats.latest_flag   = 'Y'
                            order by ats.devrevstep, ats.WAFER_ID
                            ) table
                        where table.SORT_LOT != 'LOT_NOT_FOUND'
                        group by table.devrevstep, table.WAFER_ID, table.Touchdowns, table.SORT_LOT
                        order by table.devrevstep, table.WAFER_ID
                        )	Dtable,
                        (
                        select table.devrevstep, table.WAFER_ID, table.Touchdowns, table.SORT_LOT
                        from (
                            select  ats.devrevstep, ats.WAFER_ID, afs.Touchdowns, CASE WHEN (ISNUMERIC(SUBSTRING(ats.LOT, 1, 1))!=1) THEN ats.LOT ELSE 'LOT_NOT_FOUND' END as SORT_LOT, ats.WAFER_SCRIBE
                            from (
                                select ats.WAFER_SCRIBE, ats.WAFER_ID, ats.devrevstep, idriss.Touchdowns
                                from a_testing_session   ats,
                                    (
                                        select ats.WAFER_ID, ats.devrevstep, COUNT(ats.WAFER_ID) as Touchdowns
                                        from a_testing_session   ats
                                        where ISNUMERIC(SUBSTRING(ats.LOT, 1, 1)) = 1 
                                        """ + devrevstep + """
                                        and   ats.latest_flag = 'Y'
                                        group by ats.WAFER_ID, ats.devrevstep
                                        order by ats.devrevstep, ats.WAFER_ID
                                    ) idriss

                                where
                                ISNUMERIC(SUBSTRING(ats.LOT, 1, 1)) = 1 
                                and ats.devrevstep = idriss.devrevstep
                                and ats.WAFER_ID = idriss.WAFER_ID
                                and ats.latest_flag = 'Y'
                                and ats.WAFER_SCRIBE IS NOT NULL
                                group by ats.WAFER_ID, ats.devrevstep, ats.WAFER_SCRIBE, idriss.Touchdowns
                                order by ats.devrevstep, ats.WAFER_ID
                                ) afs,
                                
                                a_testing_session   ats
                            where
                            ats.WAFER_SCRIBE = afs.WAFER_SCRIBE
                            """ + devrevstep + """
                            and ats.latest_flag   = 'Y'
                            order by ats.devrevstep, ats.WAFER_ID
                            ) table
                        where table.SORT_LOT = 'LOT_NOT_FOUND'
                        group by table.devrevstep, table.WAFER_ID, table.Touchdowns, table.SORT_LOT
                        order by table.devrevstep, table.WAFER_ID
                        ) NAtable
                    where NAtable.WAFER_ID in Dtable.WAFER_ID
                    group by NAtable.devrevstep, NAtable.WAFER_ID, NAtable.Touchdowns, NAtable.SORT_LOT
                    order by NAtable.devrevstep, NAtable.WAFER_ID
                    )
                or alltable.SORT_LOT != 'LOT_NOT_FOUND'
                group by alltable.devrevstep, alltable.WAFER_ID, alltable.Touchdowns, alltable.SORT_LOT
                order by alltable.devrevstep, alltable.WAFER_ID

                </SQL >
                </analysis>"""

    mail.Send()

    print('Request Submitted')

def remove_zero():
    global split_wfr_list
    split_wfr_list = wfr.get().replace(" ", "").split(",")
    for k in range(0,len(split_wfr_list)):
        if split_wfr_list[k].startswith("0"):
            split_wfr_list[k]=split_wfr_list[k][1:]

def run_jrp():
    global output_file_csv

    run_wfr = 0

    user_script = "Output_touchdown\\" + file_name + ".jrp"
    if wfr.get().replace(" ", "") == "":
        jsl_path = resource_path("Inputs\\touchdown.jsl")
    else:
        run_wfr = 1
        jsl_path = resource_path("Inputs\\touchdown_wfr.jsl")
        remove_zero()
        wfr_IDs_list = json.dumps(split_wfr_list).strip("[]")
        print('wfr_IDs: ', wfr_IDs_list)


    reading_file = open(jsl_path, "r")

    new_file_content = ""
    for line in reading_file:
        stripped_line = line.strip()
        new_line = stripped_line.replace("C:\Scripts", output_file_csv)
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
    
    os.system(user_script)
    print('Analysis completed check Output_touchdown folder for data. JMP report will open automatically')

def store_val():
    prod_1 = prod_code.get()
    wafr_1 = wfr.get()
    e_mail_1 = input_email.get()
    store_info = [prod_1,wafr_1,e_mail_1]

    write_hist_file = open(hist_path, "w")
    for i in store_info:
        write_hist_file.write(str(i)+'\n')
    write_hist_file.close()

cblocation = cbilocator()

hist_path = resource_path("Inputs\\INFO.tmp");
read_hist_file = open(hist_path, "r")

prod = read_hist_file.readline().strip('\n')
wafr = read_hist_file.readline().strip('\n')
e_mail = read_hist_file.readline().strip('\n')

### Main Root
root = Tk()
root.title('Touchdown [Beta] v1.10')


mainframe = ttk.Frame(root, padding="60 50 60 50")
mainframe.grid(column=0, row=0, sticky=('news'))
mainframe.columnconfigure(0, weight=3)
mainframe.rowconfigure(0, weight=3)

def callback(url):
    webbrowser.open_new(url)

link1 = Label(mainframe, text="Wiki: https://goto/touchdown", fg="blue", cursor="hand2")
link1.grid(row = 0,column = 0, sticky=W, columnspan = 2)
link1.bind("<Button-1>", lambda e: callback("https://gitlab.devtools.intel.com/ianimash/touchdown/-/wikis/Touchdown"))

link2 = Label(mainframe, text="IT support contact: ricard.menchon.enrich@intel.com or idriss.animashaun@intel.com", fg="blue", cursor="hand2")
link2.grid(row = 1,column = 0, sticky=W, columnspan = 2)
link2.bind("<Button-1>", lambda e: callback("https://outlook.com"))

label_2 = Label(mainframe, text = 'Select Site: ', bg  ='black', fg = 'white')
label_2.grid(row = 1, column = 2, sticky=E)
variable = StringVar(mainframe)
variable.set("D1C") # default value

sel_prod = OptionMenu(mainframe, variable, "F28", "D1C", "D1D", "F32", "F24", "F68", "F21")

sel_prod.grid(row = 2, column = 2, sticky=W)

label_0 = Label(mainframe, text = 'Enter Product Code  [Use % for regex]: ', bg  ='black', fg = 'white')
label_0.grid(row = 2, sticky=E)
prod_code = Entry(mainframe, width=40, relief = FLAT)
prod_code.insert(4,prod)
prod_code.grid(row = 2, column = 1, sticky=W)

label_1 = Label(mainframe, text = 'Enter List of Wafer (Optional): ', bg  ='black', fg = 'white')
label_1.grid(row = 3, sticky=E)
wfr = Entry(mainframe, width=40, relief = FLAT)
wfr.insert(4,wafr)
wfr.grid(row = 3, column = 1, sticky=W)

label_2 = Label(mainframe, text = 'Enter Email For Weekly Reports: ', bg  ='black', fg = 'white')
label_2.grid(row = 4, sticky=E)
input_email = Entry(mainframe, width=40, relief = FLAT)
input_email.insert(4,e_mail)
input_email.grid(row = 4, column = 1, sticky=W)

button_0 = Button(mainframe, text="Pull Touchdowns", height = 1, width = 20, command = cbsql_basic, bg = 'green', fg = 'white', font = '-family "SF Espresso Shack" -size 12')
button_0.grid(row = 5, column = 0, sticky=E)

button_1 = Button(mainframe, text="Automated Weekly Reports", height = 1, width = 25, command = automate, bg = 'blue', fg = 'white', font = '-family "SF Espresso Shack" -size 12')
button_1.grid(row = 5, column = 1, sticky=W)

### Main loop
root.mainloop()