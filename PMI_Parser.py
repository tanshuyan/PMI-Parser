# -*- coding: utf-8 -*-
"""
PMI Parser
Release Version 1.0

Created on Thu Mar 19 2018

@author: shu yan

----------------------------------------------------------------------------
How to compile into a .exe package using pyinstaller

Open command prompt, change directory to the folder that PMI_Parser.py is in, and do the following:

For PMI_Parser_standalone:
- set is_standalone to True in the code
- paste and run: pyinstaller -F --add-data "PMI alarms.xlsx;." PMI_Parser.py

For PMI_Parser_xlsx_required:
- set is_standalone to False in the code
- paste and run: pyinstaller -F PMI_Parser.py
"""

#Import os module
import sys, os
import pandas as pd
import pandas.io.formats.excel
import re
import datetime
import time
import xlrd
start = time.clock()
num_files = 0
#is_standalone = True if generating the standalone version (no accompanying "PMI alarms.xlsx" required; it is bundled into the .exe)
#is_standalone = False if generating the xlsx-required version
is_standalone = True

#Extracts the following information from the log file: MCCS side (A or B), type (Master or Slave), log file date and time 
#Input: file path
#Output: Tuple of (MCCS_Side, MCCS_Type, FileDateTime)
def get_logfile_properties(filename):
    re_logfile = re.compile(r"MCCS_([AB])_([EM])_(\d{2})_(\d{2})_(\d{4})_(\d{2})_(\d{2})_(\d{2})_(\d{3})\.[^/]*(?:txt|log)$")
    search_result = re_logfile.search(filename)
    if search_result:
        side, type, d, m, y, hr, min, sec, ms = search_result.groups()
        if type == "M":
            type = "Master"
        elif type == "E":
            type = "Slave"
        logfile_datetime = datetime.datetime(int(y), int(m), int(d), int(hr), int(min), int(sec), int(ms))
        return (side, type, logfile_datetime)
    else:
        #Throw some exception?
        pass
        return ("","",datetime.datetime(1,1,1))


#Searches a file object, fo
#Input: None
#Output: dataframe with search results
#Dependencies: fo (Type: file object), fname, line_no, line, alarm_df, MCCS_side ,MCCS_type ,logfile_datetime
def search_file():
    #time is defaulted to 0/0/0 00:00:00
    current_time = datetime.time(0,0,0)
        
    operations = 0
    
    
    fname_list = []
    line_no_list = []
    index_list = []
    line_list = []
    code_list = []
    chapter_list = []
    time_list = []
    mccs_side_list =[]
    mccs_type_list =[]
    logfile_datetime_list =[]

    # Read the first line from the file
    line = fo.readline()
    
    # Initialize counter for line number
    line_no = 1

    # Loop until EOF
    while line != '' :
        #search for time stamp here
        #declare a time regex in the main body and use it to find the time in here
        #Do a search for time stamp. Anytime a time stamp is found, update the current_time
        if re_time_stamp.search(line):
            current_time = pd.to_datetime(re_time_stamp.search(line).group(0), format='%d/%m/%y %H:%M:%S:%f')
        
        #search for restarts here
        search_result = re_restart.search(line)
        operations += 1
        if search_result:
            fname_list.append(fname)
            line_no_list.append(line_no)
            index_list.append(search_result.start())
            line_list.append(line)
            code_list.append(0)
            chapter_list.append("restart")
            time_list.append(current_time)
            mccs_side_list.append(MCCS_side)
            mccs_type_list.append(MCCS_type)
            logfile_datetime_list.append(logfile_datetime)
        #search for alarms here
        #We use .match() because the alarm stamp ("A") is guaranteed to be at the start of the line. This speeds up searching.
        if re_alarm_stamp.match(line):
            for code, regex_term, chapter in alarm_tuples_list:
                re_search_term = re.compile(regex_term)
                #if the alarm can be found inside the line, create a new entry and insert all relevant info
                search_result = re_search_term.search(line)
                operations += 1
                if search_result:
                    fname_list.append(fname)
                    line_no_list.append(line_no)
                    index_list.append(search_result.start())
                    line_list.append(line)
                    code_list.append(code)
                    chapter_list.append(chapter)
                    time_list.append(current_time)
                    mccs_side_list.append(MCCS_side)
                    mccs_type_list.append(MCCS_type)
                    logfile_datetime_list.append(logfile_datetime)        
        # Read next line
        line = fo.readline()

        # Increment line counter
        line_no += 1
        
    print(operations)

    output_df = pd.DataFrame({'File Path': fname_list, 'Log File Date Stamp':logfile_datetime_list, 'Side': mccs_side_list,
                    'Type':mccs_type_list, 'Line No.': line_no_list, 'Index':index_list, 'Matched term': line_list,
                    'Alarm code': code_list, 'Chapter': chapter_list, 'Alarm Time': time_list})
    return output_df
                
#Input: Directory Path
#Output: Returns a String array of the paths to each file in the directory path
def get_filepaths(directory):
    return [os.path.join(r,file) for r,d,f in os.walk(directory) for file in f]

#Main program starts here

#File types of log files to search for
file_type1 = ".log"
file_type2 = ".txt"
#PMI log time stamps are in the format dd/mm/yy HH:mm:ss:msms
re_time_stamp = re.compile(r"\d{2}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}:\d{2}")
#PMI alarms are always preceded by an "A" (this check speeds up processing times)
re_alarm_stamp = re.compile(r"^\d+ A")
#This is the marker for a PMI restart
re_restart = re.compile(r"Mode OPERATIONNEL")

#Create the path to the PMI alarms depending on whether the code is running in an .exe package (frozen) or not (not frozen)
if getattr(sys, 'frozen', False) and is_standalone:
    # if you are running in a |PyInstaller| bundle
    alarmsDir = sys._MEIPASS
    alarmsDir = os.path.join(alarmsDir, 'PMI alarms.xlsx') 
else:
    # we are running in a normal Python environment
    alarmsDir = os.getcwd()
    alarmsDir = os.path.join(alarmsDir, 'PMI alarms.xlsx') 

df = pd.DataFrame()
#Extract all PMI alarms into a dataframe
alarm_df = pd.read_excel(open(alarmsDir,'rb'), sheet_name='Sheet1')
alarm_tuples_list = list(alarm_df.itertuples(index=False))

full_file_paths = get_filepaths(os.getcwd())

#Access each .txt and .log in the directory and run a search on them for PMI alarms
for fname in full_file_paths:
   # Apply file type filter   
    if fname.endswith(file_type1) or fname.endswith(file_type2):
        num_files += 1

        #Get the MCCS side (A/B), type(Master/Slave), and logfile datetime stamp (All found in the file name) 
        MCCS_side ,MCCS_type ,logfile_datetime = get_logfile_properties(fname)

        # Open file for reading
        fo = open(fname, encoding="mac_roman")
        
        #here's the searching loop
        #loop through LIST of terms to be searched, each time calling the search_for function
        df = pd.concat([df,search_file()])
        
        # Close the files
        fo.close()

#if no alarms found, exit program
if df.empty:
    print("\nNo alarms found!")
else:
    #Arrange excel sheet columns in this order
    df = df[['File Path','Log File Date Stamp','Side','Type','Line No.','Index','Matched term','Alarm code', 'Chapter', 'Alarm Time']]
    df = df.sort_values(by=['Log File Date Stamp','Line No.'])


output_name = 'PMI alarms output ' + datetime.datetime.now().strftime("%Y_%m_%d_%H%M%S") + '.xlsx'

#Remove the default header style so that we can format it as we require later (word wrapping)
pandas.io.formats.excel.header_style = None
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(output_name, engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name='Sheet1')

#Add formatting
#Wrap and bold headers for readibility
workbook = writer.book
hdr_format = workbook.add_format()
hdr_format.set_text_wrap(True)
hdr_format.set_bold(True)
writer.sheets['Sheet1'].set_row(0, None, hdr_format)
#Bold Index column
index_format = workbook.add_format()
index_format.set_bold(True)
writer.sheets['Sheet1'].set_column('A:A', 2.3, index_format)
# Resize the column widths for readability
writer.sheets['Sheet1'].set_column('B:B', 25) #File Path
writer.sheets['Sheet1'].set_column('C:C', 17.57) #Log File Date Stamp
writer.sheets['Sheet1'].set_column('D:D', 4.7) #Side
writer.sheets['Sheet1'].set_column('E:E', 6.2) #Type
writer.sheets['Sheet1'].set_column('F:F', 5.3) #Line No.
writer.sheets['Sheet1'].set_column('G:G', 5.7) #Index
writer.sheets['Sheet1'].set_column('H:H', 75) #Matched Term
writer.sheets['Sheet1'].set_column('I:I', 6.14) #Alarm Code
writer.sheets['Sheet1'].set_column('J:J', 8.00) #Chapter
writer.sheets['Sheet1'].set_column('K:K', 17.6) #Alarm Time


# Close the Pandas Excel writer and output the Excel file.
writer.save()

print ('Time taken: %f seconds' %(time.clock() - start))
print ('Files analysed: %d' %(num_files))
input("Press Enter to continue...")