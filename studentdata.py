import xlrd
from collections import OrderedDict
import json
import os
import sys
import re


 
def driver_studentdata(filename):
    # Open the workbook and select the first worksheet
    wb = xlrd.open_workbook(os.path.join(filename))
    sh = wb.sheet_by_index(0)
    
    # List to hold dictionaries
    tag_list = []
    backlog_list = []

    
    # Iterate through each row in worksheet and fetch values into dict
    cur_row = 1
    acceptable = ['', 'A', 'A-','B','B-','C','C-','D','E','Good','Poor']
    while(cur_row < sh.nrows):
        row_values = sh.row_values(cur_row)
        tag = OrderedDict()
        tagN= OrderedDict()
        tag['Empl Id'] = int(row_values[0])
        tag['Campus Id'] = row_values[1]
        tag['Name'] = re.sub('[.,]', '', row_values[2])
        tagN['Empl Id'] = row_values[0]
        tagN['Campus Id'] = row_values[1]
        tagN['Name'] = row_values[2]
        emplid = int(row_values[0])
        courses = OrderedDict()
        logs = OrderedDict()
        while(emplid == int(row_values[0]) and cur_row < sh.nrows):
            row_values = sh.row_values(cur_row)
            desc = str(int(row_values[3]))
            course_list = []
            log_list = []
            while (desc == str(int(row_values[3])) and cur_row < sh.nrows):
                row_values = sh.row_values(cur_row)
                course = OrderedDict()
                course['Course Id'] = int(row_values[4])
                course['Course Desc'] = row_values[5]
                course['Subject'] = row_values[6]
                course['Catalog No'] = row_values[7].strip()
                course['Unit Taken'] = row_values[8]
                course['Course Grade'] = row_values[9]
                if row_values[9] in acceptable or not row_values[9].strip():
                    course_list.append(course)
                else:
                    log_list.append(course)
                cur_row += 1
            courses[desc] = course_list
            if log_list:
                logs[desc] = log_list
        tag['Courses'] = courses
        tag_list.append(tag)
        tagN['Courses'] = logs
        if logs:
            backlog_list.append(tagN)
    

    # Serialize the list of dicts to JSON
    j = json.dumps(tag_list)
    k = json.dumps(backlog_list)
    
    # Write to file
    with open(os.path.join('json','studentdatarf.json'), 'w') as f:
        f.write(j)

    with open(os.path.join('json','studentdatalog.json'), 'w') as f:
        f.write(k)

    print("Finished executing studentdata.py")
