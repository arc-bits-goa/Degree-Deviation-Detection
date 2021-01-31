import xlrd
from collections import OrderedDict
import json
import os
import sys

def driver_noofcourse(filename):

    # Open the workbook and select the first worksheet
    wb = xlrd.open_workbook(os.path.join(filename))
    sh = wb.sheet_by_index(0)
    
    # List to hold dictionaries
    tag_list = []
    
    def RepresentsInt(s):
        try: 
            int(s)
            return int(s)
        except ValueError:
            return 0

    # Iterate through each row in worksheet and fetch values into dict
    for rownum in range(2, sh.nrows):
        tag = OrderedDict()
        row_values = sh.row_values(rownum)
        tag['Discipline'] = row_values[0]
        course = OrderedDict()
        course['CDC'] = RepresentsInt(row_values[1])
        course['DEL1'] = RepresentsInt(row_values[2])
        course['DEL2'] = RepresentsInt(row_values[3])
        course['HUM'] = RepresentsInt(row_values[4])
        course['OPEN'] = RepresentsInt(row_values[5])
        tag['No of Courses'] = course
    
        tag_list.append(tag)
    
    # Serialize the list of dicts to JSON
    j = json.dumps(tag_list)
    
    # Write to file
    with open(os.path.join('json','noofcourse.json'), 'w') as f:
        f.write(j)

    print("Finished executing noofcourse.py")