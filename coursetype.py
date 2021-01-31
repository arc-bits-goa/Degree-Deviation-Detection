import xlrd
from collections import OrderedDict
import json
import os
import sys

def driver_coursetype(filePath):
    def RepresentsInt(s):
        try: 
            int(s)
            return int(s)
        except ValueError:
            return 0

    def RepresentsBool(s):
        try: 
            a = True if s == 1 else False
            return a
        except ValueError:
            return True

    # Open the workbook and select the first worksheet
    wb = xlrd.open_workbook(os.path.join(filePath))
    sh = wb.sheet_by_index(0)
    
    # List to hold dictionaries
    tag_list = []
        
    # Iterate through each row in worksheet and fetch values into dict
    tag = OrderedDict()
    for rownum in range(1, sh.nrows):
        row_values = sh.row_values(rownum)
        if str(RepresentsInt(row_values[0])) in tag:
            tag[str(RepresentsInt(row_values[0]))]['Tag'].append(row_values[2])
        else:
            desc = OrderedDict()
            desc['Comp Codes'] = row_values[1]
            desc['Tag'] = [row_values[2]]
            desc['Course Name'] = row_values[3]
            desc['Project'] = RepresentsBool(row_values[4])
            desc['Units'] = RepresentsInt(row_values[5])
            tag[str(RepresentsInt(row_values[0]))] = desc

    tag_list.append(tag)
    
    # Serialize the list of dicts to JSON
    j = json.dumps(tag)
    
    # Write to file
    with open(os.path.join('json','coursedesc.json'), 'w') as f:
        f.write(j)

    print("Finished executing coursetype.py")

    # print(json.dumps(json.loads(j), indent=4, sort_keys=True))