from suggest_courses import driver_suggest_courses
from find_lagging_students import driver_find_lagging_students
from jsontoxls_final import driver_jsontoxls_final
from logic import driver_logic
from jsontoxls_pre import driver_jsontoxls_pre
from studentdata import driver_studentdata
from makesort import driver_makesort
from noofcourse import driver_noofcourse
from check_absence_in_coursetype_excel import driver_check_absence_in_coursetype_excel
from coursetype import driver_coursetype
# import xlrd
# import xlwt
# import collections
# from collections import OrderedDict
# import json
# import os
# import sys
# import re
# import pandas


driver_coursetype('data/coursetype.xlsx')
driver_check_absence_in_coursetype_excel()
driver_noofcourse('data/noofcourse.xls')
driver_makesort()
driver_studentdata('data/sorted.xlsx')
driver_jsontoxls_pre()
driver_logic()
driver_jsontoxls_final()
driver_find_lagging_students()
driver_suggest_courses()