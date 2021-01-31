import json
import os
import pandas as pd

def driver_check_absence_in_coursetype_excel():
    TIMETABLE_PATH = './data/Time Table.xlsx'

    #Open the JSON file containing description of courses arranged
    with open(os.path.join('json',"coursedesc.json")) as json_file:
        coursedesc_arr = json.load(json_file)

    course_ids_present_in_coursedesc_json = list(coursedesc_arr.keys())

    visited_courses = []
    courses_absent_in_coursetype_excel_df = pd.DataFrame(columns=['Course ID', 'Subject', 'Catalog', 'Course Title'])

    #iterate through Timetable and check for presence in coursedesc.json
    timetable_df = pd.read_excel(TIMETABLE_PATH,sheet_name='erp time table' ,skiprows=1)
    for i in range(len(timetable_df)):
        cur_course_id = str(timetable_df['Course ID'][i]).strip()
        if cur_course_id in visited_courses:
            continue
        visited_courses.append(cur_course_id)
        if cur_course_id not in course_ids_present_in_coursedesc_json:
            courses_absent_in_coursetype_excel_df = courses_absent_in_coursetype_excel_df.append(pd.Series([cur_course_id,str(timetable_df['Subject'][i]).strip(),str(timetable_df['Catalog'][i]).strip(),str(timetable_df['Course Title'][i]).strip()], index=courses_absent_in_coursetype_excel_df.columns),  ignore_index=True )

    courses_absent_in_coursetype_excel_df.to_excel('./result/courses_absent_in_coursetype_excel.xlsx', index=False)

    print("Finished executing check_absence_in_coursetype_excel.py")
