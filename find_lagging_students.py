import json
from collections import OrderedDict
import os
import pandas as pd

def driver_find_lagging_students():

    TIMETABLE_PATH = './data/Time Table.xlsx'
    CURRENT_SEM_REGISTRATION_DATA_PATH = './data/CURRENT_SEM_REGISTRATION_DATA.xls'

    #Open the JSON file containing description of courses arranged
    with open(os.path.join('json',"coursedesc.json")) as json_file:
        coursedesc_arr = json.load(json_file)

    #Open the JSON file containing number of courses in a particular discipline
    with open(os.path.join('json',"noofcourse.json")) as json_file:
        noofcourse = json.load(json_file)

    #Open the JSON file containing the student data
    with open(os.path.join('json','studentdatarf.json')) as json_file:
        studentdatarf = json.load(json_file)

    #Get the branch from the Campus ID given
    #eg.2018A7PS0123G
    def get_branch(s):
        btype = s[4:8] #eg.A7PS
        if (btype[2:4] == 'PS') or (btype[2:4] == 'TS'):
            return btype[0:2] #return A1, A7 etc or B5 etc for only MSc.eg.2018B5PS0123G
        if (btype[0] == 'A') and (btype[2] == 'B'):
            return (btype[2:4] + btype[0:2]) #reverese dual - eg.A7B5, still returns B5A7 since other input data is not present for A7B5, but is present for B5A7, and logic.py does not take the order of courses done into account, just the total number
        else:
            return btype #simple dual degree eg. B5A7, returns BE dual and MSc dual cases as well.

    def psts(s): #returns True if student is doing thesis
        btype = s[4:8] #eg.A7PS
        if (btype[2:4] == 'TS'):
            return True
        else:
            return False

    #Check for Reverse Dual/ BE Dual Cases and other special cases
    def check_special_case(cur_student_id): 
        btype = cur_student_id[4:8] #eg.A7PS
        if btype[2:3] == 'B':
            return True #true for cases like A7B5 - reverse dual. or Msc dual eg. B1B3
        elif btype[0:1] == 'A' and btype[2:3] == 'A':
            return True #BE dual
        elif btype[0:1] == 'B' and btype[2:4] == 'PS':
            return True #only MSc.
        else:
            return False
            


    #Get the course type by comparing it with the branch, like whether it is an elective/CDC
    #returns: POMPOE, OPEN, HUM, CDC, DEL1 or DEL2
    def getcoursetype(compcode, cid, branch):
        if (compcode == '21024') or (compcode == '21023'):
            return 'POMPOE'
        if (compcode == '21591') and (psts(cid)):
            return 'OPEN'
        try:
            coursedesc_arr[compcode]
        except:
            return 'OPEN'
        branch1 = branch[0:2]
        tag1 = list(filter(lambda x: branch1 in x, coursedesc_arr[compcode]['Tag']))
        branch2 = ''
        tag2 = []
        if len(branch) == 4:
            branch2 = branch[2:4]
            tag2 = list(filter(lambda x: branch2 in x, coursedesc_arr[compcode]['Tag']))
        tag3 = list(filter(lambda x: 'HUM' in x, coursedesc_arr[compcode]['Tag']))
        if not tag1 and not tag2 and not tag3:
            return 'OPEN'
        elif tag3:
            return 'HUM'
        elif tag1 and (tag1[0][2:4] == 'CD'):
            return 'CDC'
        elif tag2 and (tag2[0][2:4] == 'CD'):
            return 'CDC'
        elif tag1 and tag1[0][2:4] == 'EL':
            return 'DEL1'
        elif tag2 and tag2[0][2:4] == 'EL':
            return 'DEL2'
        else:
            return 'OPEN'


    def get_semester_wise_course_count(path):
        semwise_courses_count_df = pd.read_excel(path)
        semwise_courses_count = {}
        for i in range(len(semwise_courses_count_df)):
            branch = str(semwise_courses_count_df['Discipline'][i])
            sem = str(semwise_courses_count_df['year'][i]).strip() + '-' + str(semwise_courses_count_df['Semester'][i]).strip()
            try:
                branch_dict = semwise_courses_count[branch]
            except:
                semwise_courses_count[branch] = {}
                branch_dict = semwise_courses_count[branch]
            branch_dict[sem] = {"DEL1": semwise_courses_count_df['DEL1'][i], "DEL2": semwise_courses_count_df['DEL2'][i], "HUM": semwise_courses_count_df['HUM'][i], "OPEN": semwise_courses_count_df['OPEN'][i]}
        return semwise_courses_count


    def get_semester_identifier(path):
        semester_identifier_df = pd.read_excel(path)
        semester_identifier_dict = {}
        for i in range(len(semester_identifier_df)):
            year = str(semester_identifier_df['year'][i])
            semester_identifier_dict[year] = {}
            semester_identifier_dict[year] [str(semester_identifier_df['1-1'][i])] = '1-1'
            semester_identifier_dict[year] [str(semester_identifier_df['1-2'][i])] = '1-2'
            semester_identifier_dict[year] [str(semester_identifier_df['2-1'][i])] = '2-1'
            semester_identifier_dict[year] [str(semester_identifier_df['2-2'][i])] = '2-2'
            semester_identifier_dict[year] [str(semester_identifier_df['3-1'][i])] = '3-1'
            semester_identifier_dict[year] [str(semester_identifier_df['3-2'][i])] = '3-2'
            semester_identifier_dict[year] [str(semester_identifier_df['4-1'][i])] = '4-1'
            semester_identifier_dict[year] [str(semester_identifier_df['4-2'][i])] = '4-2'
            semester_identifier_dict[year] [str(semester_identifier_df['5-1'][i])] = '5-1'
            semester_identifier_dict[year] [str(semester_identifier_df['5-2'][i])] = '5-2'
        return semester_identifier_dict



    def get_compcodes_list(i): #i in studentdatarf
        cur_student_id = i['Campus Id']
        list_of_compcodes = []
        for key, value in i['Courses'].items():
            for k in range(len(value)):
                compcode = str(i['Courses'][key][k]['Course Id'])
                list_of_compcodes.append(compcode)
        current_sem_registration_data_df = pd.read_excel(CURRENT_SEM_REGISTRATION_DATA_PATH)
        for i in range(len(current_sem_registration_data_df)):
            if(cur_student_id == current_sem_registration_data_df['Campus ID'][i].strip()):
                list_of_compcodes.append(str(current_sem_registration_data_df['Course ID'][i]).strip())
        return list_of_compcodes


    def get_current_sem():
        current_sem_registration_data_df = pd.read_excel(CURRENT_SEM_REGISTRATION_DATA_PATH)
        return str(current_sem_registration_data_df['Semester'][0]).strip()


    tag_list = []
    student_final_list = []
    student_lag_list = []

    semwise_courses_count = get_semester_wise_course_count('./data/Semester_wise_number_of_courses.xlsx')

    sem_identifier_dict = get_semester_identifier('./data/semester_identifier.xlsx')


    for i in studentdatarf: #go through each student
        
        if check_special_case(i['Campus Id']) == True:
            continue #skip special cases.
        current_student_dict = {}
        current_student_dict['Empl Id'] = str(i['Empl Id'])
        current_student_dict['Campus Id'] = i['Campus Id']
        current_student_dict['Name'] = i['Name'].strip()

        year_in_ID_number = i['Campus Id'][0:4] #eg.'2017'
        student_branch = get_branch(i['Campus Id'])
        current_semester = sem_identifier_dict[year_in_ID_number][get_current_sem()]
        current_student_dict['Current Semester'] = current_semester
        #eg. current_semester = '3-1'

        course_count_as_per_plan = semwise_courses_count[student_branch][current_semester]
        #eg. {"DEL1": 1, "DEL2": 0, "HUM": 3, "OPEN": 1}
        course_count_as_per_plan['DEL1'] = int(course_count_as_per_plan['DEL1'])
        course_count_as_per_plan['DEL2'] = int(course_count_as_per_plan['DEL2'])
        course_count_as_per_plan['HUM'] = int(course_count_as_per_plan['HUM'])
        course_count_as_per_plan['OPEN'] = int(course_count_as_per_plan['OPEN'])
        current_student_dict['course_count_as_per_plan'] = course_count_as_per_plan

        current_student_dict['Courses'] = {"DEL1": 0, "DEL2": 0, "HUM": 0, "OPEN":0}
        current_student_dict['LAG'] = {"DEL1": 0, "DEL2": 0, "HUM": 0, "OPEN":0}

        tag = OrderedDict()
        tag['Empl Id'] = i['Empl Id']
        tag['Campus Id'] = i['Campus Id']
        tag['Name'] = i['Name']
        POMPOE = 0
        for j in noofcourse:
            if (j['Discipline'] == get_branch(i['Campus Id'])):
                CDC_REQ = j['No of Courses']['CDC']
                DEL1_REQ = j['No of Courses']['DEL1']
                DEL2_REQ = j['No of Courses']['DEL2']
                HUM_REQ = j['No of Courses']['HUM']
                OPEN_REQ = j['No of Courses']['OPEN']
                break
        CDC_LEFT = CDC_REQ
        if psts(i['Campus Id']):
            CDC_LEFT = CDC_LEFT - 1
        DEL1_LEFT = DEL1_REQ
        DEL2_LEFT = DEL2_REQ
        HUM_LEFT = HUM_REQ
        OPEN_LEFT = OPEN_REQ

        list_of_compcodes = get_compcodes_list(i)
        current_student_dict['completed_or_registered_courses_list'] = list_of_compcodes

        for compcode in list_of_compcodes:
            coursetype = getcoursetype(compcode, i['Campus Id'], get_branch(i['Campus Id']))
            #coursetype = POMPOE, OPEN, HUM, CDC, DEL1 or DEL2
            try:
                if tag[coursetype]: #person has done this coursetype before (or in current sem)
                    if compcode in tag[coursetype]:
                        REP_FLAG = 0 
                        # 0 actually means the course is seen again during execution of program
                    else:
                        REP_FLAG = 1
            except:
                REP_FLAG = 1 
            # REP_FLAG is not to be confused with student repeating course because of NC for example.

            #REP_FLAG = 1 means encountering course for first time for this student
            if REP_FLAG == 1:

                if coursetype == 'CDC':
                    CDC_LEFT = CDC_LEFT - 1
                elif (coursetype == 'HUM' and HUM_LEFT <=0) or (coursetype == 'DEL1' and DEL1_LEFT<=0) or (coursetype == 'DEL2' and DEL2_LEFT<=0 and DEL2_REQ!=0) or (coursetype == 'OPEN'):
                    OPEN_LEFT = OPEN_LEFT - 1
                    coursetype = 'OPEN' #imp!!
                elif coursetype == 'DEL1':
                    DEL1_LEFT = DEL1_LEFT - 1
                elif coursetype == 'DEL2':
                    DEL2_LEFT = DEL2_LEFT - 1
                elif coursetype == 'HUM':
                    HUM_LEFT = HUM_LEFT - 1
                elif (coursetype == 'POMPOE') and get_branch(i['Campus Id'])[0:2]!='B3': #non-eco students
                    if POMPOE == 1:
                        OPEN_LEFT = OPEN_LEFT - 1
                        coursetype = 'OPEN'
                    else:
                        POMPOE = 1
                        CDC_LEFT = CDC_LEFT - 1
                        coursetype = 'CDC'

                elif (coursetype == 'POMPOE') and get_branch(i['Campus Id'])[0:2]=='B3': #eco students
                    CDC_LEFT = CDC_LEFT - 1
                    coursetype = 'CDC'

                #coursetype = OPEN, HUM, CDC, DEL1 or DEL2
                #not checking CDC
                if coursetype != "CDC":
                    current_student_dict['Courses'][coursetype] += 1

                try:
                    tag[coursetype].append(compcode)
                except:
                    tag[coursetype] = [compcode]

        
        #FINDING LAG:
        DEL1_LAG_FLAG = 0 
        DEL2_LAG_FLAG = 0
        HUM_LAG_FLAG = 0
        OPEN_LAG_FLAG = 0
        
        if(course_count_as_per_plan['DEL1'] > current_student_dict['Courses']['DEL1']):
            DEL1_LAG_FLAG = 1
            current_student_dict['LAG']['DEL1'] = int(course_count_as_per_plan['DEL1'] - current_student_dict['Courses']['DEL1'])

        if(course_count_as_per_plan['DEL2'] > current_student_dict['Courses']['DEL2']):
            DEL2_LAG_FLAG = 1
            current_student_dict['LAG']['DEL2'] = int(course_count_as_per_plan['DEL2'] - current_student_dict['Courses']['DEL2'])

        if(course_count_as_per_plan['HUM'] > current_student_dict['Courses']['HUM']):
            HUM_LAG_FLAG = 1
            current_student_dict['LAG']['HUM'] = int(course_count_as_per_plan['HUM'] - current_student_dict['Courses']['HUM'])

        if(course_count_as_per_plan['OPEN'] > current_student_dict['Courses']['OPEN']):
            OPEN_LAG_FLAG = 1
            current_student_dict['LAG']['OPEN'] = int(course_count_as_per_plan['OPEN'] - current_student_dict['Courses']['OPEN'])

        student_final_list.append(current_student_dict)
        
        if(DEL1_LAG_FLAG==1 or DEL2_LAG_FLAG==1 or HUM_LAG_FLAG==1 or OPEN_LAG_FLAG==1):
            student_lag_list.append(current_student_dict)


    j = json.dumps(student_final_list)
    with open(os.path.join('json','student_course_count.json'), 'w') as f:
        f.write(j)

    j = json.dumps(student_lag_list)
    with open(os.path.join('json','student_lag_list.json'), 'w') as f:
        f.write(j)

    print("Finished executing find_lagging_students.py")