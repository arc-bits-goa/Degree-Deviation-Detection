import xlwt
import json
from collections import OrderedDict
import os

def driver_jsontoxls_pre():

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
    def branch(s):
        btype = s[4:8]
        if (btype[2:4] == 'PS') or (btype[2:4] == 'TS'):
            return btype[0:2]
        if (btype[0] == 'A') and (btype[2] == 'B'):
            return (btype[2:4] + btype[0:2])
        else:
            return btype

    def psts(s):
        btype = s[4:8]
        if (btype[2:4] == 'TS'):
            return True
        else:
            return False

    #Check for Reverse Dual/ BE Dual Cases
    def specialcase(s):
        btype = s[4:8]
        if btype[2:3] == 'B':
            return True


    #Get the course type by comparing it with the branch, like whether it is an elective/CDC
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

    #Get whether the particular subject is a project or not
    tag_list = []
    c = 0
    sheetno = 1

    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Sheet " + str(sheetno))
    sheet.write(0, 0, 'Empl Id')
    sheet.write(0, 1, 'Campus Id')
    sheet.write(0, 2, 'Name')
    sheet.write(0, 3, 'Semester')
    sheet.write(0, 4, 'Course Description')
    sheet.write(0, 5, 'Course Id')
    sheet.write(0, 6, 'Subject')
    sheet.write(0, 7, 'Catalog No')
    sheet.write(0, 8, 'Unit Taken')
    sheet.write(0, 9, 'Course Grade')
    sheet.write(0, 10, 'Tag')

    for i in studentdatarf:
        tag = OrderedDict()
        coursetype_out = ''
        tag['Empl Id'] = i['Empl Id']
        tag['Campus Id'] = i['Campus Id']
        tag['Name'] = i['Name']
        PROJ_LEFT = 5
        PROJ_LIST = {}
        PROJ_FLAG = 0
        ELEC_FLAG = 0
        POMPOE = 0
        PS_FLAG = 0
        for j in noofcourse:
            if (j['Discipline'] == branch(i['Campus Id'])):
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
        for key, value in i['Courses'].items():
            for k in range(len(value)):
                coursecode = str(i['Courses'][key][k]['Subject']) + " " + str(i['Courses'][key][k]['Catalog No'])
                compcode = str(i['Courses'][key][k]['Course Id'])
                coursetype = getcoursetype(compcode, i['Campus Id'], branch(i['Campus Id']))
                try:
                    if tag[coursetype]:
                        if compcode in tag[coursetype]:
                            REP_FLAG = 0
                        else:
                            REP_FLAG = 1
                except:
                    REP_FLAG = 1

                if REP_FLAG == 1:
                    if coursetype == 'CDC':
                        CDC_LEFT = CDC_LEFT - 1
                        coursetype_out = 'CDC'
                    elif (coursetype == 'HUM' and HUM_LEFT <=0) or (coursetype == 'DEL1' and DEL1_LEFT<=0) or (coursetype == 'DEL2' and DEL2_LEFT<=0 and DEL2_REQ!=0) or (coursetype == 'OPEN'):
                        OPEN_LEFT = OPEN_LEFT - 1
                        coursetype_out = 'OPEN'
                    elif coursetype == 'DEL1':
                        DEL1_LEFT = DEL1_LEFT - 1
                        coursetype_out = 'DEL1'
                    elif coursetype == 'DEL2':
                        DEL2_LEFT = DEL2_LEFT - 1
                        coursetype_out = 'DEL2'
                    elif coursetype == 'HUM':
                        HUM_LEFT = HUM_LEFT - 1
                        coursetype_out = 'HUM'
                    elif (coursetype == 'POMPOE') and branch(i['Campus Id'])[0:2]!='B3':
                        if POMPOE == 1:
                            OPEN_LEFT = OPEN_LEFT - 1
                            coursetype_out = 'OPEN'
                        else:
                            POMPOE = 1
                            CDC_LEFT = CDC_LEFT - 1
                            coursetype_out = 'CDC'

                    elif (coursetype == 'POMPOE') and branch(i['Campus Id'])[0:2]=='B3':
                        CDC_LEFT = CDC_LEFT - 1
                        coursetype_out = 'CDC'

                    try:
                        tag[coursetype].append(compcode)
                    except:
                        tag[coursetype] = [compcode]
                    

                    #Extra Flags as mentioned by ARC
                    if (ELEC_FLAG==0) and ((coursetype == 'HUM') or (coursetype == 'DEL1') or (coursetype == 'DEL2')) and coursedesc_arr[compcode]['Units'] < 3:
                        ELEC_FLAG = 1
                        
                    c = c + 1
                    if c>=65535:
                        c=0
                        sheetno = sheetno + 1
                        sheet = workbook.add_sheet("Sheet " + str(sheetno))
                    sheet.write(c, 0, i['Empl Id'])
                    sheet.write(c, 1, i['Campus Id'])
                    sheet.write(c, 2, i['Name'])
                    sheet.write(c, 3, key)
                    sheet.write(c, 4, i['Courses'][key][k]['Course Desc'])
                    sheet.write(c, 5, i['Courses'][key][k]['Course Id'])
                    sheet.write(c, 6, '')
                    sheet.write(c, 7, (i['Courses'][key][k]['Subject']) + " " + str(i['Courses'][key][k]['Catalog No']))
                    sheet.write(c, 8, i['Courses'][key][k]['Unit Taken'])
                    sheet.write(c, 9, i['Courses'][key][k]['Course Grade'])
                    sheet.write(c, 10, coursetype_out)



    # u = json_data
    # c = 0
    # emplid = ''
    # for i in u:
    #     c = c + 1
    #     sheet.write(c, 0, i['Empl Id'])
    #     sheet.write(c, 1, i['Campus Id'])
    #     sheet.write(c, 2, i['Name'])
    #     sheet.write(c, 3, i['Description'])
    #     sheet.write(c, 4, i['Course Id'])
    #     sheet.write(c, 5, i['Subject'])
    #     sheet.write(c, 6, i['Catalog No'].split())
    #     sheet.write(c, 7, i['Unit Taken'])
    #     sheet.write(c, 8, i['Course Grade'])
    #     compcode = str(i['Course Id'])
    #     print(compcode)
    #     coursetype =

    #     sheet.write(c, 9, coursetype)
    #     if c>=65535:
    #         c=0
    #         print("New sheet")
    #         sheetno = sheetno + 1
    #         sheet.col(0).width = 256 * 15
    #         sheet.col(1).width = 256 * 15
    #         sheet.col(2).width = 256 * 40
    #         sheet.col(3).width = 256 * 5
    #         sheet.col(4).width = 256 * 6
    #         sheet.col(5).width = 256 * 10
    #         sheet.col(6).width = 256 * 7
    #         sheet.col(7).width = 256 * 5
    #         sheet.col(8).width = 256 * 5
    #         sheet = workbook.add_sheet("Sheet " + str(sheetno))
    #         sheet.write(0, 0, 'Empl Id')
    #         sheet.write(0, 1, 'Campus Id')
    #         sheet.write(0, 2, 'Name')
    #         sheet.write(0, 3, 'Description')
    #         sheet.write(0, 4, 'Course Id')
    #         sheet.write(0, 5, 'Subject')
    #         sheet.write(0, 6, 'Catalog No')
    #         sheet.write(0, 7, 'Unit Taken')
    #         sheet.write(0, 8, 'Course Grade')
    #         sheet.write(0, 9, 'Tag')

    # sheet.col(0).width = 256 * 15
    # sheet.col(1).width = 256 * 15
    # sheet.col(2).width = 256 * 40
    # sheet.col(3).width = 256 * 5
    # sheet.col(4).width = 256 * 6
    # sheet.col(5).width = 256 * 10
    # sheet.col(6).width = 256 * 7
    # sheet.col(7).width = 256 * 5
    # sheet.col(8).width = 256 * 5

    workbook.save(os.path.join('result',"final_tag.xls"))

    print('Finished executing jsontoxls_pre.py')