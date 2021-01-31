import xlwt
import json
import os

def driver_jsontoxls_final():
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Final")
    sheet.write(0, 0, 'Empl Id')
    sheet.write(0, 1, 'Campus Id')
    sheet.write(0, 2, 'Name')
    sheet.write(0, 3, 'CDC Pending')
    sheet.write(0, 4, 'DEL1 Pending')
    sheet.write(0, 5, 'DEL2 Pending')
    sheet.write(0, 6, 'HUM Pending')
    sheet.write(0, 7, 'OPEN Pending')
    sheet.write(0, 8, 'PROJ Flag')
    sheet.write(0, 9, 'ELEC Flag')

    with open(os.path.join('json',"finaldata.json")) as json_file:
        json_data = json.load(json_file)
    u = json_data
    c = 0
    for i in u:
        c = c + 1
        sheet.write(c, 0, i['Empl Id'])
        sheet.write(c, 1, i['Campus Id'])
        sheet.write(c, 2, i['Name'])
        sheet.write(c, 3, i['CDCs Left'])
        sheet.write(c, 4, i['DEL1s Left'])
        sheet.write(c, 5, i['DEL2s Left'])
        sheet.write(c, 6, i['HUMs Left'])
        sheet.write(c, 7, i['OPENs Left'])
        sheet.write(c, 8, i['PROJ Flag'])
        sheet.write(c, 9, i['ELEC Flag'])

    sheet.col(2).width = 256 * 50

    workbook.save(os.path.join('result',"final_pending_courses.xls"))

    print("Finished executing jsontoxls_final.py")
