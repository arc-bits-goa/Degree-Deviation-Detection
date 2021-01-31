import pandas as pd

def driver_makesort():

    def load_files():
        from os import listdir
        from os.path import isfile, join
        files = [f for f in listdir("./data/reg_data") if isfile(join("./data/reg_data", f))]
        files.sort()
        registration_data=pd.DataFrame()
        for i in files:
            if i[-4:] == '.xls' or i[-5:] == '.xlsx':
                data = pd.read_excel("./data/reg_data/"+i,sheet_name="sheet1",header=1)
                registration_data=registration_data.append(data)
        registration_data = registration_data.sort_values(by =['Campus ID','Semester'] )
        return registration_data 


    def label_emplid(row):
        return '311'+row['Campus ID'][0:4]+row['Campus ID'][8:12]


    def editFormat(registration_data):
        registration_data['Empl Id'] = registration_data.apply (lambda row: label_emplid(row), axis=1)
        cols = registration_data.columns.tolist()
        cols = cols[-1:] + cols[:-1]
        registration_data = registration_data[cols]
        registration_data=registration_data.drop(['Acadmic Career', 'Description','Elective OR Audit Tag',], axis=1)
        return registration_data


    def load_students(df):
        data = pd.read_excel("./data/students.xlsx")
        students=list(data['Campus ID'])
        for i in range(len(students)):
            students[i]=str(students[i]).strip()

        for i in range(len(students)):
            students[i] = '311'+students[i][0:4]+students[i][8:12]

        df=df.loc[df['Empl Id'].isin(students)]
        return df


    reg_data=load_files() 
    #all students from "./data/reg_data", all semesters. 

    reg_data=editFormat(reg_data)
    # all students in the DF obtained above - 
    # added "Empl ID". Dropped ['Acadmic Career', 'Description','Elective OR Audit Tag']


    reg_data=load_students(reg_data)
    # get the data only for those students in ("./data/students.xlsx")

    reg_data.to_excel('./data/sorted.xlsx',index=False)

    print("Finished executing makesort.py")
