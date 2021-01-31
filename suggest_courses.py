import json
import os
import pandas as pd

def driver_suggest_courses():

	TIMETABLE_PATH = './data/Time Table.xlsx'
	CURRENT_SEM_REGISTRATION_DATA_PATH = './data/CURRENT_SEM_REGISTRATION_DATA.xls'
	MAX_NUMBER_OF_SUGGESTED_COURSES = 100

	#Open the JSON file containing description of courses arranged
	with open(os.path.join('json',"coursedesc.json")) as json_file:
		coursedesc_arr = json.load(json_file)

	with open(os.path.join('json',"student_lag_list.json")) as json_file:
		student_lag_list = json.load(json_file)



	def get_branch(s):
		btype = s[4:8] #eg.A7PS
		if (btype[2:4] == 'PS') or (btype[2:4] == 'TS'):
			return btype[0:2] #return A1, A7 etc or B5 etc for only MSc.eg.2018B5PS0123G
		if (btype[0] == 'A') and (btype[2] == 'B'):
			return (btype[2:4] + btype[0:2]) #reverese dual eg.A7B5, still returns B5A7
		else:
			return btype #simple dual degree eg. B5A7


	def psts(s): #returns True if student is doing thesis
		btype = s[4:8] #eg.A7PS
		if (btype[2:4] == 'TS'):
			return True
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



	def get_lagging_course_types(current_student_dict):
		cur_lag_dict = current_student_dict['LAG']
		lagging_course_types = [] #eg. ['DEL1', 'HUM']
		for key, value in cur_lag_dict.items():
			if value > 0:
				lagging_course_types.append(key)
		return lagging_course_types



	def get_suitable_courses(lagging_course_types, cur_student_id,completed_courses_list):
		suitable_courses = {"DEL1": [], "DEL2": [], "HUM": [], "OPEN": []}
		timetable_df = pd.read_excel(TIMETABLE_PATH,sheet_name='erp time table' ,skiprows=1)
		for i in range(len(timetable_df)):
			cur_course_dict = {}
			cur_course_dict['Course ID'] = str(timetable_df['Course ID'][i]).strip()
			cur_course_dict['Subject'] = str(timetable_df['Subject'][i]).strip()
			cur_course_dict['Catalog'] = str(timetable_df['Catalog'][i]).strip()
			cur_course_dict['Course Title'] = str(timetable_df['Course Title'][i]).strip()
			cur_course_dict['Class Nbr'] = str(timetable_df['Class Nbr'][i]).strip()
			cur_course_type = getcoursetype(cur_course_dict['Course ID'],cur_student_id, get_branch(cur_student_id))
			#since we are not suggesting the OPEN courses, dont bother converting extra DELs or HUM into OPEN in this code
			#POMPOE will either be a CDC or OPEN, so skip it
			if cur_course_type == 'POMPOE':
				continue

			cur_course_dict['Course_type'] = cur_course_type
			if(cur_course_type in lagging_course_types):
				#check if previously done
				if cur_course_dict['Course ID'] not in completed_courses_list:
					suitable_courses[cur_course_type].append(cur_course_dict)
		return suitable_courses



	def append_to_simplified_excel_output(simplified_excel_output_df, current_student_dict):
		cur_empl_id = current_student_dict['Empl Id']
		cur_student_id = current_student_dict['Campus Id']
		cur_name = current_student_dict['Name']
		return simplified_excel_output_df.append(pd.Series([ cur_empl_id,cur_student_id,cur_name, current_student_dict['Current Semester'] ,current_student_dict['LAG']['DEL1'],current_student_dict['LAG']['DEL2'],current_student_dict['LAG']['HUM'],current_student_dict['LAG']['OPEN'] ], index=simplified_excel_output_df.columns),  ignore_index=True )


	def append_to_detailed_excel_output(detailed_excel_output_df, current_student_dict, suitable_courses, lagging_course_types):
		cur_empl_id = current_student_dict['Empl Id']
		cur_student_id = current_student_dict['Campus Id']
		cur_name = current_student_dict['Name']
		for cur_course_type, cur_list_of_course_dicts in suitable_courses.items():
			if cur_course_type not in lagging_course_types: 
				#don't append to output if not lagging
				continue
			if cur_course_type == 'OPEN': #don't bother about open elective suggestions
				continue
			list_to_append = [ cur_empl_id,cur_student_id,cur_name,cur_course_type]
			number_of_courses_found = 0
			for cur_course_dict in cur_list_of_course_dicts:
				list_to_append.append( cur_course_dict['Subject'].strip() + ' ' + cur_course_dict['Catalog'].strip())
				number_of_courses_found += 1
			while number_of_courses_found < MAX_NUMBER_OF_SUGGESTED_COURSES:
				list_to_append.append('-')
				number_of_courses_found += 1
			detailed_excel_output_df = detailed_excel_output_df.append(pd.Series(list_to_append, index=detailed_excel_output_df.columns),  ignore_index=True )
		return detailed_excel_output_df


	simplified_excel_output_df = pd.DataFrame(columns=['Empl Id', 'Campus Id', 'Name', 'Current Sem', 'DEL1_LAG', 'DEL2_LAG', 'HUM_LAG', 'OPEN_LAG'])

	detailed_excel_output_columns_list = ['Empl Id', 'Campus Id', 'Name', 'Course_type']
	for i in range(MAX_NUMBER_OF_SUGGESTED_COURSES):
		temp = 'option-' + str(i+1)
		detailed_excel_output_columns_list.append(temp)
	detailed_excel_output_df = pd.DataFrame(columns=detailed_excel_output_columns_list)


	for current_student_dict in student_lag_list:
		cur_empl_id = current_student_dict['Empl Id']
		cur_student_id = current_student_dict['Campus Id']
		cur_name = current_student_dict['Name']

		lagging_course_types = get_lagging_course_types(current_student_dict)

		suitable_courses = get_suitable_courses(lagging_course_types,cur_student_id,current_student_dict['completed_or_registered_courses_list'])

		simplified_excel_output_df = append_to_simplified_excel_output(simplified_excel_output_df, current_student_dict)

		detailed_excel_output_df = append_to_detailed_excel_output(detailed_excel_output_df, current_student_dict, suitable_courses, lagging_course_types)

	simplified_excel_output_df.to_excel('./result/simplified_lag_output.xlsx', index=False)
	detailed_excel_output_df.to_excel('./result/detailed_lag_output.xlsx', index=False)


	print("Finished executing suggest_courses.py")

	print("All python files executed. Please check result folder")