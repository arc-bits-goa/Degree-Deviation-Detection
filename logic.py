import json
from collections import OrderedDict
import os

def driver_logic():

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

	# #Check for special cases like Reverse Dual/ BE Dual Cases
	# def specialcase(s): 
	# 	btype = s[4:8] 
	# 	if btype[2:3] == 'B':
	# 		return True 


	#Get the course type by comparing it with the branch, like whether it is an elective/CDC
	#returns: POMPOE, OPEN, HUM, CDC, DEL1 or DEL2
	def getcoursetype(compcode, cid, branch):
		if (compcode == '21024') or (compcode == '21023'):
			return 'POMPOE'
		if (compcode == '21591') and (psts(cid)): #practice school for thesis student
			return 'OPEN' 
		try:
			coursedesc_arr[compcode]
		except:
			return 'OPEN' #open elective if course type not mentioned by excel file
		branch1 = branch[0:2] #eg.A7, B5, A3
		tag1 = list(filter(lambda x: branch1 in x, coursedesc_arr[compcode]['Tag']))
		#eg. tag1 = ['A7CDC']
		branch2 = ''
		tag2 = []
		if len(branch) == 4:
			branch2 = branch[2:4] #eg.A7 as the second degree
			tag2 = list(filter(lambda x: branch2 in x, coursedesc_arr[compcode]['Tag']))
		tag3 = list(filter(lambda x: 'HUM' in x, coursedesc_arr[compcode]['Tag']))
		#eg. tag3 = ['HUM']
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
	def proj(compcode):
		try:
			coursedesc_arr[compcode]
		except:
			return False
		return coursedesc_arr[compcode]['Project']

	tag_list = []

	for i in studentdatarf:
		tag = OrderedDict()
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
		for key, value in i['Courses'].items():
			#key is semester number like "1142"
			#value is a list, having one dict per course in that sem
			for k in range(len(value)): #len(value) is number of courses in that sem
				coursecode = str(i['Courses'][key][k]['Subject']) + " " + str(i['Courses'][key][k]['Catalog No']) #eg. "BIO F111"
				compcode = str(i['Courses'][key][k]['Course Id']) #eg. "21002"
				coursetype = getcoursetype(compcode, i['Campus Id'], get_branch(i['Campus Id']))
				#coursetype = POMPOE, OPEN, HUM, CDC, DEL1 or DEL2
				try:
					if tag[coursetype]: #person has done this coursetype before (or in current sem)
						if compcode in tag[coursetype]:
							REP_FLAG = 0 
							# 0 actually means the course has been seen before for this student
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
						OPEN_LEFT = OPEN_LEFT - 1 #extra HUM, DEL counts as OPEN
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

					try:
						tag[coursetype].append(compcode)
					except:
						tag[coursetype] = [compcode]
					
					#Extra Flags as mentioned by ARC
					if proj(compcode):
						try:
							PROJ_LIST[compcode] = PROJ_LIST[compcode] + 1
						except:
							PROJ_LIST[compcode] = 1
					
					if ((coursetype == 'HUM') or (coursetype == 'DEL1') or (coursetype == 'DEL2')) and coursedesc_arr[compcode]['Units'] < 3:
						ELEC_FLAG = 1 


		for key,value in PROJ_LIST.items():
			if PROJ_LEFT > 0 and PROJ_FLAG != 1:
				PROJ_LEFT -= value
			
			if (PROJ_LEFT <= 0) or (value >= 3):
				PROJ_FLAG = 1 


		tag['CDCs Left'] = CDC_LEFT
		tag['DEL1s Left'] = DEL1_LEFT
		tag['DEL2s Left'] = DEL2_LEFT
		tag['OPENs Left'] = OPEN_LEFT
		tag['HUMs Left'] = HUM_LEFT
		tag['PROJ Flag'] = PROJ_FLAG
		tag['ELEC Flag'] = ELEC_FLAG

		tag_list.append(tag)
		
	# Serialize the list of dicts to JSON
	j = json.dumps(tag_list)
	
	# Write to file
	with open(os.path.join('json','finaldata.json'), 'w') as f:
		f.write(j)

	print("Finished executing logic.py")
