import gspread
from oauth2client.service_account import ServiceAccountCredentials
import math
from operator import itemgetter
from gspread.models import Cell


# https://docs.gspread.org/en/v3.7.0/


# Setup Connection and Credeentials
credentials = "credentials.json"
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',
						"https://www.googleapis.com/auth/drive.file",
						"https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(credentials, scope)
client = gspread.authorize(creds)


WritingDocSheet = "test2"
ReadingDocSheet = "12_july_Sprint3"

class Spreadsheet:

	def __init__(self,scope, creds, client):
		self.scope = scope
		self.creds = creds
		self.client = client


	def writeOnSheet(self, emails, names):
		spreadsheet = self.client.open("12_July")
		worksheet =  spreadsheet.add_worksheet("Emails", len(emails), 2)

		for i in range(1, len(emails)+1):
			worksheet.update_cell(i, 1, names[i-1])
			worksheet.update_cell(i, 2, emails[i-1])

		print("Writing Done")


	def readFromSheet(self):
		spreadsheet = self.client.open("12_July").worksheet("Emails")
		requestData = spreadsheet.get_all_records()
		print(requestData)

		requestData = spreadsheet.get_all_values()
		print(requestData)

		values_list = spreadsheet.row_values(1)
		print(values_list)

		cellValue = spreadsheet.acell("A2")
		print(cellValue)


		print("readingDone")



	def prepareCategoryReport(self, percentageOfAttempts, contest, k, worksheet):
		Category_A, Category_B, Category_C, Category_D, Category_E, Category_F = 0,0,0,0,0,0
		Categories_Name = ["Category_A", "Category_B", "Category_C", "Category_D", "Category_E", "Category_F"]


		for i in percentageOfAttempts:
			if i == 100: Category_A += 1
			if i<100 and i>=75: Category_B += 1
			if i<75 and i>=50: Category_C += 1
			if i>=25 and i<50: Category_D += 1
			if i>=10 and i<25: Category_E += 1
			if i<10: Category_F += 1

		Category_Values = list((Category_A, Category_B, Category_C, Category_D, Category_E, Category_F))
		print(Category_Values)


		worksheet.update_cell(k,1, contest)

		worksheet.update_cell(k+2,3, "Count Of Students")
		worksheet.update_cell(k+2,5, "Percentage Of Students")

		sumOfCategoryValues = sum(Category_Values)

		cells = []
		for i in range(1, len(Category_Values)+1):
			cells.append(Cell(row=i+k+2, col=1, value=Categories_Name[i-1]))
			cells.append(Cell(row=i+k+2, col=3, value=Category_Values[i-1]))
			percentageValue = str(math.ceil((Category_Values[i-1]/sumOfCategoryValues)*100))+" %"
			cells.append(Cell(row=i+k+2, col=5, value=percentageValue))

		worksheet.update_cells(cells)

		# for i in range(1, len(Category_Values)+1):
		#   worksheet.update_cell(i+k+2, 1, Categories_Name[i-1])
		#   worksheet.update_cell(i+k+2, 3, Category_Values[i-1])
			
		#   percentageValue = str(math.ceil((Category_Values[i-1]/sumOfCategoryValues)*100))+" %"
		#   worksheet.update_cell(i+k+2, 5, percentageValue)




	# Step 1 : Fetch the Regular Attempted Column in a list
	def generateReportForGroups(self, group, groupFileName):
		global ReadingDocSheet
		global WritingDocSheet

		spreadsheet = self.client.open(ReadingDocSheet).worksheet(groupFileName)

		# Spreadsheet for writing
		spreadsheetWriting = self.client.open(WritingDocSheet)

		# group = "Alpha_Report1"
		worksheet =  spreadsheetWriting.add_worksheet(group,1000, 1000)

		
		# Get Percentage of Regular Contest
		Regular_Accepted = spreadsheet.col_values(2)[1:]
		Regular_Accepted = list(map(int, Regular_Accepted))


		Regular_Attempted = spreadsheet.col_values(5)[1:]
		Regular_Attempted = list(map(int, Regular_Attempted))
		max_Attempted = int(max(Regular_Attempted))

		percentageOfAttempts = [math.ceil((int(x)/max_Attempted)*100) for x in Regular_Accepted]


		# Arrange in Categories
		percentageOfAttempts.sort()
		print("Regular ", percentageOfAttempts)

		# Regular Contest Done
		self.prepareCategoryReport(percentageOfAttempts, "Regular Contest Report", 2, worksheet)
		print("Regular Contest Report : Done")



		# Get Percentage of Timed Contest
		Timed_Accepted = spreadsheet.col_values(6)[1:]
		Timed_Accepted = list(map(int, Timed_Accepted))

		Timed_Attempted = spreadsheet.col_values(9)[1:]
		Timed_Attempted = list(map(int, Timed_Attempted))

		max_Attempted = int(max(Timed_Attempted))
		
		# To avoid division by zero
		# if max_Attempted == 0:
		# 	max_Attempted = 1

		percentageOfAttempts = [math.ceil((int(x)/max_Attempted)*100) for x in Timed_Accepted]


		# Arrange in Categories
		percentageOfAttempts.sort()
		print("Timed ",percentageOfAttempts)

		self.prepareCategoryReport(percentageOfAttempts, "Timed Contest Report", 14, worksheet)


	def generateReport_1ForAllGroups(self):

		# Venu Report
		cohort = "Cohort-14-Unit-3"
		self.generateReportForGroups("Report1", cohort)

		# Varun Report
		# self.generateReportForGroups("Charlie_Report1", "Cohort-13-Unit-5_charlie" )
		# self.generateReportForGroups("Delta_Report1", "Cohort-13-Unit-5_delta" )

		# Lohit Report
		# self.generateReportForGroups("Charlie_Report1", "Cohort-13-Unit-5_charlie" )
		# self.generateReportForGroups("Delta_Report1", "Cohort-13-Unit-5_delta" )



	def generateReport_3ForAllGroups(self):

		#Venu Report
		cohort = "Cohort-14-Unit-3"
		self.generateReport_3ForGroups("Report3", cohort )

		#Varun Report
		# self.generateReport_3ForGroups("Charlie_Report3", "Cohort-13-Unit-5_charlie")
		# self.generateReport_3ForGroups("Delta_Report3", "Cohort-13-Unit-5_delta")
		


		
	
	def generateReport_3ForGroups(self, group, groupFileName):

		global ReadingDocSheet
		global WritingDocSheet

		spreadsheet = self.client.open(ReadingDocSheet).worksheet(groupFileName)

		# Spreadsheet for writing
		spreadsheetWriting = self.client.open(WritingDocSheet)
		# group = "Alpha_Report1"
		worksheet =  spreadsheetWriting.add_worksheet(group,1000, 1000)

		
		# Get Percentage of Regular Contest
		Regular_Accepted = spreadsheet.col_values(2)[1:]
		Regular_Accepted = list(map(int, Regular_Accepted))


		Regular_Attempted = spreadsheet.col_values(5)[1:]
		Regular_Attempted = list(map(int, Regular_Attempted))
		max_Attempted = int(max(Regular_Attempted))



		# # To avoid max attempt as 0
		# if max_Attempted == 0:
		# 	max_Attempted = 1

		print("max max_Attempted",max_Attempted)
		print(Regular_Accepted)
		percentageOfAttempts = [math.ceil((int(x)/max_Attempted)*100) for x in Regular_Accepted]
		print("percentageOfAttempts ", percentageOfAttempts)
		categoryListRegular = []
		for i in percentageOfAttempts:
			if i == 100: categoryListRegular.append(1)
			if i<100 and i>=75: categoryListRegular.append(2)
			if i<75 and i>=50: categoryListRegular.append(3)
			if i>=25 and i<50: categoryListRegular.append(4)
			if i>=10 and i<25: categoryListRegular.append(5)
			if i<10: categoryListRegular.append(6)

		Timed_Accepted = spreadsheet.col_values(6)[1:]
		Timed_Accepted = list(map(int, Timed_Accepted))

		Timed_Attempted = spreadsheet.col_values(9)[1:]
		Timed_Attempted = list(map(int, Timed_Attempted))

		max_Attempted = int(max(Timed_Attempted))

		# To avoid division by zero
		# if max_Attempted == 0:
		# 	max_Attempted = 1

		percentageOfAttempts = [math.ceil((int(x)/max_Attempted)*100) for x in Timed_Accepted]

		categoryListTimed = []
		for i in percentageOfAttempts:
			if i == 100: categoryListTimed.append(1)
			if i<100 and i>=75: categoryListTimed.append(2)
			if i<75 and i>=50: categoryListTimed.append(3)
			if i>=25 and i<50: categoryListTimed.append(4)
			if i>=10 and i<25: categoryListTimed.append(5)
			if i<10: categoryListTimed.append(6)


		# Create Final Sorted arrangement of ID's, Name's and Categories
		StudentId = spreadsheet.col_values(1)[1:]
		StudentNames = spreadsheet.col_values(14)[1:]

		print("yoo",StudentNames)
		print(len(categoryListTimed)," ", len(categoryListRegular)," ", len(StudentId), len(StudentNames))

		print("yes",categoryListRegular)
		Categories_Name = ["none","Category_A", "Category_B", "Category_C", "Category_D", "Category_E", "Category_F"]
		FinalCategoryValue = []

		for i, j, id_s, name in zip(categoryListTimed, categoryListRegular, StudentId, StudentNames):
			value = min(i,j)
			FinalCategoryValue.append((id_s,name,value))

		FinalCategoryValue.sort(key=lambda x:x[2], reverse=True)

		print(FinalCategoryValue)
		FinalCategorization = []
		for id_s, name, value in FinalCategoryValue:
			FinalCategorization.append((id_s,name,Categories_Name[value]))

		print(FinalCategorization)

		worksheet.update_cell(2,1, "Student Id")
		worksheet.update_cell(2,3, "Student Name")
		worksheet.update_cell(2,5, "Category")



		cells = []
		for i in range(1, len(FinalCategorization)+1):
			cells.append(Cell(row=i+3, col=1, value=FinalCategorization[i-1][0]))
			cells.append(Cell(row=i+3, col=3, value=FinalCategorization[i-1][1]))
			cells.append(Cell(row=i+3, col=5, value=FinalCategorization[i-1][2]))

		worksheet.update_cells(cells)






		
		
# sheet.writeOnSheet(emails, names)


# Generate Categories Report
sheet1 = Spreadsheet(scope,creds, client)
sheet1.generateReport_1ForAllGroups()


# Generate Student-Categories Mapping Report
sheet2 = Spreadsheet(scope,creds, client)
sheet2.generateReport_3ForAllGroups()

