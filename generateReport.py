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




class Spreadsheet:

    def __init__(self,scope, creds, client):
        self.scope = scope
        self.creds = creds
        self.client = client


    def writeOnSheet(self, emails, names):
        spreadsheet = self.client.open("test")
        worksheet =  spreadsheet.add_worksheet("Emails", len(emails), 2)

        for i in range(1, len(emails)+1):
            worksheet.update_cell(i, 1, names[i-1])
            worksheet.update_cell(i, 2, emails[i-1])

        print("Writing Done")


    def readFromSheet(self):
        spreadsheet = self.client.open("test").worksheet("Emails")
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
        spreadsheet = self.client.open("Samurai_Sprint2_05th July 2021- Last 7 Days").worksheet(groupFileName)

        # Spreadsheet for writing
        spreadsheetWriting = self.client.open("Samurai_cohort-13_Sprint2_Report")
        # group = "Alpha_Report1"
        worksheet =  spreadsheetWriting.add_worksheet(group,1000, 1000)

        
        # Get Percentage of Regular Contest
        Regular_Accepted = spreadsheet.col_values(2)[1:]
        max_Attempted = int(max(spreadsheet.col_values(5)[1:]))
        percentageOfAttempts = [math.ceil((int(x)/max_Attempted)*100) for x in Regular_Accepted]


        # Arrange in Categories
        percentageOfAttempts.sort()
        print("Regular ", percentageOfAttempts)

        # Regular Contest Done
        self.prepareCategoryReport(percentageOfAttempts, "Regular Contest Report", 2, worksheet)
        print("Regular Contest Report : Done")






        # Get Percentage of Timed Contest
        Timed_Accepted = spreadsheet.col_values(6)[1:]
        max_Attempted = int(max(spreadsheet.col_values(9)[1:]))
        percentageOfAttempts = [math.ceil((int(x)/max_Attempted)*100) for x in Timed_Accepted]


        # Arrange in Categories
        percentageOfAttempts.sort()
        print("Timed ",percentageOfAttempts)

        self.prepareCategoryReport(percentageOfAttempts, "Timed Contest Report", 14, worksheet)


    def generateReport_1ForAllGroups(self):
    	pass
        
        # self.generateReportForGroups("Alpha_Report1", "Cohort-13-Unit-5_alpha" )
        # self.generateReportForGroups("Bravo_Report1", "Cohort-13-Unit-5_bravo" )






    def generateReport_3ForAllGroups(self):
        # self.generateReport_3ForGroups("Alpha_Report3", "Cohort-13-Unit-5_alpha")
        self.generateReport_3ForGroups("Bravo_Report3", "Cohort-13-Unit-5_bravo")
        


        
    
    def generateReport_3ForGroups(self, group, groupFileName):
        spreadsheet = self.client.open("Samurai_Sprint2_05th July 2021- Last 7 Days").worksheet(groupFileName)

        # Spreadsheet for writing
        spreadsheetWriting = self.client.open("Samurai_cohort-13_Sprint2_Report")
        # group = "Alpha_Report1"
        worksheet =  spreadsheetWriting.add_worksheet(group,1000, 1000)

        
        # Get Percentage of Regular Contest
        Regular_Accepted = spreadsheet.col_values(2)[1:]
        max_Attempted = int(max(spreadsheet.col_values(5)[1:]))
        percentageOfAttempts = [math.ceil((int(x)/max_Attempted)*100) for x in Regular_Accepted]
        categoryListRegular = []
        for i in percentageOfAttempts:
            if i == 100: categoryListRegular.append(1)
            if i<100 and i>=75: categoryListRegular.append(2)
            if i<75 and i>=50: categoryListRegular.append(3)
            if i>=25 and i<50: categoryListRegular.append(4)
            if i>=10 and i<25: categoryListRegular.append(5)
            if i<10: categoryListRegular.append(6)

        Timed_Accepted = spreadsheet.col_values(6)[1:]
        max_Attempted = int(max(spreadsheet.col_values(9)[1:]))
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






        
        




sheet = Spreadsheet(scope,creds, client)
# sheet.writeOnSheet(emails, names)

# sheet.generateReport_1ForAllGroups()

sheet.generateReport_3ForAllGroups()


