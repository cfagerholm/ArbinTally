#run path as of 4/4/23:   C:\Users\cara.fagerholm\source\repos\Costing\Costing\Costing.py
### Needed edits: 
### - for part 1, file identification, remove tests that started after the end of the month being evaluated 
#    (Improve to remove selection of files that were not started until after the month ended)
### - 

import os.path
from datetime import date
from datetime import datetime, timedelta
import csv


##############################################################################################################

import openpyxl
import pandas as pd

#Save Input.xlsx and run ResToExcel.exe
#ResToExcel.exe
#With all the imported files at the filepath from Input.xlsx, open and check the activity of each channel

#ImportedPath = "C:\\Users\\Arbin\\Desktop\\Output_Files" #Outputfolder on the desktop (can make sure this is input in the section above if I get to it)
ImportedPath = "H:\\Tests\\Test Profiles\\Test 43 - Ten Nine\\Test 43ao - LTNT dryMix Validation -MT\\TestProcess"
FileList = os.listdir(ImportedPath)

#establish what month we are in at the time of evaluation and comparator values 
today = datetime.now()
EOMTime = today.replace(day=1, hour = 0, minute = 0, second = 0)
dt = EOMTime - timedelta(days=1)
LastMo = dt.replace(day=1)
print("The first day of last month was:", LastMo)
print('eom time' + str(EOMTime))

import datetime #this import line can't go before datetime.now (may need new environment....)
#loop through each file indicated by the user
for input in FileList:
#for input in InputList: 
	if input.endswith('xlsx'): #confirm if we want this (skip xls files)
#input = "220222-43ao-2_NMC811_LTNT_R1077-81-9-5-5.xlsx"

############Indent all below
		InPath = os.path.join(ImportedPath, input)
		df = pd.ExcelFile(InPath) #, on_demand=True)
		ShtList = df.sheet_names
		#GlobalName = ShtList[0] #print(df.parse(sheetname=ShtList[0]))

		#In the global tab, check how many cells there are
		NumCells = len(df.parse(sheetname=ShtList[0]).index) - 3
		print(input)
		print("The Number of cells in this test:", NumCells)  #print(ShtList)
		SkipDaysLog = []
		CellDaysLog = []
		CellDaysSum = timedelta(days=0)
		StartTime = df.parse('Global_Info', header = 3).at[0, 'Start_DateTime'] ##########this gets intentionally overwritten if multi-month test!!!

		#Find each sheet that starts with "Stat" (for each cell in NumCells:)
		for Sht in ShtList: 
			if Sht.startswith("Statistic"):
				#check all the active dates in this sheet in the third column     #print(len(df.parse(sheetname=Sht).index)) #len starts at 0
				# make a df out of the sheet we want to check
				Shtdf = df.parse(Sht)  #print(Shtdf.head())
				SlowCellLog = timedelta(days=0)
				for cycle in Shtdf.index: #cycle starting at 0 instead of 1
					#subtract each date from the prior, if timedelta >= 2days, then log #print(Shtdf['Date_Time'].loc[data.index[cycle]])	
					if cycle != 0:
						CurTime = Shtdf.at[cycle, 'Date_Time']

						#if data is in the reivewed time range
						if CurTime > LastMo and CurTime < EOMTime:
							PrevTime = Shtdf.at[(cycle-1), 'Date_Time']
							if CurTime > (PrevTime + timedelta(days=2)):
								LagLog = CurTime - PrevTime
								print('lag log:')
								print(LagLog)
								SlowCellLog = SlowCellLog + LagLog
								print(SlowCellLog)
								print('\n')
								print('\n')

				SkipDaysLog.append(SlowCellLog)
				#ShtStartTime = Shtdf.at[0, 'Date_Time']
				#print(len(Shtdf.index))
				if len(Shtdf.index) > 1:
					
					# if this data set is older than the month being reviewed, correct the start time for the sheet to the beginning of the month being reviewed
					if StartTime < LastMo:
						StartTime = LastMo

					# obtain the sheet end time
					ShtEndTime = Shtdf.at[len(Shtdf.index)-1, 'Date_Time']

					# Calculate the numenr of active cell days to add to the sum, depending on 
					#	if the sheet ended before the end of the calendar month being reviewed or not (continued cycling thru the end of the month)
					if ShtEndTime < EOMTime:
						CellDaysLog.append(ShtEndTime - StartTime)
						CellDaysSum = CellDaysSum + (ShtEndTime - StartTime)
					else:
						CellDaysLog.append(EOMTime - StartTime)
						CellDaysSum = CellDaysSum + (EOMTime - StartTime)

				# Add a minimum of one day of cycling if the cell was actively tested in the timeframe
				else:
					CellDaysLog.append(timedelta(days=1))
					CellDaysSum = CellDaysSum + timedelta(days=1)

		# Debug prints, can be removed
		print("Skipdays: The number of days a cell did not log any cycles:", SkipDaysLog)
		print("Test days: The number of days a cell was on test:", CellDaysLog) #SumDays = sum(CellDaysLog, datetime.timedelta())
		print("Total test days, ignoring skip days:")
		print(CellDaysSum)

		# Initialize a variable to hold the total sum of timedelta objects
		total_skipDaysLog = datetime.timedelta()

		# Iterate through the list and add each timedelta object to the total sum
		for td in SkipDaysLog:
			if td.total_seconds() > 0:
				total_skipDaysLog += td

		# Write to output file, 
		#	Include total days that a cell was on test, and
		#	Include "skipdays" (when there was no cycle collected for an extended period of time, which may be subtracted from tot cell days at discretion of user)
		#with open("C:\\Users\\Arbin\\Desktop\\Output_Files\\CellDaysOnTestLog.csv", 'a') as fcelldays:
		#with open("H:\\Tests\\Test Profiles\\Test 43 - Ten Nine\\Test 43ao - LTNT dryMix Validation -MT\\TestProcess\\CellDaysOnTestLog.csv", 'a') as fd:
		with open(os.path.join(ImportedPath, "CellDaysOnTestLog.csv"), 'a') as fd:
			fd.write('\n')
			fd.write(str(date.today()))
			fd.write('\n')
			fd.write(input)
			fd.write('\n')
			if total_skipDaysLog.total_seconds() > 0:
				fd.write('skipdayslog:, {}'.format(str(total_skipDaysLog.total_seconds() / 86400)))
			else:
				fd.write('skipdayslog: none')

			fd.write('\n')
			if CellDaysSum.total_seconds() >= 0:
				Total_CellDaysSum = CellDaysSum.total_seconds() / 86400
				fd.write('TotTestDays, ignoring skip days:, {}'.format(str(Total_CellDaysSum)))
			else:
				fd.write('TotTestDays - ignoring skip days:, none')
				
		

#Traceback (most recent call last):
#  File "Costing.py", line 48, in <module>
#  File "PyInstaller\loader\pyimod02_importers.py", line 499, in exec_module
#  File "openpyxl\__init__.py", line 6, in <module>
#  File "PyInstaller\loader\pyimod02_importers.py", line 499, in exec_module
#  File "openpyxl\workbook\__init__.py", line 4, in <module>
#  File "PyInstaller\loader\pyimod02_importers.py", line 499, in exec_module
#  File "openpyxl\workbook\workbook.py", line 9, in <module>
#  File "PyInstaller\loader\pyimod02_importers.py", line 499, in exec_module
#  File "openpyxl\worksheet\_write_only.py", line 13, in <module>
#  File "openpyxl\worksheet\_writer.py", line 23, in init openpyxl.worksheet._writer
#ModuleNotFoundError: No module named 'openpyxl.cell._writer'
#[19212] Failed to execute script 'Costing' due to unhandled exception!

