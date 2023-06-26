#run path as of 4/4/23:   C:\Users\cara.fagerholm\source\repos\Costing\Costing\Costing.py
### Needed edits: 
### - for part 1, file identification, remove tests that started after the end of the month being evaluated (Improve to remove selection of files that were not started until after the month ended)
### - 

import os.path
from datetime import date
from datetime import datetime, timedelta
import csv

ArbinPath = "C:\\ArbinSoftware\\MITS_PRO\\Data"
#ArbinPath = "H:\\Tests\\Test Profiles\\Test 43 - Ten Nine\\Test 43ao - LTNT dryMix Validation -MT\\TestProcess"
FileList = os.listdir(ArbinPath)
InputList = []
print(InputList)

#find the first day of last month so we can check any files modified since then
today = date.today()
dt = today.replace(day=1)
#print(dt) #should print just the date (not the time) for the first of the month
dt = dt - timedelta(days=1)
LastMo = dt.replace(day=1)
print("The first day of last month was:", LastMo)

with open("C:\\Users\\Arbin\\Desktop\\Output_Files\\FilesToInput.csv", 'wt') as fi:
#with open("H:\\Tests\\Test Profiles\\Test 43 - Ten Nine\\Test 43ao - LTNT dryMix Validation -MT\\TestProcess\\FilesToInput.csv", 'wt') as fi:
	fi.write("the first day of last month was:")
	fi.write(str(LastMo))
	fi.write('\n')
#f in Filelist is just the file name
	print(FileList)
	for f in FileList:
		if f.startswith("23"):
			Mtime = os.path.getmtime(os.path.join(ArbinPath, f))
			MT = datetime.fromtimestamp(Mtime) #.strftime('%Y-%m-%d %H:%M:%S')
			ModTime = MT.date()
			if (ModTime >= LastMo) and (f.endswith('res')):
				#Copy filename from FileList into the InputList
				InputList.append(f)
				fi.write('\n')
				fi.write(f)
				print(InputList)



####Print the InputList into Input.xlsx
####InputXL = "C:\\csImport\\input.xlsx"
####OutputPath = "C:\Users\Arbin\Desktop\Output_Files"  #paste this into the 3rd column for every file to process
####recall ArbinPath from above to paste into 1st input col (for every file to process)

##############################################################################################################

#####import openpyxl
#####import pandas as pd

######Save Input.xlsx and run ResToExcel.exe
######ResToExcel.exe
######With all the imported files at the filepath from Input.xlsx, open and check the activity of each channel

######ImportedPath = "C:\\Users\\Arbin\\Desktop\\Output_Files" #Outputfolder on the desktop (can make sure this is input in the section above if I get to it)
#####ImportedPath = "H:\\Tests\\Test Profiles\\Test 43 - Ten Nine\\Test 43ao - LTNT dryMix Validation -MT\\TestProcess"
#####FileList = os.listdir(ImportedPath)

#####today = datetime.now()
#####EOMTime = today.replace(day=1, hour = 0, minute = 0, second = 0)
#####dt = EOMTime - timedelta(days=1)
#####LastMo = dt.replace(day=1)
#####print("The first day of last month was:", LastMo)
#####print('eom time' + str(EOMTime))

#####import datetime #can't go before datetime.now (may need new environment....)
#####for input in FileList:
######for input in InputList: 
#####	if input.endswith('xlsx'): #confirm if we want this (skip xls files)
######input = "220222-43ao-2_NMC811_LTNT_R1077-81-9-5-5.xlsx"

#################Indent all below
#####		InPath = os.path.join(ImportedPath, input)
#####		df = pd.ExcelFile(InPath) #, on_demand=True)
#####		ShtList = df.sheet_names
#####		#GlobalName = ShtList[0] #print(df.parse(sheetname=ShtList[0]))

#####		#In the global tab, check how many cells there are
#####		NumCells = len(df.parse(sheetname=ShtList[0]).index) - 3
#####		print(input)
#####		print("The Number of cells in this test:", NumCells)  #print(ShtList)
#####		SkipDaysLog = []
#####		CellDaysLog = []
#####		CellDaysSum = timedelta(days=0)
#####		StartTime = df.parse('Global_Info', header = 3).at[0, 'Start_DateTime'] ##########this gets intentionally overwritten if multi-month test!!!

#####		for Sht in ShtList: #Find each sheet that starts with "Stat" (for each cell in NumCells:)
#####			if Sht.startswith("Statistic"):
#####				#check all the active dates in this sheet in the third column     #print(len(df.parse(sheetname=Sht).index)) #len starts at 0
#####				# make a df out of the sheet we want to check
#####				Shtdf = df.parse(Sht)  #print(Shtdf.head())
#####				SlowCellLog = timedelta(days=0)
#####				for cycle in Shtdf.index: #cycle starting at 0 instead of 1
#####					#subtract each date from the prior, if timedelta >= 2days, then log #print(Shtdf['Date_Time'].loc[data.index[cycle]])	
#####					if cycle != 0:
#####						CurTime = Shtdf.at[cycle, 'Date_Time']

#####						#if data is in the reivewed time range
#####						if CurTime > LastMo and CurTime < EOMTime:
#####							PrevTime = Shtdf.at[(cycle-1), 'Date_Time']
#####							if CurTime > (PrevTime + timedelta(days=2)):
#####								LagLog = CurTime - PrevTime
#####								print('lag log:')
#####								print(LagLog)
#####								SlowCellLog = SlowCellLog + LagLog
#####								print(SlowCellLog)
#####								print('\n')
#####								print('\n')

#####				SkipDaysLog.append(SlowCellLog)
#####				#ShtStartTime = Shtdf.at[0, 'Date_Time']
#####				#print(len(Shtdf.index))
#####				if len(Shtdf.index) > 1:
					
#####					# if this data set is older than the month being reviewed, correct the start time for the sheet to the beginning of the month being reviewed
#####					if StartTime < LastMo:
#####						StartTime = LastMo

#####					# obtain the sheet end time
#####					ShtEndTime = Shtdf.at[len(Shtdf.index)-1, 'Date_Time']

#####					# Calculate the numenr of active cell days to add to the sum, depending on 
#####					#	if the sheet ended before the end of the calendar month being reviewed or not (continued cycling thru the end of the month)
#####					if ShtEndTime < EOMTime:
#####						CellDaysLog.append(ShtEndTime - StartTime)
#####						CellDaysSum = CellDaysSum + (ShtEndTime - StartTime)
#####					else:
#####						CellDaysLog.append(EOMTime - StartTime)
#####						CellDaysSum = CellDaysSum + (EOMTime - StartTime)

#####				# Add a minimum of one day of cycling if the cell was actively tested in the timeframe
#####				else:
#####					CellDaysLog.append(timedelta(days=1))
#####					CellDaysSum = CellDaysSum + timedelta(days=1)

#####		print("Skipdays: The number of days a cell did not log any cycles:", SkipDaysLog)
#####		print("Test days: The number of days a cell was on test:", CellDaysLog) #SumDays = sum(CellDaysLog, datetime.timedelta())
#####		print("Total test days, ignoring skip days:")
#####		print(CellDaysSum)

#####		# Initialize a variable to hold the total sum of timedelta objects
#####		total_skipDaysLog = datetime.timedelta()

#####		# Iterate through the list and add each timedelta object to the total sum
#####		for td in SkipDaysLog:
#####			if td.total_seconds() > 0:
#####				total_skipDaysLog += td

#####		#with open("C:\\Users\\Arbin\\Desktop\\Output_Files\\CellDaysOnTestLog.csv", 'a') as fcelldays:
#####		#with open("H:\\Tests\\Test Profiles\\Test 43 - Ten Nine\\Test 43ao - LTNT dryMix Validation -MT\\TestProcess\\CellDaysOnTestLog.csv", 'a') as fd:
#####		with open(os.path.join(ImportedPath, "CellDaysOnTestLog.csv"), 'a') as fd:
#####			fd.write('\n')
#####			fd.write(str(date.today()))
#####			fd.write('\n')
#####			fd.write(input)
#####			fd.write('\n')
#####			if total_skipDaysLog.total_seconds() > 0:
#####				fd.write('skipdayslog:, {}'.format(str(total_skipDaysLog.total_seconds() / 86400)))
#####			else:
#####				fd.write('skipdayslog: none')

#####			fd.write('\n')
#####			if CellDaysSum.total_seconds() >= 0:
#####				Total_CellDaysSum = CellDaysSum.total_seconds() / 86400
#####				fd.write('TotTestDays, ignoring skip days:, {}'.format(str(Total_CellDaysSum)))
#####			else:
#####				fd.write('TotTestDays - ignoring skip days:, none')
				
		

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

