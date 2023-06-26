import openpyxl
import xlrd
import os

#Save Input.xlsx and run ResToExcel.exe
#ResToExcel.exe

#With all the imported files at the filepath from Input.xlsx, open and check the activity of each channel

#ImportedPath = "C:\\Users\\Arbin\\Desktop\\Output_Files" #Outputfolder on the desktop (can make sure this is input in the section above if I get to it)
#for input in InputList:

ImportedPath = "H:\\Tests\\Test Profiles\\Test 43 - Ten Nine\\Test 43ao - LTNT dryMix Validation -MT\\TestProcess"
input = "220222-43ao-2_NMC811_LTNT_R1077-81-9-5-5"

############Indent all below
InPath = os.path.join(ImportedPath, input)

InFile = xlrd.open_workbook(InPath, on_demand=True)
ShtList = InFile.sheet_names()
dfGlobal = InFile.sheet_by_index(0)

#In the global tab, check how many cells there are
NumCells = dfGlobal.nrows
print(ShtList)

	#for Sht in ShtList: #for cell in NumCells:
	#	#Find next sheet that starts with "Channel"
	#	if Sht.value.startswith("Channel"):
	#		#check all the active dates in this sheet in the third column
	#		print(Sht.nrows)
	#		for cycle in  Sht.nrows
	#		#subtract each date from the prior, if timedelta >= 2days, then log
	#			if (Sht[].value isvalue) and (Sht[ +1].value isvalue):
	#				if Sht[+1].value - Sht[].value > 2
	#					SlowCycleLog = SlowCycleLog.append(Sht[+1].value - Sht[].value)
		
