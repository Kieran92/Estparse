import os
import json
import os.path
import sys
import smartsheet
import requests

def estparse():
	base = os.path.dirname(__file__)
	filepath = os.path.abspath(os.path.join(base, "..", "1-Obj-1.txt"))
	#base = os.getcwd()
	#os.chdir("../../")
	f =  open(filepath, "r")
	assignInfo = []#all of the information from the file
	dateRecieved = ''
	insured = ''
	contact = ''
	contactPhone = ''
	claimNo = ''
	email = ''
	VIN = ''
	vehicleType = ''
	handlingAdjEmail = ''

	for line in f:
		assignInfo.append((line.lstrip().rstrip()))
	x = assignInfo[0].split("Date: ")
	dateRecieved = x[1]

	for value in assignInfo:
		if "Estimate:                             Claim Number:  " in value:
			x = value.split("Claim Number:  ")
			claimNo = x[1]
		if "Insured:                                 Last Name:  " in value:
			x = value.split("Last Name:  ")
			#print x[1]
			insured =  x[1]
		if  " Description" in value:
			x = value.split("Description:  ")
			#vehicleType = x[1]
		if "VIN:  " in value:
			x = value.split("VIN:  ")
			VIN = x[1]
			#print VIN
		if "@" in value:
			handlingAdjEmail =  value
	contact = insured
	return dateRecieved, insured, contact, claimNo, VIN, vehicleType, handlingAdjEmail

def smartsheetupload(sheetName):
	sheetDict = {}
	columnDict = {} 
	sheet = smartsheet.Smartsheet('<access code>')
	me = sheet.Users.get_current_user()	
	#print me
	action = sheet.Sheets.list_sheets(include_all=True)
	sheets = action.data
	for val in sheets:
		sheetDict[val.name] = val.id
	#print sheetDict

	if sheetDict.get(sheetName):
		sheetID =  sheetDict.get(sheetName)
		#getting all of the columns
		action = sheet.Sheets.get_columns(sheetID, include_all=True)
		columns = action.data
		for val in columns:
			columnDict[val.title] = val.id
		
		print columnDict
		row_a = smartsheet.models.Row()
		row_a.to_top = True

		row_a.cells.append({
		    'column_id': 663502235953028,
		    'value': 'balla',
		    'strict': False
		})

		# Add rows to sheet.
		action = sheet.Sheets.add_rows(sheetID, [row_a])
	else:
		makeSheetQuery = raw_input("Create new sheet (y/n): ")
		if makeSheetQuery == "y":
			newSheetName = raw_input("Enter the name of the new sheet: ")
			#todo: create sheet and the columns for the sheet
		elif makeSheetQuery == "n":
			print "terminating"
			return None 

	#return None

def main():
	#print "Welcome to Wawanesa Estimate Parser"
	sheet = raw_input("Enter the name of the sheet that you wish to update: ")
	#print sheet
	#estimate = raw_input("Enter the name of the of the estimate: ")
	dateRecieved, insured, contact, claimno, vin, vt, handlingadjemail = estparse()
	smartsheetupload(sheet)


if __name__ == "__main__":
    main()
