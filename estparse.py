import os
import json
import os.path
import sys
import smartsheet
import requests
import re

def estparse(folder, fileName):
	base = os.path.dirname(__file__)
	filepath = os.path.abspath(os.path.join(base, "..","..", folder, fileName))
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
	deductable = ''
	city = ''
	payGST = False 
	flag = False # here so that I can get the insured city

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
		if "Deductible:  " in value:
			x = value.split("Deductible:  ")
			deductible = x[1]
		if "City:  " in value:
			x = value.split("City:  ")
			city = x[1]
		if "@" in value:
			y = value.split(' ')
			for word in y:
				if  "@" in word:
					handlingAdjEmail = word
		if "First Name:  " in value:
			x = value.split("First Name:  ")
			contact = x[1]
		if "GST/HST Registered?: " in value:
			x = value.split("GST/HST Registered?: " )
			if x[1] == "Yes":
				payGST = True 
		if "Phone 1:  " in value:
			x = value.split("Phone 1:  ")
			phone = x[1]
		if "Vehicle:                               Description:  " in value:
			x = value.split("Description:  ")
			vehicleType = x[1]


	return dateRecieved, insured, contact, claimNo, VIN, vehicleType, handlingAdjEmail, city, phone, deductible, payGST

def smartsheetupload(sheetName, folder):
	sheetDict = {}
	columnDict = {} 
	allRows = []
	base = os.path.dirname(__file__)
	filepath = os.path.abspath(os.path.join(base, "..","..", folder))

	sheet = smartsheet.Smartsheet('')
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
			print val.type
			columnDict[val.title] = val.id
		
		print columnDict
		for filename in os.listdir(filepath):
			dateRecieved, insured, contact, claimno, vin, vt, handlingadjemail, city, phone, deductible, payGST= estparse(folder, filename)
			row = smartsheet.models.Row()
			row.to_bottom = True 
			row.cells.append({
			    'column_id': columnDict.get("Claim Number"),
			    'value': claimno ,
			    'strict': False
			   })
			
			row.cells.append({
			    'column_id': columnDict.get("Insured") ,
			    'value': insured ,
			    'strict': False
			   })
			row.cells.append({
			    'column_id': columnDict.get("Contact") ,
			    'value': contact ,
			    'strict': False
			    })

			row.cells.append({
			    'column_id': columnDict.get("Assignment Received") ,
			    'value': dateRecieved,
			    'strict': False
			    })

			row.cells.append({
			    'column_id': columnDict.get("Contact Number") ,
			    'value': phone,
			    'strict': False
			    }) 

			row.cells.append({
			    'column_id': columnDict.get("City") ,
			    'value': city,
			    'strict': False
			    }) 

			row.cells.append({
			    'column_id': columnDict.get("VIN") ,
			    'value': vin,
			    'strict': False
			    }) 

			row.cells.append({
			    'column_id': columnDict.get("Year Make Model") ,
			    'value': vt,
			    'strict': False
			    }) 

			row.cells.append({
			    'column_id': columnDict.get("Deductible") ,
			    'value': deductible,
			    'strict': False
			    }) 
			
			row.cells.append({
			    'column_id': 7339369083234180,
			    'value': True
			   })
			row.cells.append({
			    'column_id': columnDict.get("Handling Adjuster Email") ,
			    'value': handlingadjemail,
			    'strict': False
			    }) 
			allRows.append(row)
		



		# Add rows to sheet.
		action = sheet.Sheets.add_rows(sheetID, allRows)

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
	folder = raw_input("Enter the name of the folder containing the dispatches: ")
	smartsheetupload(sheet, folder)


if __name__ == "__main__":
    main()
