# Copyright 2016 Kieran Boyle
# 
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
# 
#     http://www.apache.org/licenses/LICENSE-2.0
# 
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

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
		if "@w" in value:
			y = re.split(',|0|1|2|3|4|5|6|7|8|9|ext| ', value)
			for word in y:
				if  "@w" in word:
					handlingAdjEmail = word
		if "First Name:  " in value:
			x = value.split("First Name:  ")
			contact = x[1]
		if "GST/HST Registered?: " in value:
			x = value.split("GST/HST Registered?: " )
			payGST = x[1] 
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
	uploadedIds= []
	rowIDs = []
	cells = []
	allCells = []
	base = os.path.dirname(__file__)
	filepath = os.path.abspath(os.path.join(base, "..","..", folder))

	sheet = smartsheet.Smartsheet('')
	me = sheet.Users.get_current_user()	
	action = sheet.Sheets.list_sheets(include_all=True)
	sheets = action.data
	for val in sheets:
		sheetDict[val.name] = val.id


	if sheetDict.get(sheetName):
		sheetID =  sheetDict.get(sheetName)
		wholeSheet = sheet.Sheets.get_sheet(sheetID)
		rowList = list(wholeSheet.rows)
		for row in rowList:
			rowIDs.append(row.id)
		print "Checking for duplicates"
		for rowID in rowIDs:
			rowContent = sheet.Sheets.get_row(sheetID, rowID)
			allCells.append(list(rowContent.cells))
		for cellList in allCells:
			uploadedIds.append(cellList[0].value)

		#print uploadedIds

		action = sheet.Sheets.get_columns(sheetID, include_all=True)
		columns = action.data
		for val in columns:
			columnDict[val.title] = val.id

		print "Uploading Entries"
		for filename in os.listdir(filepath):
			dateRecieved, insured, contact, claimno, vin, vt, handlingadjemail, city, phone, deductible, payGST= estparse(folder, filename)
			if claimno not in uploadedIds:
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
				    'column_id': columnDict.get("Customer Pays GST"),
				    'value': payGST
				   })
				row.cells.append({
				    'column_id': columnDict.get("Handling Adjuster Email") ,
				    'value': handlingadjemail,
				    'strict': False
				    }) 
				allRows.append(row)
		# Add rows to sheet.
		action = sheet.Sheets.add_rows(sheetID, allRows)
		return None
	else:
		makeSheetQuery = raw_input("Retry (y/n): ")
		if makeSheetQuery == "y":
			#semi recursive call to retry the sheet upload and 
			retrySheet = raw_input("Enter the name of the sheet that you wish to update: ")
			retryFolder = raw_input("Enter the name of the folder containing the dispatches: ")
			smartsheetupload(retrySheet, retryFolder)
		elif makeSheetQuery == "n":
			print "terminating"
			return None 

def main():
	#print "Welcome to Wawanesa Estimate Parser"
	sheet = raw_input("Enter the name of the sheet that you wish to update: ")
	#print sheet
	folder = raw_input("Enter the name of the folder containing the dispatches: ")
	smartsheetupload(sheet, folder)


if __name__ == "__main__":
    main()
