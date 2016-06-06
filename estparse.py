import os
import os.path
import sys
import smartsheet
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
	#print assignInfo
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

def smartsheetupload(dateRecieved, insured, contact, claimNo, VIN, vehicleType, handlingAdjEmail):
	sheet = smartsheet.Smartsheet('1ep68bn2j4frngtyuumt0x3896')
	me = sheet.Users.get_current_user()	
	print me
	#return None

def main():
	#print "Welcome to Wawanesa Estimate Parser"
	#sheet = raw_input("Enter the name of the sheet that you wish to update: ")
	#print sheet
	#estimate = raw_input("Enter the name of the of the estimate: ")
	dateRecieved, insured, contact, claimno, vin, vt, handlingadjemail = estparse()
	smartsheetupload(dateRecieved, insured, contact, claimno, vin, vt, handlingadjemail)


if __name__ == "__main__":
    main()
