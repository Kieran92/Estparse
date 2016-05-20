import os
def estparse():
	f =  open("1-Obj-1.txt", "r")
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


	print dateRecieved, insured, contact, claimNo, VIN, vehicleType, handlingAdjEmail

def main():
	print "Welcome to Wawanesa Estimate Parser"
	sheet = raw_input("Enter the name of the sheet that you wish to update: ")
	print sheet
	estimate = raw_input("Enter the name of the of the estimate: ")
	moreEsts = "yes"
	while True:
		estparse()
		moreEsts = raw_input("Do you wish to add another estimate to this sheet (yes/no)?: ")
		if moreEsts != "yes":
			break
		else:
			estimate = raw_input("Enter the name of the of the estimate: ")
			moreEsts = "no"


if __name__ == "__main__":
    main()
