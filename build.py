import os
import sys
import csv
import calendar
import zipfile
import gzip
import json

ds = "/"
if (sys.platform == "win32"):
	ds = "\\"

def listFiles(dir, usefullname=True, recurse=False, appendStr = ""):
	thelist = []
	for name in os.listdir(dir):
		full_name = os.path.join(dir, name)

		if os.path.isdir(full_name) and recurse==True:
			thelist.extend(listFiles(full_name, usefullname, appendStr+name+ds))
		else:
			if usefullname==True:
				thelist.extend([appendStr + full_name])
			else:
				thelist.extend([appendStr + name])
	return thelist

def get_str_extension(str):
	findex = str.rfind('.')
	h_ext = str[findex+1:len(str)]
	return h_ext

# http://stackoverflow.com/a/12886818
def unzip(source_filename, dest_dir):
    with zipfile.ZipFile(source_filename) as zf:
        for member in zf.infolist():
            # Path traversal defense copied from
            # http://hg.python.org/cpython/file/tip/Lib/http/server.py#l789
            words = member.filename.split('/')
            path = dest_dir
            for word in words[:-1]:
                while True:
                    drive, word = os.path.splitdrive(word)
                    head, word = os.path.split(word)
                    if not drive:
                        break
                if word in (os.curdir, os.pardir, ''):
                    continue
                path = os.path.join(path, word)
            zf.extract(member, path)

def formatTuple3(param1, length1, param2, length2, param3, length3):
	return ("| " + param1).ljust(length1) + ("| " + param2).ljust(length2) + "|" + (str(param3) + " |").rjust(length3)

def findIOSMonthlySummary(csvfiles, directory, monthIndex, year):

	for file in csvfiles:

		#print '---' + file

		with open(directory + ds + file, 'rb') as csvfile:
			filereader = csv.reader(csvfile, delimiter=',', quotechar='"')
			for row in filereader:
				month = row[0][row[0].rfind("(")+1 : row[0].rfind(",")]
				csvyear = row[0][ row[0].rfind(" ")+1 : row[0].rfind(")")]
				#print calendar.month_name[int(monthIndex)] + ' against ' + month
				#print csvyear + ' against ' + str(year)
				if calendar.month_name[int(monthIndex)] == month and int(csvyear) == int(year):
					correctfile = file
					return correctfile

	print "Could not find monthly summary file for month " + str(int(monthIndex)) + " and year " + str(year)
	exit(0)

def doesBalanceColumnExistInSummary(summaryfile):
	f = open(summaryfile, "r")
	contents = f.read()
	f.close()
	return contents.find("Balance") >= 0

def getBalanceForCurrency(summaryfile, currency):
	balanceIndex = 2
	with open(summaryfile, 'rb') as csvfile:
		filereader = csv.reader(csvfile, delimiter=',', quotechar='"')
		i = 0
		for row in filereader:
			if i > 2 and row[0] is not '':

				if len(row) < 11:
					print 'Balance carried forward ' + currency + " - " +(summaryfile)
					return 0

				region = row[0]
				rowcurrency = region[len(region)-4:len(region)-1]
				if currency == rowcurrency:
					value = row[balanceIndex]
					if len(value) == 0:
						value = 0
					return value
			i += 1

	print "Could not find balance currency " + currency + " in file " + summaryfile
	exit(0)

def findIOSExchangeRateForCurrency(summaryfile, currency):
	exchangeRateIndex = 8
	if (doesBalanceColumnExistInSummary(summaryfile)):
		exchangeRateIndex += 1

	with open(summaryfile, 'rb') as csvfile:
		filereader = csv.reader(csvfile, delimiter=',', quotechar='"')
		i = 0
		for row in filereader:
			if i > 2 and row[0] is not '':

				if len(row) < 11:
					print 'Balance carried forward ' + currency + " - " +(summaryfile)
					return [ True, 1.0 ]

				region = row[0]
				rowcurrency = region[len(region)-4:len(region)-1]
				if currency == rowcurrency:
					return [ False, row[exchangeRateIndex] ]
			i += 1

	print "Could not find exchange rate for currency " + currency + " in file " + summaryfile
	exit(0)

def findIOSWithholdingTaxForCurrency(summaryfile, currency):
	withholdingTaxIndex = 6
	if (doesBalanceColumnExistInSummary(summaryfile)):
		withholdingTaxIndex  += 1

	with open(summaryfile, 'rb') as csvfile:
		filereader = csv.reader(csvfile, delimiter=',', quotechar='"')
		i = 0
		for row in filereader:
			if i > 2 and row[0] is not '':
				if len(row) < 11:
					#print 'Withholding tax carried forward ' + currency
					return 0

				region = row[0]
				rowcurrency = region[len(region)-4:len(region)-1]
				if currency == rowcurrency:
					return float(row[withholdingTaxIndex])
			i += 1

	print "Could not find withholding tax for currency " + currency + " in file " + summaryfile
	exit(0)

markdownTable = ""
mytable = ""

def monthForApps(a):
	global markdownTable
	global mytable

	monthtotal = 0
	for app in apps:
		monthtotal += apps[app]
		markdownTable += "| " + year + " | " + app + " | " + str(apps[app]) + " |" + nl
		mytable += formatTuple3( calendar.month_name[int(month)][0:3] + " " + year + " ", 12, app, 36, round(apps[app], 2), 12)  + nl

	markdownTable += "| " + year + " | total | " + str(monthtotal) + " |" + nl
	mytable += formatTuple3( calendar.month_name[int(month)][0:3] + " " + year + " ", 12, "total", 36, round(monthtotal, 2), 12)  + nl

	markdownTable += "---" + nl
	mytable += "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -" + nl

if __name__ == "__main__":

	print "======================="
	print "Generating Sales Report"
	print "* Author: Ashley Gwinnell / Force Of Habit"
	print "* Version: 0.1.0"

	# Default args
	directory = os.getcwd()
	platform = "none"
	sku_map = {}

	# Overwrite default args with parameters
	count = 0
	for item in sys.argv:
		if count == 0:
			count += 1
			continue
		parts = item.split("=")
		if len(parts) > 2:
			print("paramter " + item + " is invalid")
			continue
		if parts[0] == "directory":
			directory = parts[1]
			if (directory[len(directory)-1:len(directory)] == "/"):
				directory = directory[0:len(directory)-1]
		elif parts[0] == "platform":
			platform = parts[1]
		elif parts[0] == "sku_map":
			# e.g. python build.py platform=ios sku_map='{"tipjar.small":"makenines","tipjar.medium":"makenines","tipjar.large":"makenines"}' directory=~/AppStore/
			sku_map = json.loads(parts[1])
		elif parts[0] == "sku_map_file":
			f = open(parts[1])
			fcontents = f.read()
			f.close()
			sku_map = json.loads(fcontents)
		count += 1

	if (platform == "none"):
		print "WARN: Parameter 'platform' is not set. Pass it on the command line e.g. platform=googleplay"
		exit(0)

	print "======================="
	print "directory: " + directory
	print "platform: " + platform

	ignoreFiles = set(['.DS_Store', 'build.py'])
	files = filter(lambda x: x not in ignoreFiles, listFiles(directory, False))

	if (platform == "android" or platform == "googleplay"):
		nl = "\n"
		markdownTable += "| MONTH 	| APP 	| INCOME	|" + nl
		markdownTable += "|---		|---	|---:		|" + nl

		mytable += "-------------------------------------------------------------" + nl
		mytable += "| GOOGLE PLAY SALES                                         |" + nl
		mytable += "| MONTH     | APP                               | INCOME    |" + nl
		mytable += "-------------------------------------------------------------" + nl

		packageNameIndex = 7	# H column
		amountIndex = 18		# S column

		for monthFile in files:

			with open(directory + ds + monthFile, 'rb') as csvfile:

				apps = {}

				month = monthFile[13:15]
				year = monthFile[9:13]

				filereader = csv.reader(csvfile, delimiter=',', quotechar='"')
				i = 0
				for row in filereader:
					if ( i > 0 ):
						if (row[packageNameIndex] not in apps):
							apps[row[packageNameIndex]] = 0
						apps[row[packageNameIndex]] += float(row[amountIndex])
					i += 1

				monthForApps(apps)


	elif (platform == "ios" or platform == "appstore"):

		zipfiles = filter(lambda x: get_str_extension(x) == "zip", files)
		csvfiles = filter(lambda x: get_str_extension(x) == "csv", files)

		nl = "\n"
		markdownTable += "| MONTH 	| APP 	| INCOME	|" + nl
		markdownTable += "|---		|---	|---:		|" + nl

		mytable += "-------------------------------------------------------------" + nl
		mytable += "| IOS APPSTORE SALES                                        |" + nl
		mytable += "| MONTH     | APP                               | INCOME    |" + nl
		mytable += "-------------------------------------------------------------" + nl

		vendorIdentifierIndex = 4
		quantityIndex = 5
		amountIndex = 7
		currencyIndex = 8
		saleOrReturnIndex = 9

		balanceCarriedForward = {} #balanceCarriedForward['app']['currency']

		#
		# for each month
		#
		for file in zipfiles:

			apps = {}
			numApps = 0

			#print file
			outfolder = directory + ds + file[0:len(file)-4]
			#print outfolder
			if (not os.path.isdir(outfolder)):
				os.makedirs(outfolder)
				unzip(directory + ds + file, outfolder)

			month = 0
			year = 0

			#
			# for each region
			#
			regionfiles = filter(lambda x: get_str_extension(x) == "gz", listFiles(outfolder))
			for regionfile in regionfiles:
				#print regionfile
				gzfile = gzip.open(regionfile, 'rb')
				tsv = gzfile.read()
				gzfile.close()

				tsvname = regionfile[0:len(regionfile)-3]
				f = open(tsvname, "w")
				f.write(tsv)
				f.close()

				month = tsvname[tsvname.find("_")+1:tsvname.find("_")+3]
				year = tsvname[tsvname.find("_")+3:tsvname.find("_")+5]

				summaryfile = directory + ds + findIOSMonthlySummary(csvfiles, directory, month, 2000 + int(year))
				#print 'month ' + month
				#print 'month ' + summaryfile

				balanceExists = doesBalanceColumnExistInSummary(summaryfile)

				with open(tsvname, 'rb') as tsvfile:
					i = 0
					#print regionfile
					regioncurrency = ''
					filereader = csv.reader(tsvfile, delimiter='\t')
					for row in filereader:
						if i > 0 and len(row) > 2:
							#print row#[vendorIdentifierIndex]

							skuname = sku_map[ row[vendorIdentifierIndex] ] if row[vendorIdentifierIndex] in sku_map else row[vendorIdentifierIndex]
							if (skuname not in apps):
								apps[skuname] = 0
								numApps += 1

							amountInBuyersCurrency = float(row[amountIndex])

							#print row[saleOrReturnIndex]

							regioncurrency = row[currencyIndex]
							exchangeratedata = findIOSExchangeRateForCurrency(summaryfile, regioncurrency)
							if (exchangeratedata[0] == True):
								# balance carried forward
								#print 'balance ' + str(amountInBuyersCurrency) + " " + regioncurrency + ' carried forward for app ' + skuname
								#print balanceCarriedForward
								if skuname not in balanceCarriedForward:
									balanceCarriedForward[skuname] = {}
								if str(regioncurrency) not in balanceCarriedForward[skuname]:
									balanceCarriedForward[skuname][regioncurrency] = 0.0
								balanceCarriedForward[skuname][regioncurrency] += amountInBuyersCurrency
							else:
								exchangerate = float(exchangeratedata[1])
								apps[skuname] += amountInBuyersCurrency * exchangerate

						i += 1

					if regioncurrency is not '':
						# minus any witholding taxes
						withholdingAmount = findIOSWithholdingTaxForCurrency(summaryfile, regioncurrency)
						#print str(month) + " " + str(year) + " - " + regioncurrency + " : " + str(withholdingAmount)
						if (withholdingAmount < 0):
							print "--------"
							print "Withholding " + str(withholdingAmount) + " in " + regioncurrency
							print "Splitting evenly amongst " + str(numApps) + " apps..."
							print "--------"
							exchangerate = float(findIOSExchangeRateForCurrency(summaryfile, regioncurrency)[1])
							for app in apps:
								apps[app] += (withholdingAmount / numApps) * exchangerate


			# check carried forward balance
			if doesBalanceColumnExistInSummary(summaryfile) and balanceCarriedForward:
				for appid in balanceCarriedForward:
					#print appid
					#print balanceCarriedForward[appid]
					for balcurrency in balanceCarriedForward[appid]:
						exchangeratedata = findIOSExchangeRateForCurrency(summaryfile, balcurrency)
						#print exchangeratedata
						if not exchangeratedata[0]:
							amountInBuyersCurrency = balanceCarriedForward[appid][balcurrency]
							skuname = sku_map[ appid ] if appid in sku_map else appid
							#print ('----')
							#print skuname
							#print apps
							apps[skuname] += amountInBuyersCurrency * float(exchangeratedata[1])
				balanceCarriedForward = {}
				#print balanceCarriedForward

			monthForApps(apps)

	pass

	print "======================="
	print mytable

	# Write raw format
	f = open(os.getcwd() + ds + "report-" + platform + ".txt", "w")
	f.write(mytable)
	f.close()

	# Write markdown format
	f = open(os.getcwd() + ds + "report-" + platform + ".md", "w")
	f.write(markdownTable)
	f.close()

	print "======================="
	print "Done!"
	print "======================="

