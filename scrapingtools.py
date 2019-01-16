import xlsxwriter
from uszipcode import SearchEngine
import xlrd

search = SearchEngine(simple_zipcode=True)

#find a substring between two other substrings; found on stack overflow
def find_between(s, first, last):
	try:
		start = s.index(first) + len(first)
		end = s.index(last, start)
		return s[start:end]
	except ValueError:
		return ''

#get todays date in format: mm_dd_yyyy
def getTodaysDate():
	today = datetime.datetime.today()
	month = today.month
	day = today.day
	year = today.year
	if today.month < 10:
		month = '0' + str(month)
	if today.day < 10:
		day = '0' + str(day)
	return str(month) + '_' + str(day) + '_' + str(year)

#return current time in format hh_mm
def getTime():
	current_time = datetime.datetime.today().time()
	hour = current_time.hour
	if hour < 10:
		hour = '0' + str(hour)
	minute = current_time.minute
	if minute < 10:
		minute = '0' + str(minute)
	return str(hour) + '_' + str(minute)

#given a filename with the extension CSV, reads the file and converts a comma seperated list into a list
def convertCSVToList(filename):
	if '.csv' not in str(filename):
		filename = str(filename) + '.csv'
	f = open(filename, mode='r', encoding='utf-8-sig')
	lines = f.readlines()
	f.close()
	data = []
	for line in lines:
		data_line = line.split(',')
		data.append(data_line)
	return data

#provide filename and a list of lists
def createCSVFromList(filename, csv_data):
	if '.csv' not in str(filename):
		filename = str(filename) + '.csv'
	f = open(filename, mode='w')
	for row in csv_data:
		row_value = ''
		for element in row:
			row_value += (str(element) + ',')
		f.write(row_value[:-1].strip() + '\n')
	f.close()

'''
convert .xlsx file to dictionary
filename is path to .xlsx file
dictionary format: {worksheet_name : [[column_data], [column_data_2], [column_data_k]]}
'''
def convertExcelToDictionary(filename):
	if '.xlsx' not in str(filename):
		filename = str(filename) + '.xlsx'
	wb = xlrd.open_workbook(filename)
	sheets = {}
	for i in range(0, wb.nsheets):
		active_sheet = wb.sheet_by_index(i)
		name = active_sheet.name
		sheets.update({name : []})
		num_of_rows = active_sheet.nrows
		for k in range(0, num_of_rows):
			values = []
			for element in active_sheet.row(k):
				values.append(str(element.value).strip())
			sheets[name].append(values)
	return sheets

'''
sheet_data should be a dictionary where:
each key is the name of a worksheet
each value paired to each key is a list of a lists
each inner-most list is a list of elements that should be added to that worksheet
ie: {sheet_name: [[1,2,3], [4,5,6], [7,8,9]]}
'''
def convertDictionaryToExcel(sheet_data, excel_name): #sheet_data = {sheet_name: [list, of, list, of, elements]}
	if '.xlsx' not in str(excel_name):
		excel_name = str(excel_name) + '.xlsx'
	final_workbook = xlsxwriter.Workbook(excel_name)
	sheet_name = excel_name.split('/')[-1].split('.')[0]
	for sheet_name in sheet_data:
		final_workbook.add_worksheet(sheet_name)
		row_num = 0 #Row 1
		column_num = 0 #column A
		for row in sheet_data[sheet_name]:
			for element in row:
				final_workbook.sheetnames[sheet_name].write_string(row_num, column_num, element)
				column_num += 1
			row_num += 1
			column_num = 0
	final_workbook.close()

def getLocationDataFromZip(zipcode):
	zipcode = int(zipcode)
	zipcode_search = search.by_zipcode(zipcode)
	lat = float((zipcode_search.bounds_north + zipcode_search.bounds_south)/2)
	lng = float((zipcode_search.bounds_west + zipcode_search.bounds_east)/2)
	coords = (lat, lng)
	city = zipcode_search.major_city
	county = zipcode_search.county
	state = zipcode_search.state
	try:
		state = us.states.lookup(state).name
	except:
		pass
	return {'zipcode': zipcode, 'coordinates': {'lat': coords[0], 'lng': coords[1]}, 'city': city, 'county': county, 'state': state}

def getLocationDataFromCoords(lat, lng):
	result = getLocationDataFromZip(search.by_coordinates(lat, lng)[0].zipcode)
	result['coordinates'] = {'lat': lat, 'lng': lng}
	return result

#client secret filename should be a json file provided by google
def loginToGoogle(client_secret_filename):
	scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
	creds = ServiceAccountCredentials.from_json_keyfile_name(client_secret_filename, scope)
	client = gspread.authorize(creds)
	return client

def getSheet(client, sheet_name):
	sheet = client.open(sheet_name)
	return sheet

def getExistingWorksheets(sheet):
	existing_worksheets = sheet.worksheets()
	return existing_worksheets

def getWorksheetData(worksheet):
	return worksheet.get_all_records()