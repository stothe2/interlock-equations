from xlrd import open_workbook
from xlsxwriter import Workbook
from pyeda.inter import *

from Tree import *

# Workbook created by __.
# Might need update <--
wb = open_workbook(filename='227-174783-002_C.xls', formatting_info=True)
matrixSheet = wb.sheet_by_name('Interlock Matrix')
stringsSheet = wb.sheet_by_name('Interlock Strings')
rimSheet = wb.sheet_by_name('IO Mapping Table')
# Array containing positions (int) of all rows with 'ILK'
interlockNumArray = []
# Array containing positions (int) of all rows with 'Where Used'
whereUsedNumArray = []
# Array containing starting row number (int) of rows containing I/O members
ioMemNumArray = []

# Workbook created by Lam Controls group.
# Might need update <--
controlsWb = open_workbook(filename='Signal List_13.xlsx')

# Might need update <--
def get_mapped_name(descriptionName, rimNumber):
	if rimNumber == '1':
		s = controlsWb.sheet_by_name('RIM1 ^ -001')
	elif rimNumber == '2':
		s = controlsWb.sheet_by_name('RIM2 ^ -001')
	elif rimNumber == '3':
		s = controlsWb.sheet_by_name('RIM3 ^ -001')

	if not s:
		raise Exception('%s not a valid RIM#' % rimNumber)

	nameCol = s.col_values(0)
	hardwareNameCol = s.col_values(1)

	for index, hardwareName in enumerate(hardwareNameCol):
		if hardwareName.lower() == descriptionName.lower():
			return nameCol[index]
	return None

def replace_signal_name_doi(name, rimNumber):
	if name[0] == 'M':
		name = name[2:]
	if name[0] == 'R':
		name = name[3:]

	descriptionColumn = rimSheet.col_values(4)
	doColumn = rimSheet.col_values(16)

	for index, signal in enumerate(doColumn):
		if name.find(signal) != -1 and signal != '':
			mappedName = get_mapped_name(descriptionColumn[index], rimNumber)
			# Might need update <--
			if name == 'DI_239':
				return 'DI100'
			if name == 'DO_0206':
				return 'DO36'
			if name == 'DI_209':
				return 'DI100'
			if name == 'DI_211':
				return 'DI101'
			if name == 'DI_213':
				return 'DI102'
			if name == 'DI_215':
				return 'DI103'
			if name == 'DI_217':
				return 'DI104'
			if not mappedName:
				raise Exception('%s does not map to Controls Group names' % name)
			return mappedName

def replace_signal_name_so(name, rimNumber):
	if name[0] == 'M':
		name = name[2:]
	if name[0] == 'R':
		name = name[3:]

	descriptionColumn = rimSheet.col_values(4)
	signalColumn = rimSheet.col_values(0)

	for index, signal in enumerate(signalColumn):
		if name.find(signal) != -1:
			mappedName = get_mapped_name(descriptionColumn[index], rimNumber)
			# Might need update <--
			if name == 'SO_128':
				return name
			if name == 'SO_130':
				return name
			if name == 'SO_137':
				return name
			if name == 'SO_138':
				return name
			if name == 'SO_146':
				return 'DO81'
			if name == 'SO_152':
				return name
			if name == 'SO_153':
				return name
			if name == 'SO_159':
				return name
			if name == 'SO_160':
				return name
			if name == 'SO_161':
				return name
			if name == 'SO_162':
				return name
			if name == 'SO_163':
				return name
			if name == 'SO_201':
				return name
			if name == 'SO_209':
				return name
			## this is due to a space in Spare_HEATER _SSR_17!
			if name == 'SO_218':
				return 'DO108'
			## this is due to AMPDS_vaporizer_inlet_valve)open_cmd!
			if name == 'SO_268':
				return 'DO36'
			if not mappedName:
				raise Exception('%s does not map to Controls Group names' % name)
			return mappedName

def get_rim_number(name, soName):
	signalColumn = rimSheet.col_values(0)
	doColumn = rimSheet.col_values(16)
	rimColumn = rimSheet.col_values(32)

	for index, item in enumerate(signalColumn):
		if name.find(item) != -1:
			if rimColumn[index].find('RIM') != -1:
				# The RIM Column is currently formatted like "RIM _".
				# We only want the the number, not the word "RIM" too.
				# Might need update <--
				temp = rimColumn[index].split()
				return str(temp[1])

	for index, item in enumerate(doColumn):
		if name.find(item) != -1:
			if rimColumn[index].find('RIM') != -1:
				# See above.
				temp = rimColumn[index].split()
				return str(temp[1])

def write_data(sheet, soArray, dependencyMap):
	# Keeps track of what row you're writing data to.
	rownum = 1
	for item in soArray:
		data = dependencyMap[item]
		ilk = data[0]
		equation = data[1]
		# Obtain RIM# for SO.
		rimNumber = get_rim_number(item, item)
		# Replace SO name with the name we use.
		mappedItem = replace_signal_name_so(item, rimNumber)
		# Ta-da!
		newItem = 'R' + rimNumber + '_' + mappedItem
		sheet.write(rownum, 0, newItem)
		sheet.write(rownum, 1, ilk)
		sheet.write(rownum, 2, equation)
		rownum = rownum + 1

def loop(tree, pointer, ilk):
		array = get_io_members(ilk)
		insert(array, tree, pointer)

def insert(array, tree, pointer):
	for index, item in enumerate(array):
		# If last I/O member for a given ILK, set "close, )".
		if index == len(array) - 1:
			tree.set_extra_close(pointer)
		pointer = tree.insert_child(pointer, item)

def match_helper(col, soName):
	for item in col:
		if item.find(soName) != -1:
			return True
	return False

# Found soName in given search columms?
def match(col1, col2, soName):
	found1 = match_helper(col1, soName)
	found2 = match_helper(col2, soName)
	if found1 or found2:
		return True
	return False

# Helper method for get_io_members.
# Returns row number where given ILK is present.
# If ILK not found, returns None.
def helper(name):
	for num in interlockNumArray:
		if name == stringsSheet.cell(num, 11).value:
			return int(interlockNumArray.index(num))
	return None

def get_io_members(name):
	temp = []
	# Finds interlock on worksheet.
	index = helper(name)
	if index == None:
		raise Exception('InterlockDoesNotExistException')
	# Generates list of I/O members of the given interlock.
	start = ioMemNumArray[index] + 1
	stop = whereUsedNumArray[index] - 1
	c1 = stringsSheet.col_values(11, start, stop)
	c2 = stringsSheet.col_values(26, start, stop)
	for item in c1:
		if item == '':
			continue
		temp.append(item)
	for item in c2:
		if item == '':
			continue
		temp.append(item)
	return temp

def generate_soArray():
	soArray = list()
	# Array containing entire column 'L' from Interlock Strings sheet.
	searchCol1 = stringsSheet.col_values(11)
	# Array containing entire column 'AA' from Interlock Strings sheet.
	searchCol2 = stringsSheet.col_values(26) 
	for s in searchCol1:
		if s.find('SO_') != -1:
			# Appends only the signal number, not the descriptive name.
			soArray.append(s[:6])
	for s in searchCol2:
		if s.find('SO_') != -1:
			# Appends only the signal number, not the descriptive name.
			soArray.append(s[:6])
	soArray.sort()
	return soArray

def main():
	# Array with all SO signal names, generated from Interlock Strings sheet.
	soArray = generate_soArray()
	# Array with all DO/DI signal names, generated from the Interlock Matrix sheet.
	doiArray = matrixSheet.row_values(2, 14, 146)

	# Array containing entire column 'A' from Interlock Strings sheet.
	# The array is used to determine row numbers corresponding to 'ILK',
	# 'Where Used', and 'I/O memebers' data.
	filterArray = stringsSheet.col_values(0)

	# Might need update <--
	for index, item in enumerate(filterArray):
		if item == "Interlock Number:":
			interlockNumArray.append(index)
		elif item == "Where Used:":
			whereUsedNumArray.append(index)
		elif item == "I/O Members:":
			ioMemNumArray.append(index)

	# Dictionary mapping signals to their equations
	dependencyMap = dict()

	for soName in soArray:
		dependencyMap[soName] = []

		# Looks for soName occurence in the Interlock Strings sheet.
		# After the occurence is detemined, the position is used to match
		# the signal with its main ILK.
		for index, item in enumerate(whereUsedNumArray):
			if index == (len(whereUsedNumArray) - 1):
				searchCol1 = stringsSheet.col_values(11, item+1)
				searchCol2 = stringsSheet.col_values(26, item+1)
			else:
				searchCol1 = stringsSheet.col_values(11, item+1, interlockNumArray[index+1])
				searchCol2 = stringsSheet.col_values(26, item+1, interlockNumArray[index+1])

			found = match(searchCol1, searchCol2, soName)
			if not found:
				continue
			break

		# TODO: can raise exception if soName not found in the sheet.

		# Appends the main ILK at the front of the array.
		dependencyMap[soName].append(stringsSheet.cell(interlockNumArray[index], 11).value)
		# Beginning row number for I/O members of the given ILK.
		start = ioMemNumArray[index] + 1
		# End row number for I/O members of the given ILK.
		stop = whereUsedNumArray[index] - 1
		# Column 'L' of the I/O members.
		c1 = stringsSheet.col_values(11, start, stop)
		# Column 'AA' of the I/O members.
		c2 = stringsSheet.col_values(26, start, stop)
		# Appends DO/DI signals from the first column.
		for doName in c1:
			if doName == '':
				continue
			dependencyMap[soName].append(doName)
		# Appends DO/DI signals from the second column.
		for doName in c2:
			if doName == '':
				continue
			dependencyMap[soName].append(doName)

	# For each SO signal, we flatten the equation, starting from the top,
	# i.e., the main ILK name, and its corresponding I/O members.
	for soName in soArray:
		array = dependencyMap[soName]
		# main ILK
		ilk = array[0]
		mainIlk = ilk
		# Array of DO/DI signals.
		array = array[1:]

		# Separate check for ILK_01, since it only contains one I/O member.
		# Might need update <--
		if mainIlk == 'ILK_01':
			dependencyMap[soName] = [mainIlk, 'DI_000']
			continue

		tree = Tree()
		# Main ILK inserted at root, and p is the pointer to the root.
		p = tree.insert_root(ilk)

		# Checks if any ILKs are present in the tree which we haven't
		# accounted for (not "visited" even once), and flattens the
		# equation by finding that ILK's I/O members, and so on...
		while tree.has_ilk():
			pointer, ilk = tree.get_ilk()
			loop(tree, pointer, ilk)

		# Ensures proper traversal order.
		# New ILK branches are always on the left, and the continuing I/O
		# members are on the right.
		tree.reverse_children()

		# New array with all signals for a given soName, with these
		# signals formatted properly (i.e., NOT changed to ~, ILKs skipped,
		# and parenthesis in place for logic reduction).
		arrayCopy = tree.format_list()

		# RIM Number generation.
		for index, item in enumerate(arrayCopy):
			# Temporary hack for DO_0265.
			# Might need update <--
			if item.find('DO_0265') != -1:
				if item.find('~') != -1:
					temp = item.split('~')
					t = temp[0] + 'M_' + temp[1]
				else:
					t = 'M_' + item
				arrayCopy[index] = t
				continue

			temp = ''
			if item.find('~') != -1:
				temp = item.split('~')
				rimNumber = get_rim_number(temp[1], soName)
			else:
				rimNumber = get_rim_number(item, soName)

			# If RIM Number not found, just skip.
			if not rimNumber:
				continue

			if temp:
				mappedItem = replace_signal_name_doi(temp[1], rimNumber)
				t = '~R' + rimNumber + '_' + mappedItem
			else:
				mappedItem = replace_signal_name_doi(item, rimNumber)
				t = 'R' + rimNumber + '_' + mappedItem
			arrayCopy[index] = t

		string = ''
		for index, item in enumerate(arrayCopy):
			string = string + ' ' + item
			if index <= len(arrayCopy) - 2:
				if item.find('D') != -1:
					if arrayCopy[index+1].find('D') != -1 or arrayCopy[index+1].find('(') != -1:
						string = string + ' &'
				elif item.find(')') != -1:
					if arrayCopy[index+1].find('(') != -1 or arrayCopy[index+1].find('D') != -1:
						string = string + ' &'

		f = expr(string)
		f = f.simplify()
		f = f.factor()

		dependencyMap[soName] = [mainIlk, str(f)]

	# Write data to excel file.
	book = Workbook('ILKEqns-8-14.xlsx')
	sheet = book.add_worksheet('Flattened Interlock Equations')

	sheet.write(0, 0, 'SO')
	sheet.write(0, 1, 'ILK')
	sheet.write(0, 2, 'Dependency')
	# Set column width.
	sheet.set_column(2, 2, 300)
	sheet.set_column(0, 0, 10)
	
	write_data(sheet, soArray, dependencyMap)

	book.close()

if __name__ == '__main__':
	main()
