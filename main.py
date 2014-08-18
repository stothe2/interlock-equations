from xlrd import open_workbook
from xlsxwriter import Workbook
from pyeda.inter import *

from Tree import *

# Workbook created by __.
# Might need update <--
wb = open_workbook(filename='2.xls', formatting_info=True)
matrixSheet = wb.sheet_by_name('Interlock Matrix')
stringsSheet = wb.sheet_by_name('Interlock Strings')
# Array containing positions (int) of all rows with 'ILK'
interlockNumArray = []
# Array containing positions (int) of all rows with 'Where Used'
whereUsedNumArray = []
# Array containing starting row number (int) of rows containing I/O members
ioMemNumArray = []

def write_data(sheet, soArray, dependencyMap):
	# Keeps track of what row you're writing data to.
	rownum = 1
	for item in soArray:
		data = dependencyMap[item]
		ilk = data[0]
		equation = data[1]
		sheet.write(rownum, 0, item)
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
	book = Workbook('ILKEqns.xlsx')
	sheet = book.add_worksheet('Flattened Interlock Equations')

	sheet.write(0, 0, 'SO')
	sheet.write(0, 1, 'ILK')
	sheet.write(0, 2, 'Dependency')
	# Set column width.
	sheet.set_column(2, 2, 300)
	
	write_data(sheet, soArray, dependencyMap)

	book.close()

if __name__ == '__main__':
	main()
