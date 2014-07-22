from xlrd import open_workbook
from xlsxwriter import Workbook
from pyeda.inter import *

from Tree import *

wb = open_workbook(filename='2.xls', formatting_info=True)
matrixSheet = wb.sheet_by_name('Interlock Matrix')
stringsSheet = wb.sheet_by_name('Interlock Strings')
rimSheet = wb.sheet_by_name('IO Mapping Table')
interlockNumArray = []
whereUsedNumArray = []
ioMemNumArray = []

def get_rim_number(name):
	signalColumn = rimSheet.col_values(0)
	rimColumn = rimSheet.col_values(32)

	for index, item in enumerate(signalColumn):
		if name.find(item) != -1:
			if rimColumn[index].find('RIM') != -1:
				temp = rimColumn[index].split()
				return str(temp[1])

def write_data(sheet, soArray, dependencyMap):
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
		if index == len(array) - 1:
			tree.set_extra_close(pointer)
		pointer = tree.insert_child(pointer, item)

def match_helper(col, soName):
	for item in col:
		if item.find(soName) != -1:
			return True
	return False
##
# found soName in given search columms?
def match(col1, col2, soName):
	found1 = match_helper(col1, soName)
	found2 = match_helper(col2, soName)
	if found1 or found2:
		return True
	return False

def helper(name):
	for num in interlockNumArray:
		if name == stringsSheet.cell(num, 11).value:
			return int(interlockNumArray.index(num))
	return None

def get_io_members(name):
	temp = []
	# find interlock on worksheet
	index = helper(name)
	if index == None:
		raise Exception('InterlockDoesNotExistException')
	# list I/O members of the interlock
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
	searchCol1 = stringsSheet.col_values(11)
	searchCol2 = stringsSheet.col_values(26) 
	for s in searchCol1:
		if s.find('SO_') != -1:
			soArray.append(s[:6])
	for s in searchCol2:
		if s.find('SO_') != -1:
			soArray.append(s[:6])
	soArray.sort()
	return soArray

def main():
	#wb = open_workbook(filename='1.xls', formatting_info=True)
	#matrixSheet = wb.sheet_by_name('Interlock Matrix')
	soArray = generate_soArray()
	doiArray = matrixSheet.row_values(2, 14, 146)

	# number of rows and columns in matrix
	nrows = len(soArray) # 176
	ncols = len(doiArray) # 132

	colourMap = dict()

	'''
	# dark blue rgb(0, 51, 102)
	# light blue rbd(204, 255, 255)
	for i, so in enumerate(soArray):
		dependencyList = []
		# matirx begins at row 3, col 14
		for j, doi in enumerate(doiArray):
			xfx = matrixSheet.cell_xf_index(i+3, j+14)
			xf = wb.xf_list[xfx]
			bgx = xf.background.pattern_colour_index
			rgb = wb.colour_map[bgx]
			if rgb == (0, 51, 102):
				dependencyList.append(doi)

		colourMap[so] = dependencyList

	# list of rownums for interlock number and where used
	#stringsSheet = wb.sheet_by_name('Interlock Strings')
	#interlockNumArray = []
	#whereUsedNumArray = []
	#ioMemNumArray = []
	'''
	filterArray = stringsSheet.col_values(0)

	for index, item in enumerate(filterArray):
		if item == "Interlock Number:":
			interlockNumArray.append(index)
		elif item == "Where Used:":
			whereUsedNumArray.append(index)
		elif item == "I/O Members:":
			ioMemNumArray.append(index)

	dependencyMap = dict()

	for soName in soArray:
		dependencyMap[soName] = []
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
		dependencyMap[soName].append(stringsSheet.cell(interlockNumArray[index], 11).value)
		start = ioMemNumArray[index] + 1
		stop = whereUsedNumArray[index] - 1
		c1 = stringsSheet.col_values(11, start, stop)
		c2 = stringsSheet.col_values(26, start, stop)
		for doName in c1:
			if doName == '':
				continue
			dependencyMap[soName].append(doName)
		for doName in c2:
			if doName == '':
				continue
			dependencyMap[soName].append(doName)

	
	for soName in soArray:
		array = dependencyMap[soName]
		ilk = array[0]
		mainIlk = ilk
		array = array[1:]

		if mainIlk == 'ILK_01':
			dependencyMap[soName] = [mainIlk, 'DI_000']
			continue

		tree = Tree()
		p = tree.insert_root(ilk)
		
		while tree.has_ilk():
			pointer, ilk = tree.get_ilk()
			loop(tree, pointer, ilk)

		tree.reverse_children()

		arrayCopy = tree.format_list()

		##
		# Add the rim number generation function here
		for index, item in enumerate(arrayCopy):
			rimNumber = get_rim_number(item)
			if not rimNumber:
				continue
			t = 'R' + rimNumber + '_' + item
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


	# write data to excel file
	book = Workbook('ILKEqns.xlsx')
	sheet = book.add_worksheet('Flattened Interlock Equations')

	sheet.write(0, 0, 'SO')
	sheet.write(0, 1, 'ILK')
	sheet.write(0, 2, 'Dependency')
	# set width
	sheet.set_column(2, 2, 300)
	
	write_data(sheet, soArray, dependencyMap)

	book.close()

if __name__ == '__main__':
	main()
