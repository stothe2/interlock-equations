from xlrd import open_workbook
from pyeda.inter import *

from Tree import *

wb = open_workbook(filename='1.xls', formatting_info=True)
matrixSheet = wb.sheet_by_name('Interlock Matrix')
stringsSheet = wb.sheet_by_name('Interlock Strings')
interlockNumArray = []
whereUsedNumArray = []
ioMemNumArray = []

def loop(p, check, tree, checkArray, ilk):
	if not checkArray:
		p, ilk = tree.get_ilk()
	if ilk.find('NOT_') != -1:
		ilk = ilk[4:]
	array = get_io_members(ilk)
	insert(p, array, tree, check)
	check = tree.has_ilk()

def insert(pointer, array, tree, check):
	for item in array:
		pointer = tree.insert_child(pointer, item)
		if item.find('ILK') != -1:
			ilk = item[8:]
			loop(pointer, check, tree, True, ilk)

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
	if not index:
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

def main():
	#wb = open_workbook(filename='1.xls', formatting_info=True)
	#matrixSheet = wb.sheet_by_name('Interlock Matrix')
	soArray = matrixSheet.col_values(1, 3, 179)
	doiArray = matrixSheet.row_values(2, 14, 146)

	# number of rows and columns in matrix
	nrows = len(soArray) # 176
	ncols = len(doiArray) # 132

	colourMap = dict()

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
	filterArray = stringsSheet.col_values(0)
	#interlockNumArray = []
	#whereUsedNumArray = []
	#ioMemNumArray = []
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

	'''
	for soName in soArray:
		array = dependencyMap[soName]
		ilkName = array[0]
		array = array[1:]
		newArray = loop(array, [])
		dependencyMap[soName] = newArray
	'''

	array = dependencyMap['SO_030']
	ilkName = array[0]
	array = array[1:]

	tree = Tree()

	for index, doi in enumerate(array):
		if index == 0:
			pointer = tree.insert_root(doi)
		else:
			pointer = tree.insert_child(pointer, doi)

	check = tree.has_ilk()
	while check:
		check = loop(None, check, tree, None, '')

	print(dependencyMap['SO_030'])
	tree.pre()

	finalList = ilkName + tree.to_string()

	print(finalList)

if __name__ == '__main__':
	main()
