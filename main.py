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

controlsWb = open_workbook(filename='Signal List_11.xlsx')

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
			## this is due to capitalization difference!
			#if name == 'SO_179':
			#	return 'DO00'
			if name == 'SO_201':
				return name
			if name == 'SO_209':
				return name
			## this is due to a space in Spare_HEATER _SSR_17! wtf
			if name == 'SO_218':
				return 'DO108'
			## this is due to AMPDS_vaporizer_inlet_valve)open_cmd! ftw
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
				temp = rimColumn[index].split()
				return str(temp[1])

	for index, item in enumerate(doColumn):
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
		# obtain RIM# for SO
		rimNumber = get_rim_number(item, item)
		# Replace SO name with the name we use
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
			# temporary hack for DO_0265
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

	# write data to excel file
	book = Workbook('ILKEqns.xlsx')
	sheet = book.add_worksheet('Flattened Interlock Equations')

	sheet.write(0, 0, 'SO')
	sheet.write(0, 1, 'ILK')
	sheet.write(0, 2, 'Dependency')
	# set width
	sheet.set_column(2, 2, 300)
	sheet.set_column(0, 0, 10)
	
	write_data(sheet, soArray, dependencyMap)

	book.close()

if __name__ == '__main__':
	main()
