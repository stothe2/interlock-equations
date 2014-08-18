class Tree:

	class _Node:
		def __init__(self, data):
			self.data = data
			self.parent = None
			self.children = list()
			# The colour is used to determine if a node has been
			# visited before or not. This is necessary, but only helpful
			# for "ILK" nodes when flattening the equations.
			# black, True: visited
			# red, False: not visited
			self.colour = False
			# extra data (for formatting)
			self.extra = False

	def __init__(self):
		self._root = None
		self._elements = 0

	def has_parent(self, p):
		n = self._convert(p)
		return n.parent != None

	def has_children(self, p):
		n = self._convert(p)
		return len(n.children) != 0

	def insert_root(self, data):
		if self._root:
			raise Exception('TreeNotEmptyException')
		self._root = self._Node(data)
		self._elements = self._elements + 1
		return self._root

	def insert_child(self, p, data):
		p = self._convert(p)
		n = self._Node(data)
		n.parent = p
		p.children.append(n)
		
		self._elements = self._elements + 1
		return n

	def root(self):
		if not self._root:
			raise Exception('EmptyTreeException')
		return self._root

	# Not really well-written. Would be more apt/necessary if the
	# script was written in Java.
	def _convert(self, p):
		# check for isinstance
		if not p:
			raise Exception('not really a position')
		# Hmm, not really typecasting...
		# Hey, don't judge!
		n = p
		return n

	def size(self):
		return self._elements

	def empty(self):
		return len(self._elements) == 0

	def pre(self):
		self._pre_traversal(self._root)

	def _pre_traversal(self, n):
		if n:
			print(n.data)
			if self.has_children(n):
				for child in n.children:
					n = child
					self._pre_traversal(n)

	def has_ilk(self):
		array = self._to_list(self._root, [])
		for item in array:
			if item.data.find('ILK_') != -1 and not item.colour:
				return True
		return False

	# Finds ILK name in the tree which hasn't ever been "visited",
	# and returns a pointer to that node and the ILK name.
	def get_ilk(self):
		array = self._to_list(self._root, [])
		for item in array:
			# ILK branch not visited till present?
			if item.data.find('ILK_') != -1 and not item.colour:
				# Set current ILK branch to visited once.
				item.colour = True
				# Only 'ILK'?
				if len(item.data) <= 7:
					ilk = item.data
				# 'NOT_ILK'?
				elif item.data.find('NOT_ILK_') != -1:
					ilk = item.data[12:]
				# gotta check...
				else:
					ilk = item.data[8:]
				return item, ilk
		raise Exception('NoILKPresentException')

	def to_list(self):
		return self._to_list(self._root, [])

	def _to_list(self, n, array):
		if n:
			array.append(n)
			if self.has_children(n):
				for child in n.children:
					n = child
					self._to_list(n, array)
		return array

	def format_list(self):
		array = self._format_list(self._root, [])
		arrayCopy = list()
		for index, item in enumerate(array):
			if item.find('ILK_') != -1 or item.find('DI_000') != -1:
				continue
			if item.find('NOT_DO_') != -1:
				temp = item.split(':')
				item = temp[0]
				item = '~' + item[4:]
			else:
				temp = item.split(':')
				item = temp[0]
			arrayCopy.append(item)
		return arrayCopy

	def _format_list(self, n, array):
		if n:
			array.append(n.data)
			if self.has_children(n):
				if n.data.find('NOT_ILK_') != -1:
					array.append('~(')
				elif n.data.find('ILK_') != -1:
					array.append('(')
				for child in n.children:
					n = child
					self._format_list(n, array)
					if n.extra:
						array.append(')')
		return array

	def set_extra_close(self, n):
		n.extra = True

	def reverse_children(self):
		self._reverse_children(self._root)

	def _reverse_children(self, n):
		if n:
			if self.has_children(n):
				n.children.reverse()
				for child in n.children:
					n = child
					self._reverse_children(n)
