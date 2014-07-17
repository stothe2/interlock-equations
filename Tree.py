class Tree:

	class _Node:
		def __init__(self, data):
			self.data = data
			self.parent = None
			self.children = list()
			# black, True: visited
			# red, False: not visited
			self.colour = False

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

	def _convert(self, p):
		# check for isinstance
		if not p:
			raise Exception('not really a position')
		# hmm, not really typecasting
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

	def get_ilk(self):
		array = self._to_list(self._root, [])
		for item in array:
			if item.data.find('ILK_') != -1 and not item.colour:
				item.colour = True
				return item, item.data[8:]
		raise Exception('NoILKPresentException')

	def _to_list(self, n, array):
		if n:
			array.append(n)
			if self.has_children(n):
				for child in n.children:
					n = child
					self._to_list(n, array)
		return array

	def to_string(self):
		array = self._to_list(self._root, [])
		string = ''
		for item in array:
			string = string + ' ' + item.data
		return string

'''
	def __iter__(self):
		return self

	def __next__(self):
		root = self._root
		if not root:
			StopIteration
		else:
			# ...
'''
