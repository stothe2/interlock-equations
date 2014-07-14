class BinarySearchTree:

	class _Node:
		def __init__(self, key, value):
			self.key = key
			self.value = value
			self._left = None
			self._right = None

	def __init__(self):
		self._root = None

	def _insert(self, n, k, v):
		if not n:
			x = self._Node(k, v)
			return x
		if (k < n.key):
			n._left = self._insert(n._left, k, v)
			# check balance?
		elif (k > n.key):
			n._right = self._insert(n._right, k, v)
			# check balance?
		else:
			# illegal argument
			raise ValueError('duplicate key ' + k)
		return n

	def insert(self, k, v):
		self._root = self._insert(self._root, k, v)

	def get(self, k):
		n = self._find_for_sure(k)
		return n.value

	def _find_for_sure(self, k):
		n = self._find(k)
		if not n:
			raise KeyError('unknown key ' + str(k))
		return n

	def _find(self, k):
		n = self._root
		while n:
			if k < n.key:
				n = n._left
			elif k > n.key:
				n = n._right
			else:
				return n

	def _has(self, n, k):
		if not n:
			return False
		if k < n.key:
			return _has(n.left, k)
		elif k > n.key:
			return _has(n.right, k)
		else:
			return True

	def has(self, k):
		return self._has(self._root, k)

	def remove(self, k):
		value = self.get(k)
		self._root = self._remove(self._root, k)
		return value

	def _remove(self, n, k):
		if not n:
			raise KeyError('unknown key ' + str(k))
		if k < n.key:
			n._left = self._remove(n._left, k)
		elif k > n.key:
			n._right = self._remove(n._right, k)
		else:
			n = self._remove_two(n)
		return n

	def _remove_two(self, n):
		if not n._left:
			return n._right
		if not n._right:
			return n._left

		# ...

	def pre(self):
		self._preorder_traversal(self._root)

	def _preorder_traversal(self, n):
		if n:
			print(n.value)
			self._preorder_traversal(n._left)
			self._preorder_traversal(n._right)


	def __iter__(self):
		return self
