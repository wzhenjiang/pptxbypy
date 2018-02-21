#! /usr/local/bin/python3

''' this is a sample code to test list
'''

import collections

# use named tuple
Card = collections.namedtuple('Card', ['suit', 'rank'])

# unit test code below
if __name__ == '__main__':
	suits = ['Spade', 'Heart', 'Diamond', 'Clube']
	ranks = [ str(rank) for rank in range(2,11) ] + 'J Q K'.split()

	deck = [ Card(suit, rank) for suit in suits for rank in ranks]

	print( "%s" % deck )
