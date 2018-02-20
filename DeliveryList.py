#! /usr/local/bin/python3

''' DeliveryList Class
	Purpose: host delivery tracking records
	Usage: use add to add it and use tuple operation to access data
'''

import collections

# Record class is used for each record that is hosted in DeliveryList to ease access
Record = collections.namedtuple('Record', ['Id', 'Item', 'Detail', 'MigrationPlan', 'KeyAsk', 'Owner', 'DueDate', 'C', 'Status'])
# set default data, please note color code of this item is inconsistant
Record.__new__.__defaults__ = (-1, 'N/A', 'N/A'  , 'N/A'  , 'N/A'  , 'N/A' , 'Mar 28, 2018', '0x00ff00', 'Open') 

class DeliveryList():
	''' class to host all delivery tracking records
	'''
	def __init__(self):
		self._records = [] 

	def get_title(self):
		return Record()._fields
	
	def add(self,record):
		self._records.append(record)

	def __len__(self):
		return len(self._records)

	def __getitem__(self,position):
		return self._records[position]
		
# unit test code below
if __name__ == '__main__':
	dl = DeliveryList()
	strp = ''
	for title in dl.get_title():
		strp += title
		strp += ' '
	print("%s" % strp)
	print("%d records so far" % len(dl))
	dl.add(Record())
	print("%d records so far" % len(dl))
