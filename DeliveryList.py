#! /usr/local/bin/python3

''' DeliveryList Class
	Purpose: host delivery tracking records
	Usage: use add to add it and use tuple operation to access data
'''

from collections import namedtuple

# Record class is used for each record that is hosted in DeliveryList to ease access
Record = namedtuple('Record', ['id', 'item', 'detail', 'migration_plan', 'key_ask', 'owner', 'due_date', 'status'])
# set default data, please note color code of this item is inconsistant
DEFAULTDLVALUE = (-1, 'N/A', 'N/A'  , 'N/A'  , 'N/A'  , 'N/A' , 'Mar 28, 2018', '0x00ff00', 'Open')
Record.__new__.__defaults__ = DEFAULTDLVALUE

class DeliveryList():
	''' class to host all delivery tracking records
	Value added function: get_color_code
	'''
	def __init__(self):
		self._records = [] 
		self.cdict = { 'Open':0x800000, 'New':0x008080, 'Ongoing': 0x008080, 'Close': 0x008000, 'Done': 0x008000 }
		self.titlelist = ('Id', 'Item', 'Detail', 'MigrationPlan', 'KeyAsk', 'Owner', 'DueDate', 'C', 'Status')
	def get_title(self):
		return self.titlelist
		# return Record()._fields

	def get_color_code(self,position):
		''' Get color code per status info
		'''
		if self._records[position].status in self.cdict:
			return self.cdict[self._records[position].status]
		else:
			return 0x000000
	
	def __len__(self):
		return len(self._records)

	def __getitem__(self,position):
		return self._records[position]
	
def print_dl(dl):
	''' Fucntion to print sample content
	'''
	for s in dl.get_title():
		print("%s" % s, end='\t')
	print()
	for x in dl._records:
		pos = dl._records.index(x)
		print("%d %s %s %06x" % (x.id, x.item, x.status, dl.get_color_code(pos)))

# unit test code below
if __name__ == '__main__':
	dl = DeliveryList()
	strp = ''
	for title in dl.get_title():
		strp += title
		strp += ' '
	print("%s" % strp)
	print("%d records so far" % len(dl))
	value1 = (1, 'MN FaskTrack', 'N/A'  , 'N/A'  , 'N/A'  , 'N/A' , 'Mar 28, 2018', 'Done')
	value2 = (2, 'ION FaskTrack', 'N/A'  , 'N/A'  , 'N/A'  , 'N/A' , 'Mar 28, 2018', 'Open')
	dl._records.append(Record(*value1))
	dl._records.append(Record(*value2))
	
	print_dl(dl)
