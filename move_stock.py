#!/usr/bin/env python
# encoding: utf-8
# @NuttyKnot with ♥
# 20120903
 
import xlrd, sys, ConfigParser
from operator import itemgetter, attrgetter
from collections import deque

if len(sys.argv) < 2:
	print "usage: %s <SOH excel filename>" % sys.argv[0]
	exit(1)

# read and parse config
config = ConfigParser.SafeConfigParser()
config.read('config.txt')

all_branch = set()
class_list = {}
class_set = {}
class_name = ['A', 'B', 'C']

# branch classes
for class_desc in class_name:
	class_list[class_desc] = config.get('Branches', class_desc).split(',')
	class_set[class_desc] = set(class_list[class_desc])
	all_branch = all_branch | class_set[class_desc]

if len(all_branch) != reduce(lambda x,y: x+len(y), class_list.itervalues(), 0):
	print "duplicated branch in class!"
	exit(1)

non_c_branch_list = []
for class_desc in ['A', 'B']:
	non_c_branch_list.extend(class_list[class_desc])

promotion_config = config.get('Basic', 'PromotionBranch')
promotion_list = promotion_config.split(',')
number_of_promotion_branch = 0 if len(promotion_config) == 0 else len(promotion_list)

min_ageing = config.get('Basic', 'MinAgeing')
# min promotional item per branch
min_item_per_branch = config.get('Basic', 'MinItemPerBranch')

# open excel
filename = sys.argv[1]
workbook = xlrd.open_workbook(filename, on_demand=True)
worksheet = workbook.sheet_by_index(1)

# open output file
output = open('%s.txt' % filename, 'w')

branch_map = {}
reverse_branch_map = {}

# find and map all branches name and its column numbers
for i in xrange(25,worksheet.ncols,2):
	cell_type = worksheet.cell_type(1, i)
	cell_value = worksheet.cell_value(1, i).strip()
	print '"%s" <=> column#: %s' % (cell_value, i)
	branch_map[cell_value] = i
	reverse_branch_map[i] = cell_value

def colToBranchClass(col):
	branch = reverse_branch_map[col]
	for class_desc in class_name:
		if branch in class_set[class_desc]:
			return class_desc
	print "branch is not in any class!"
	assert false # branch is not in any class

# find col# of each branch order by branch class
def branchDescToCode(branch_list):
	return filter(None, map(lambda x: branch_map['%sCT' % x] if '%sCT' % x in branch_map else None, branch_list))
non_c_branch_col_list = branchDescToCode(non_c_branch_list)
class_col_list = {}
for class_desc in class_name:
	class_col_list[class_desc] = branchDescToCode(class_list[class_desc])

def sortBranch(row):
	i = row
	# loop in each class
	def loopInClass(col_list):
		more_than_one = []
		temp = []
		for j in col_list:	
			cell_type = worksheet.cell_type(i, j)
			# filter blank branch
			if cell_type == 0 or cell_type == 6:
				continue
			item_count = int(worksheet.cell_value(i, j))
			if item_count == 0:
				continue
			item = (reverse_branch_map[j], item_count)
			if item_count > 1:
				more_than_one.append(item)
			else:
				temp.append((reverse_branch_map[j], item_count))
		temp = sorted(temp, key=itemgetter(1), reverse=True)
		# return list of branch that has this item
		# sort by
		# 1. branch that has more than one item
		# 2. branch that has oldest ageing
		more_than_one.extend(temp)
		return more_than_one
	sorted_branch = []
	for class_desc in reversed(class_name):
		sorted_branch.extend(loopInClass(class_col_list[class_desc]))
	return deque(sorted_branch)

# loop through each SKU
# start at row 3
for i in xrange(3,worksheet.nrows):
	percent_marked_down = worksheet.cell_value(i, 21) # get %Mark down
	sku = worksheet.cell_value(i, 5)
	# skip no mark-down or no promotional branch
	if percent_marked_down != "0.00%" and number_of_promotion_branch != 0:
		pocket = set()
		optional_pocket = set()
		pocketed_item_count = 0
		# loop through branches
		# divide item into old and non-old
		for j in xrange(25,worksheet.ncols,2):
			cell_type = worksheet.cell_type(i, j)
			# filter blank branch
			if cell_type == 0 or cell_type == 6:
				continue
			item_count = int(worksheet.cell_value(i, j))
			if item_count == 0:
				continue
			item = (reverse_branch_map[j], item_count, )

			optional_pocket.add(item) # add to pocket

			"""
			if(int(worksheet.cell_value(i, j+1)) >= min_ageing):
				pocket.add(item) # add to pocket
				pocketed_item_count = pocketed_item_count + item_count
			else:
				optional_pocket.add(item) # add to pocket
			"""

		min_pocketed_item = number_of_promotion_branch * min_item_per_branch

		# sort old item
		sorted_optional_pocket = sorted(optional_pocket, key=itemgetter(1), reverse=True)
		while int(pocketed_item_count) < int(min_pocketed_item) and len(sorted_optional_pocket) > 0:
			item = sorted_optional_pocket.pop()
			pocket.add(item)
			pocketed_item_count = pocketed_item_count + item[1]

		"""
		# if picked item is not enough,
		# sort by ageing desc and pick item
		if len(optional_pocket) > 0 and pocketed_item_count < min_pocketed_item:
			sorted_optional_pocket = sorted(optional_pocket, key=itemgetter(1), reverse=True)
			while pocketed_item_count < min_pocketed_item and len(sorted_optional_pocket) > 0:
				item = sorted_optional_pocket.pop()
				pocket.add(item)
				pocketed_item_count = pocketed_item_count + item[1]
		"""

		#print pocket
		# optional
		#pocket = sorted(pocket, key=itemgetter(1), reverse=True)
		#print sorted_optional_pocket

		# distribute
		avg_item_per_branch = int(pocketed_item_count / number_of_promotion_branch)
		current_promotion_branch_index = 0
		current_promotion_branch = promotion_list[current_promotion_branch_index]
		current_promotion_branch_count = 0
		while len(pocket) > 0:			
			item = pocket.pop()
			# if item in this branch has more than avg, split
			if item[1] > avg_item_per_branch + 1:
				pocket.add((item[0], item[1] - avg_item_per_branch))
				item[1] = avg_item_per_branch
			output.write('%s,%s,%s,%s\n' % (sku, item[0], current_promotion_branch,item[1]))
			print '%s: %s,%s,%s,%s # promotional move' % (i, sku, item[0], current_promotion_branch,item[1])
			current_promotion_branch_count = current_promotion_branch_count + item[1]
			if current_promotion_branch_count > avg_item_per_branch:
				current_promotion_branch_index = current_promotion_branch_index + 1
				current_promotion_branch = promotion_list[current_promotion_branch_index]
	else:
		# loop through branches 
		sorted_branch = []
		has_calc_sorted_branch = False
		for j in non_c_branch_col_list:
			# filter blank branch			
			cell_type = worksheet.cell_type(i, j)			
			if cell_type == 0 or cell_type == 6 or (cell_type == 2 and int(worksheet.cell_value(i, j)) == 0):
				if not has_calc_sorted_branch:
					sorted_branch = sortBranch(i)					
					has_calc_sorted_branch = True
				if len(sorted_branch) > 0:
					left = sorted_branch[0]
				else:
					# out-of-stock
					print '%s: # %s out-of-stock at %s' % (i, sku, reverse_branch_map[j])
					continue					
				if left[1] > 1:
					sorted_branch[0] = left[0], left[1] - 1
				else:
					sorted_branch.popleft()
				output.write('%s,%s,%s,%s\n' % (sku, left[0], reverse_branch_map[j], 1))
				print '%s: %s,%s,%s,%s # out-of-stock move' % (i, sku, left[0], reverse_branch_map[j], 1)

output.close()				

exit()

# debuging tool
num_rows = worksheet.nrows - 1
num_cells = worksheet.ncols - 1

curr_row = 1047
while curr_row < num_rows:
	curr_row += 1
	row = worksheet.row(curr_row)
	curr_cell = -1
	while curr_cell < num_cells:
		curr_cell += 1
		# Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank
		cell_type = worksheet.cell_type(curr_row, curr_cell)
		cell_value = worksheet.cell_value(curr_row, curr_cell)
		print '(%s,%s) %s:%s' % (curr_row, curr_cell, cell_type, cell_value)
	if curr_row>10:
		exit()

