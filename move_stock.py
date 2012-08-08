#!/usr/bin/env python
# encoding: utf-8
 
import xlrd, sys, ConfigParser
from operator import itemgetter, attrgetter
from collections import deque

if len(sys.argv) < 2:
	print "usage: %s <SOH excel filename>" % sys.argv[0]
	exit(1)

# read and parse config
config = ConfigParser.SafeConfigParser()
config.read('config.txt')

# branch classes
class_a_list = config.get('Branches', 'A').split(',')
class_a = set(class_a_list)
class_b_list = config.get('Branches', 'B').split(',')
class_b = set(class_b_list)
class_c_list = config.get('Branches', 'C').split(',')
class_c = set(class_c_list)

all_branch = class_a | class_b | class_c

if len(all_branch) != len(class_a) + len(class_b) + len(class_c):
	print "duplicated branch in class!"
	exit(1)

non_c_branch_list = []
non_c_branch_list.extend(class_a_list)
non_c_branch_list.extend(class_b_list)

promotion_list = config.get('Basic', 'PromotionBranch').split(',')
promotion = set(promotion_list)
number_of_promotion_branch = len(promotion)

min_aging = config.get('Basic', 'MinAging')
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

# find col# of each branch order by branch class
def branchDescToCode(branch_list):
	branch_code_list = []
	for code in branch_list:
		try:
			branch_code_list.append(branch_map['%sCT' % code])
		except KeyError:
			pass
	return branch_code_list

non_c_branch_col_list = branchDescToCode(non_c_branch_list)
class_c_col_list = branchDescToCode(class_c_list)
class_b_col_list = branchDescToCode(class_b_list)
class_a_col_list = branchDescToCode(class_a_list)

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
		# 2. branch that has oldest aging
		more_than_one.extend(temp)
		return more_than_one
	sorted_branch = []
	sorted_branch.extend(loopInClass(class_c_col_list))
	sorted_branch.extend(loopInClass(class_b_col_list))
	sorted_branch.extend(loopInClass(class_a_col_list))
	return deque(sorted_branch)

# loop through each SKU
# start at row 3
for i in xrange(3,worksheet.nrows):
	percent_marked_down = worksheet.cell_value(i, 21) # get %Mark down
	sku = worksheet.cell_value(i, 5)
	if percent_marked_down != "0.00%":
		pocket = set()
		optional_pocket = set()
		pocketed_item_count = 0
		# loop through branches
		# pick all "old" item
		# while store "non-old" item in seperated list
		for j in xrange(25,worksheet.ncols,2):
			cell_type = worksheet.cell_type(i, j)
			# filter blank branch
			if cell_type == 0 or cell_type == 6:
				continue
			item_count = int(worksheet.cell_value(i, j))
			if item_count == 0:
				continue
			item = (reverse_branch_map[j], item_count)
			if(int(worksheet.cell_value(i, j+1)) >= min_aging):
				pocket.add(item) # add to pocket
				pocketed_item_count = pocketed_item_count + item_count
			else:
				optional_pocket.add(item) # add to pocket
		min_pocketed_item =  number_of_promotion_branch * min_item_per_branch

		# if picked item is not enough,
		# sort by aging desc and pick item
		if len(optional_pocket) > 0 and pocketed_item_count < min_pocketed_item:
			sorted_optional_pocket = sorted(optional_pocket, key=itemgetter(1), reverse=True)
			while pocketed_item_count < min_pocketed_item and len(sorted_optional_pocket) > 0:
				item = sorted_optional_pocket.pop()
				pocket.add(item)
				pocketed_item_count = pocketed_item_count + item[1]
		#print pocket

		# optional
		#pocket = sorted(pocket, key=itemgetter(1), reverse=True)
		#print pocket

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
			#print '%s: %s,%s,%s,%s # promotional move' % (i, sku, item[0], current_promotion_branch,item[1])
			current_promotion_branch_count = current_promotion_branch_count + item[1]
			if current_promotion_branch_count > avg_item_per_branch:
				current_promotion_branch_index = current_promotion_branch_index + 1
				current_promotion_branch = promotion_list[current_promotion_branch_index]
	else:
		# loop through branches 
		for j in non_c_branch_col_list:
			# filter blank branch
			sorted_branch = []
			cell_type = worksheet.cell_type(i, j)			
			if cell_type == 0 or cell_type == 6 or (cell_type == 2 and int(worksheet.cell_value(i, j)) == 0):
				if len(sorted_branch) == 0:
					sorted_branch = sortBranch(i)					
				left = sorted_branch[0]
				if left[1] > 1:
					sorted_branch[0] = left[0], left[1] - 1
				else:
					sorted_branch.popleft()
				output.write('%s,%s,%s,%s\n' % (sku, left[0], reverse_branch_map[j], 1))
				print '%s: %s,%s,%s,%s # out-of-stock move' % (i, sku, left[0], reverse_branch_map[j], 1)

output.close()				

exit()

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

