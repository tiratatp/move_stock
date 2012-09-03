#!/usr/bin/env python
# encoding: utf-8
# @NuttyKnot with ♥
# 20120903
 
import xlrd, sys, ConfigParser, math
from operator import itemgetter, attrgetter
from collections import deque

if len(sys.argv) < 2:
	print "usage: %s <SOH excel filename>" % sys.argv[0]
	exit(1)

# check if config.txt exist, if not create
try:
	with open('config.txt') as f: pass
except IOError as e:
	print "cannot find config.txt, creating default config.txt!"
	with open('config.txt', 'w') as f:
		f.write("""[Basic]
# put blank if there is no branch in promotion
PromotionBranch = R2
# minimum ageing to consider moving; currently not using this value
MinAgeing = 180
# number of item per promotional branch per sku
PromotionalItemPerBranch = 2

[Branches]
# ensure there is no space between branch and comma
A = CL,ZW,LP
B = BN,PK,PY,PT,R2,CT,SC
C = RI,RS,KS,HY,RH

# @NuttyKnot with ♥""")

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

min_ageing = config.getint('Basic', 'MinAgeing')
# min promotional item per branch
promotional_item_per_branch = config.getint('Basic', 'PromotionalItemPerBranch')

# open excel
filename = sys.argv[1]
workbook = xlrd.open_workbook(filename, on_demand=True)
worksheet = workbook.sheet_by_index(1)

# open output file
output = open('%s.csv' % filename, 'w')
log = open('%s.log' % filename, 'w')

def printlog(str):
	print str
	log.write("%s\n" % str)

branch_map = {}
reverse_branch_map = {}

# find and map all branches name and its column numbers
for i in xrange(25,worksheet.ncols,2):
	cell_type = worksheet.cell_type(1, i)
	cell_value = worksheet.cell_value(1, i).strip()
	printlog('"%s" <=> column#: %s' % (cell_value, i))
	branch_map[cell_value] = i
	reverse_branch_map[i] = cell_value

# convert column to class
def colToBranchClass(col):
	branch = reverse_branch_map[col][:2]
	for class_desc in class_name:
		if branch in class_set[class_desc]:
			return class_desc
	printlog("branch is not in any class!")
	assert False # branch is not in any class

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
			item = (reverse_branch_map[j], item_count, int(worksheet.cell_value(i, j+1)))
			if item_count > 1:
				more_than_one.append(item)
			else:
				temp.append(item)
		more_than_one = sorted(more_than_one, key=itemgetter(2), reverse=True)		
		temp = sorted(temp, key=itemgetter(2), reverse=True)
		# return list of branch that has this item
		# sort by
		# 1. branch that has more than one item
		# 2. branch that has oldest ageing
		more_than_one.extend(temp)
		return more_than_one
	sorted_branch = []
	for class_desc in reversed(class_name):
		sorted_branch.extend(loopInClass(class_col_list[class_desc]))
	printlog("Sorted: %s" % deque(sorted_branch))
	return deque(sorted_branch)

# loop through each SKU
# start at row 3
for i in xrange(3,worksheet.nrows):
	percent_marked_down = worksheet.cell_value(i, 21) # get %Mark down
	sku = worksheet.cell_value(i, 5)
	# skip no mark-down or no promotional branch
	if percent_marked_down != "0.00%" and number_of_promotion_branch != 0:
		pocket = set()
		pocketed_item_count = 0
		# loop through branches
		for j in xrange(25,worksheet.ncols,2):
			cell_type = worksheet.cell_type(i, j)
			# filter blank branch
			if cell_type == 0 or cell_type == 6:
				continue
			item_count = int(worksheet.cell_value(i, j))
			if item_count == 0:
				continue

			pocket.add((reverse_branch_map[j], item_count, int(worksheet.cell_value(i, j+1)), colToBranchClass(j))) # add to pocket

		if len(pocket) == 0:
			continue

		# sort items
		# order by ageing then by class
		pocket = sorted(pocket, key=itemgetter(2, 3), reverse=True)
		# convert to deque
		pocket = deque(pocket)

		printlog("Sorted: %s" % pocket)

		# count sum of item in pocket
		count_pocket = reduce(lambda x,y:x+int(y[1]), pocket , 0)
		min_pocketed_item = number_of_promotion_branch * promotional_item_per_branch
		pocketed_item_count = count_pocket if count_pocket < min_pocketed_item else min_pocketed_item

		# avg_item_per_branch = promotional_item_per_branch if there is abundant of that sku
		# else
		# avg_item_per_branch = pocketed_item_count/number_of_promotion_branch

		# distribute
		avg_item_per_branch = math.ceil(float(pocketed_item_count) / number_of_promotion_branch)
		current_promotion_branch_index = 0
		current_promotion_branch = promotion_list[current_promotion_branch_index]
		current_promotion_branch_count = 0
		while len(pocket) > 0:			
			item = pocket.popleft()
			# if item in this branch has more than avg, split
			if item[1] > avg_item_per_branch:
				pocket.appendleft((item[0], item[1] - avg_item_per_branch))
				item = (item[0], avg_item_per_branch)
			output.write('%s,%s,%s,%s\n' % (sku, item[0], current_promotion_branch,item[1]))
			printlog('%s: %s,%s,%s,%s # promotional move' % (i, sku, item[0], current_promotion_branch,item[1]))
			current_promotion_branch_count = current_promotion_branch_count + item[1]
			if current_promotion_branch_count >= avg_item_per_branch:
				current_promotion_branch_index = current_promotion_branch_index + 1
				try:
					current_promotion_branch = promotion_list[current_promotion_branch_index]
				except IndexError:
					break
				current_promotion_branch_count = 0
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
					printlog('%s: # %s out-of-stock at %s' % (i, sku, reverse_branch_map[j]))
					continue					
				if left[1] > 1:
					sorted_branch[0] = left[0], left[1] - 1
				else:
					sorted_branch.popleft()
				output.write('%s,%s,%s,%s\n' % (sku, left[0], reverse_branch_map[j], 1))
				printlog('%s: %s,%s,%s,%s # out-of-stock move' % (i, sku, left[0], reverse_branch_map[j], 1))

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

