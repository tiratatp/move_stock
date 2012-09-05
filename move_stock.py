#!/usr/bin/python
# -*- coding: utf-8 -*-

# @NuttyKnot with ♥
# 20120903

print """
move_stock, a program to move stock in Central Department Store, Thailand
Copyright (C) 2012  Tiratat Patana-anake

This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program; if not, write to the Free Software
Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
"""

import xlrd
import sys
import ConfigParser
import math
import os
from operator import itemgetter
from collections import deque
from os import path

if len(sys.argv) < 2:
    print 'usage: %s <SOH excel filename>' % sys.argv[0]
    sys.exit(1)

# check if config.txt exist, if not create

config_path = path.join(path.dirname(sys.argv[0]), 'config.txt')
try:
    with open(config_path) as f:
        pass
except IOError, e:
    print 'cannot find config.txt, creating default config.txt!'
    with open(config_path, 'w') as f:
        f.write("""[Basic]
# put blank if there is no branch in promotion
PromotionBranch = R2
# number of item per promotional branch per sku
PromotionalItemPerBranch = 1

[Branches]
# ensure there is no space between branch and comma
A = CL,ZW,LP
B = BN,PK,PY,PT,R2,CT,SC
C = RI,RS,KS,HY,RH

# For move_stock, Tiratat Patana-anake""")

# read and parse config

config = ConfigParser.SafeConfigParser()
config.read(config_path)

all_branch = set()
class_list = {}
class_name = ['A', 'B', 'C']

# branch classes

for class_desc in class_name:
    class_list[class_desc] = config.get('Branches',
            class_desc).split(',')
    all_branch = all_branch | set(class_list[class_desc])

if len(all_branch) != reduce(lambda x, y: x + len(y),
                             class_list.itervalues(), 0):
    print 'duplicated branch in class!'
    sys.exit(1)

non_c_branch_list = []
for class_desc in ['A', 'B']:
    non_c_branch_list.extend(class_list[class_desc])

promotion_config = config.get('Basic', 'PromotionBranch')
promotion_list = promotion_config.split(',')
number_of_promotion_branch = (0 if len(promotion_config)
                              == 0 else len(promotion_list))

promotional_item_per_branch = config.getint('Basic',
        'PromotionalItemPerBranch')

# open excel

filename = sys.argv[1]
workbook = xlrd.open_workbook(filename, on_demand=True)
worksheet = workbook.sheet_by_index(1)  # second workbook

# open output file

output = open('%s.csv' % filename, 'w')
log = open('%s.log' % filename, 'w')


def printlog(str):
    print str
    log.write('%s\n' % str)


branch_map = {}
reverse_branch_map = {}

# find and map all branches name and its column numbers

for i in xrange(25, worksheet.ncols, 2):
    cell_type = worksheet.cell_type(1, i)
    cell_value = worksheet.cell_value(1, i).strip()
    printlog('"%s" <=> column#: %s' % (cell_value, i))
    branch_map[cell_value] = i
    reverse_branch_map[i] = cell_value


# convert column to class

def colToBranchClass(col):
    branch = (reverse_branch_map[col])[:2]
    for class_desc in class_name:
        if branch in class_list[class_desc]:
            return class_desc
    return None


# find col# of each branch order by branch class

def branchDescToCode(branch_list):
    return filter(None, map(lambda x: (branch_map['%sCT' % x] if '%sCT'
                  % x in branch_map else None), branch_list))


non_c_branch_col_list = branchDescToCode(non_c_branch_list)
class_col_list = {}
for class_desc in class_name:
    class_col_list[class_desc] = \
        branchDescToCode(class_list[class_desc])


# for out-of-stock move

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
            item = (reverse_branch_map[j], item_count,
                    int(worksheet.cell_value(i, j + 1)))
            if item_count > 1:
                more_than_one.append(item)
            else:
                temp.append(item)
        more_than_one = sorted(more_than_one, key=itemgetter(2),
                               reverse=True)
        temp = sorted(temp, key=itemgetter(2), reverse=True)

        # return list of branch that has this item
        # sort by
        # 0. class
        # 1. branch that has more than one item
        # 2. branch that has oldest ageing

        more_than_one.extend(temp)
        return more_than_one

    sorted_branch = []
    for class_desc in reversed(class_name):
        sorted_branch.extend(loopInClass(class_col_list[class_desc]))
    printlog('Sorted: %s' % deque(sorted_branch))
    return deque(sorted_branch)


# loop through each SKU
# start at row 3

for i in xrange(3, worksheet.nrows):
    percent_marked_down = worksheet.cell_value(i, 21)  # get %Mark down
    sku = worksheet.cell_value(i, 5)

    # skip no mark-down or no promotional branch

    if percent_marked_down != '0.00%' and number_of_promotion_branch \
        != 0:
        pocket = set()
        pocketed_item_count = 0

        # loop through branches

        for j in xrange(25, worksheet.ncols, 2):
            cell_type = worksheet.cell_type(i, j)

            # filter blank branch

            if cell_type == 0 or cell_type == 6:
                continue
            item_count = int(worksheet.cell_value(i, j))
            if item_count == 0:
                continue

            # if pocketed item branch == promotional branch, pick another pocketed item

            if (reverse_branch_map[j])[:2] in promotion_list:
                continue
            branch_class = colToBranchClass(j)
            if branch_class == None:
                continue
            pocket.add((reverse_branch_map[j], item_count,
                       int(worksheet.cell_value(i, j + 1)),
                       branch_class))  # add to pocket

        if len(pocket) == 0:
            continue

        # sort items
        # order by ageing then by class

        pocket = sorted(pocket, key=itemgetter(2, 3), reverse=True)

        # convert to deque

        pocket = deque(pocket)

        printlog('Sorted: %s' % pocket)

        # count sum of item in pocket

        count_pocket = reduce(lambda x, y: x + int(y[1]), pocket, 0)
        min_pocketed_item = number_of_promotion_branch \
            * promotional_item_per_branch
        pocketed_item_count = (count_pocket if count_pocket
                               < min_pocketed_item else min_pocketed_item)

        # avg_item_per_branch = promotional_item_per_branch if there is abundant of that sku
        # else
        # avg_item_per_branch = pocketed_item_count/number_of_promotion_branch

        # distribute

        done = False
        avg_item_per_branch = math.ceil(float(pocketed_item_count)
                / number_of_promotion_branch)
        current_promotion_branch_index = -1
        current_promotion_branch_count = None

        while len(pocket) > 0 and not done:

                      # get current stock at promotion branch
                      # if current stock >= avg_item_per_branch, ignore this branch

            while current_promotion_branch_count == None \
                or current_promotion_branch_count \
                >= avg_item_per_branch:
                current_promotion_branch_index = \
                    current_promotion_branch_index + 1
                try:
                    current_promotion_branch = \
                        promotion_list[current_promotion_branch_index]
                except IndexError:
                    done = True
                    break
                try:
                    current_promotion_branch_count = \
                        int(worksheet.cell_value(i, branch_map['%sCT'
                            % current_promotion_branch]))
                except ValueError:
                    current_promotion_branch_count = 0
            if done:
                break
            item = pocket.popleft()

            # if item in this branch has more than avg, split............

            if item[1] > avg_item_per_branch:
                pocket.appendleft((item[0], item[1]
                                  - (avg_item_per_branch
                                  - current_promotion_branch_count)))
                item = (item[0], avg_item_per_branch
                        - current_promotion_branch_count)
            output.write('%s,%s,%sCT,%s\n' % (sku, item[0],
                         current_promotion_branch, item[1]))
            printlog('%s: %s,%s,%s,%s # promotional move' % (i, sku,
                     item[0], current_promotion_branch, item[1]))
            current_promotion_branch_count = \
                current_promotion_branch_count + item[1]
    else:

        # loop through branches

        sorted_branch = []
        has_calc_sorted_branch = False
        for j in non_c_branch_col_list:

            # filter blank branch............

            cell_type = worksheet.cell_type(i, j)
            if cell_type == 0 or cell_type == 6 or cell_type == 2 \
                and int(worksheet.cell_value(i, j)) == 0:
                if not has_calc_sorted_branch:
                    sorted_branch = sortBranch(i)
                    has_calc_sorted_branch = True
                if len(sorted_branch) > 0:
                    left = sorted_branch[0]
                else:

                    # out-of-stock

                    printlog('%s: # %s out-of-stock at %s' % (i, sku,
                             reverse_branch_map[j]))
                    continue
                if left[1] > 1:
                    sorted_branch[0] = (left[0], left[1] - 1)
                else:
                    sorted_branch.popleft()
                output.write('%s,%s,%s,%s\n' % (sku, left[0],
                             reverse_branch_map[j], 1))
                printlog('%s: %s,%s,%s,%s # out-of-stock move' % (i,
                         sku, left[0], reverse_branch_map[j], 1))

output.close()
log.close()
