# Copyright (c) 2014 Giulio De Pasquale
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
#to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
#copies of the Software, and to permit persons to whom the Software is
#furnished to do so, subject to the following conditions:
#
#The above copyright notice and this permission notice shall be included in
#all copies or substantial portions of the Software.
#
#THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
#IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
#FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
#AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
#LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
#OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
#THE SOFTWARE.

__author__ = 'Giulio De Pasquale'

from sys import argv
from tempfile import TemporaryFile
from collections import defaultdict

import xlrd
from xlwt import Workbook


def printBanner():
    ver = "1.0.1"
    print "Thanks for using XLS Polisher " + str(ver) + "\n\n"


def printChoicesMenu():
    allowed_choices = [1, 2, 3, 4, 5, 8, 9]
    print "1) REMOVE a column"
    print "2) SELECT a filter on a specific column"
    print "3) LIST COLUMNS marked to be deleted"
    print "4) LIST active FILTERS"
    print "5) LIST Column's UNIQUE Elements"
    print "8) APPLY changes"
    print "9) EXIT"
    return allowed_choices


def getColIdxFromName(name):
    colNamesDict = getColTitles()
    for key, value in colNamesDict.iteritems():
        if name == value:
            return key


def printColTitles():
    for col in getColTitles().values():
        if col not in col_names_to_delete:
            print col


def getColTitles():
    """
    [just the title]
    """
    colNamesDict = {}
    for i in range(0, sheet.ncols):
        colNamesDict[i] = sheet.cell_value(0, i)
    return colNamesDict


def printDelPendingColumns():
    if len(col_names_to_delete) > 0:
        print "There is(are) " + str(len(col_names_to_delete)) + " column(s) pending deletion, which is(are):"
        for name in col_names_to_delete:
            print name
    else:
        print "There are no columns pending deletion."


def printFilters():
    if len(col_filter_strict) or len(col_filter_loose) or len(col_filter_delete_strict) or len(
            col_filter_delete_loose) > 0:
        print "There are these ACTIVE filters:"
        if len(col_filter_loose) > 0:
            print "SHOW ROWS CONTAINING:"
            for key, value in col_filter_loose.iteritems():
                for content in value:
                    print "COL: " + str(key) + " - FILTER: " + str(content)
        if len(col_filter_strict) > 0:
            print "SHOW ROWS STRICTLY CONTAINING:"
            for key, value in col_filter_strict.iteritems():
                for content in value:
                    print "COL: " + str(key) + " - FILTER: " + str(content)
        if len(col_filter_delete_loose) > 0:
            print "DELETE ROWS CONTAINING:"
            for key, value in col_filter_delete_loose.iteritems():
                for content in value:
                    print "COL: " + str(key) + " - FILTER: " + str(content)
        if len(col_filter_delete_strict) > 0:
            print "DELETE ROWS STRICTLY CONTAINING:"
            for key, value in col_filter_delete_strict.iteritems():
                for content in value:
                    print "COL: " + str(key) + " - FILTER: " + str(content)
    else:
        print "There are NO ACTIVE filters."


def listElements():
    mne = []
    printColTitles()
    print "Which column would you like to inspect?"
    col = getCharInput(getColTitles().values())
    col = getColIdxFromName(col)
    print "Would you like to print every element? [Y/N]"
    choice = getCharInput(["y", "n"])
    for row in range(1, sheet.nrows):
        val = sheet.cell_value(row, col)
        if val not in mne:
            mne.append(val)
    if choice == "y":
        for val in mne:
            print val
    print "Number of elements: " + str(len(mne))


def removeColumn():
    print "Which column would you like to DELETE? [MULTIPLE ENTRIES DELIMITED BY \" | \""
    printColTitles()
    col_to_delete = raw_input()
    col_to_delete = col_to_delete.split(" | ")
    for name in col_to_delete:
        if name in getColTitles().values():
            col_names_to_delete.append(name)
        else:
            print "Sorry but I can't find " + str(name) + " in the file you supplied."
    return


def addFilter():
    print "Which column would you like to inspect?"
    printColTitles()
    col_to_filter = getCharInput(getColTitles().values())
    if col_to_filter not in col_names_to_delete:
        col_to_filter = getColIdxFromName(col_to_filter)
        print "Would you like to DELETE rows with particular content? [Y/N]"
        choice = getCharInput(["y", "n"])
        if choice.lower() == "y":
            src_content = raw_input("What would you like to DELETE?[MULTIPLE ENTRIES DELIMITED BY \" | \"\n>> ")
            src_content = src_content.split(" | ")
            for filter in src_content:
                print "Are you sure that you want to DELETE \"" + filter + "\"? [Y/N]"
                choice = getCharInput(["y", "n"])
                if choice.lower() == "y":
                    print "Do you want to STRICTLY filter \"" + filter + "\"?"
                    choice = getCharInput(["y", "n"])
                    if choice.lower() == "y":
                        col_filter_delete_strict[col_to_filter].append(filter)
                    else:
                        col_filter_delete_loose[col_to_filter].append(filter)
        elif choice.lower() == "n":
            src_content = raw_input("What would you like to FILTER?[MULTIPLE ENTRIES DELIMITED BY \" | \"\n>> ")
            src_content = src_content.split(" | ")
            for filter in src_content:
                print "Do you want to STRICTLY filter \"" + filter + "\"?"
                choice = getCharInput(["y", "n"])
                if choice.lower() == "y":
                    col_filter_strict[col_to_filter].append(filter)
                else:
                    col_filter_loose[col_to_filter].append(filter)
            return


def getWorkingSheetFromFile(filename):
    xls = xlrd.open_workbook(filename)
    # Opens the first sheet in the XLS file
    sheet = xls.sheet_by_index(0)
    return sheet


def populateColNumsToDelete(colNamesDict):
    for key, value in colNamesDict.iteritems():
        if value in col_names_to_delete:
            col_nums_to_delete.append(key)


def populateRowNumsToDelete():
    # checking user selected rows
    if len(col_filter_delete_strict) > 0:
        for row in range(0, sheet.nrows):
            for col in range(0, sheet.ncols):
                if col in col_filter_delete_strict and sheet.cell_value(row, col) in col_filter_delete_strict[col]:
                    row_nums_to_delete.append(row)

    # checking strict filter
    if len(col_filter_strict) > 0:
        for row in range(1, sheet.nrows):
            for col in range(0, sheet.ncols):
                if col in col_filter_strict and not sheet.cell_value(row, col) in col_filter_strict[col]:
                    row_nums_to_delete.append(row)

    if len(col_filter_loose) > 0:
        for row in range(1, sheet.nrows):
            for col in range(0, sheet.ncols):
                if col in col_filter_loose:
                    for name in col_filter_loose[col]:
                        if name not in sheet.cell_value(row, col):
                            row_nums_to_delete.append(row)

    if len(col_filter_delete_loose) > 0:
        for row in range(1, sheet.nrows):
            for col in range(0, sheet.ncols):
                if col in col_filter_delete_loose:
                    for name in col_filter_delete_loose[col]:
                        if name in sheet.cell_value(row, col):
                            row_nums_to_delete.append(row)


def writeFile(pf_sheet):
    # actual row/col to write since there may be some rows/cols that have to be jumped
    populateColNumsToDelete(getColTitles())
    populateRowNumsToDelete()
    col_write = 0
    col_wrote = False
    row_write = 0
    for col in range(0, sheet.ncols):
        for row in range(0, sheet.nrows):
            if col not in col_nums_to_delete and row not in row_nums_to_delete:
                if not col_wrote:
                    col_wrote = True
                pf_sheet.write(row_write, col_write, sheet.cell_value(row, col))
                row_write += 1
        if col_write >= (sheet.ncols - len(col_nums_to_delete)):
            col_write = 0
        if col_wrote:
            col_write += 1
            col_wrote = False
        row_write = 0


def mainLoop():
    print "What would you like to do? "
    allowed_choices = printChoicesMenu()
    choice = getInput(allowed_choices)
    while choice != 9:
        if choice == 1:
            removeColumn()
        elif choice == 2:
            addFilter()
        elif choice == 3:
            printDelPendingColumns()
        elif choice == 4:
            printFilters()
        elif choice == 5:
            listElements()
        elif choice == 8:
            filename = raw_input("How would you like to name your new file?\n>> ")
            if ".xls" in filename:
                filename.strip(".xls")
            polished_file = Workbook()
            # Creates sheet with original name
            pf_sheet = polished_file.add_sheet(sheet.name)
            writeFile(pf_sheet)
            polished_file.save(filename + ".xls")
            polished_file.save(TemporaryFile())
            return
        printChoicesMenu()
        print "What would you like to do? "
        choice = getInput(allowed_choices)


def getInput(allowed):
    choice = input(">> ")
    while choice not in allowed:
        choice = input(">> ")
    return choice


def getCharInput(allowed):
    choice = raw_input(">> ")
    while choice.lower() not in allowed and choice not in allowed and choice.upper() not in allowed:
        print "What you entered was not valid."
        choice = raw_input(">> ")
    return choice

####
# FI VARIABLES
####

col_names_to_delete = []
col_nums_to_delete = []
row_nums_to_delete = []
col_filter_delete_strict = defaultdict(list)
col_filter_strict = defaultdict(list)
col_filter_loose = defaultdict(list)
col_filter_delete_loose = defaultdict(list)

####
# MAIN
####

if len(argv) > 1:
    file = argv[1]
else:
    file = raw_input("Hello which XLS file would you like to parse?")

sheet = getWorkingSheetFromFile(file)
printBanner()
mainLoop()