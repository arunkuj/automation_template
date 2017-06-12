'''
    This is the main file which drives the Automation Framework.
    Usage : python main.py
'''

import sys
import os
import re
import unittest
import subprocess

import HTMLTestRunner
import xlrd

from collections import OrderedDict
from time import sleep
from time import localtime
from ConfigParser import ConfigParser

# append path into system path
testDir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.normpath(os.path.join(testDir, "..")))

# and change working directory to automation folder
os.chdir(testDir)

# input files
confile = 'config.ini'
tc_file = 'testcases.txt'
testcase_folder = 'testcases'

# Get the time stamp
dictionary = {'1': 'Jan', 
              '2': 'Feb',
              '3': 'Mar',
              '4': 'Apr',
              '5': 'May',
              '6': 'Jun',
              '7': 'Jul',
              '8': 'Aug',
              '9': 'Sept',
              '10': 'Oct',
              '11': 'Nov',
              '12': 'Dec'}
time_stamp = localtime()
Test_Start_Time = (str(time_stamp.tm_mday) + 
                   dictionary[str(time_stamp.tm_mon)] + 
                   str(time_stamp.tm_year) + '_' + 
                   str(time_stamp.tm_hour) + 'hr' + 
                   str(time_stamp.tm_min) + 'min' + 
                   str(time_stamp.tm_sec) + 'sec')

input_dict = {}
dup_dict = {}
outputfile = None

if not confile in os.listdir(testDir):
    print "Error : Config file does not exist in %s" % testDir
    sys.exit()

# parse config.ini
def read_config(file):
    config = ConfigParser()
    config.read(file)   # this takes care of closing the file too

    for key, val in config.items('input_info'):
        input_dict[key] = val


def get_tc_to_be_run():
    """
    parse the testcase.xls to get the scripts to be run
    """
    try:
        book = xlrd.open_workbook('testcases.xls')
    except:
       print "\nFAIL:ailed to open testcases.xls"
       sys.exit(-1)
    else:
        sheet = book.sheet_by_name('Features')
        for i in range(sheet.nrows):
            row = sheet.row_values(i)
            if 'Feature' in row:
                feature_index = row.index('Feature')
                selected_index = row.index('Selected')

                list1 = sheet.col_values(feature_index)[1:]
                list2 = sheet.col_values(selected_index)[1:]

                features = []

                dictionary = OrderedDict(zip(list1, list2))

                for key, value in dictionary.iteritems():
                    if value == 'Yes' or value == 'YES':
                        features.append(key)

        testcases = []

        for feature in features:
            sheet = book.sheet_by_name(feature)
            for i in range(sheet.nrows):
                row = sheet.row_values(i)
                if 'TC ID' in row:
                    tcid_index = row.index('TC ID')
                    selected_index = row.index('Selected')

                    list3 = sheet.col_values(tcid_index)[1:]
                    list4 = sheet.col_values(selected_index)[1:]

                    list3_int = []

                    for i in list3:
                        list3_int.append(int(i))

                    dictionary = OrderedDict(zip(list3_int, list4))

                    for key, value in dictionary.iteritems():
                        if value == 'Yes' or value == 'YES':
                            testcases.append(key)
        return testcases    

def main():
    config_data = read_config(confile)
    testcases_to_be_run = get_tc_to_be_run()

    inputReportAttrs = list()
    outputfile = None

    if testcases_to_be_run == ['']:
        print "EXIT:No TC ID's have been mentioned in testcases.txt file"
        sys.exit()
    
    if 'outputfile' in config_data:
        outputfile = input_dict['outputfile']
    else:
        outputfile = 'Report.html'

    if os.path.isfile(outputfile):
        os.remove(outputfile)

    testRunner = None
    f = open(outputfile, 'w')
    testRunner = HTMLTestRunner.HTMLTestRunner(
                                stream=f,
                                title=input_dict['suite'], verbosity=2,
                                inputReportAttrs=inputReportAttrs,
                                )

    suites = unittest.TestSuite()

    for testcaseid in testcases_to_be_run:
        pattern = 'test_' + testcaseid + '.py'
        test = unittest.defaultTestLoader.discover(testcase_folder, pattern)

        if test.countTestCases() == 0:
            print "Error : Module %s doesn't exist in test case folder"\
                  % pattern
            sys.exit()
        if test.countTestCases() > 1:
            print dir(test)
            print "Error : Module %s exist in %s location, Please verify "\
                  % (pattern, test.countTestCases())
            sys.exit()
        if test.countTestCases() == 1:
            suites.addTest(test)

    testSuite = unittest.TestSuite(suites)
    testRunner.run(testSuite)


if __name__ == '__main__':
    main()
