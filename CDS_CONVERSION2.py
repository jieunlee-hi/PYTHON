# -*- coding:utf-8 -*-
import re
from datetime import datetime, timedelta
import time
import traceback
import pandas as pd
import sys, os, traceback
from selenium.webdriver.common.keys import Keys
import os
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
import sys, os, traceback, glob
import win32com.client as win32
import sys
import os.path

outpath = "C:\\CDS/"


def CONVERSION2():
    if len(sys.argv) is 1:
        filename = glob.glob('CDS_result2.xls')  # There is no option.
        filename = ''.join(filename)
        print(filename)
        if os.path.isfile(outpath + filename + "x") == True:
            os.unlink(outpath + filename + "x")

        excel = win32.gencache.EnsureDispatch('Excel.Application')
        fn_ext = os.path.splitext(filename)
        print('Input file name: %s' % fn_ext[0])
        print('Input file ext: %s' % fn_ext[1])

        if fn_ext[1] == '.xls':
            out_filename = fn_ext[0] + '.xlsx'
        else:
            out_filename = fn_ext[0] + fn_ext[1] + '.xlsx'

        print('Out file name: %s' % out_filename)
        wb = excel.Workbooks.Open(outpath + filename)
        wb.SaveAs(outpath + out_filename, FileFormat=56)  # FileFormat 56 is for .xls extension
        wb.Close()
        excel.Application.Quit()
try :
    CONVERSION2()
    #DELETE()
except:
    pass