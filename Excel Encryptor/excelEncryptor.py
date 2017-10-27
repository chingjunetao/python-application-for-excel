# -*- coding: utf-8 -*-
"""
Created on Mon Oct 9 10:28:26 2017
@author: June Tao Ching
"""

import openpyxl as xl
from win32com.client import Dispatch
import os
from shutil import copyfile

#Ignore Excel format warning
import warnings
warnings.filterwarnings("ignore")

try:
    #Check config file
    if(os.path.exists('ContentConfig.txt') == False):
        raise Exception('Warning! Config file does not exist! \n')
    
    #Get input from user
    inputFilename = input("Please enter game document name >>")
    
    outputFilename = os.path.splitext(inputFilename)[0] + '_lite' + os.path.splitext(inputFilename)[1]
    
    print('Encrypting... \n')
    
    outputFolder = 'lite version'
    if(os.path.exists(outputFolder) == False):
        os.makedirs(outputFolder)
    
    copyfile(inputFilename,outputFolder + '\\' + outputFilename)
    
    workbook = xl.load_workbook(outputFolder + '\\' + outputFilename)
    allSheets = workbook.get_sheet_names()
    
    #read Copy Content
    with open('ContentConfig.txt', 'r') as myfile:
        content = myfile.read().split('\n')
    
    sheetToHide = list(set(allSheets)-set(content))
    
    #VBA
    xl = Dispatch("Excel.Application")
    xl.Visible = False  # You can remove this line if you don't want the Excel application to be visible
    
    
    outputPath = os.getcwd() + '\\' + outputFolder + '\\' + outputFilename
    outputWorkbook = xl.Workbooks.Open(outputPath)
    
    for sheet in sheetToHide:
        outputWorkbook.Worksheets(sheet).visible = False
    
    
    outputWorkbook.Protect('aasdfA#%',True,True)
    outputWorkbook.Close(SaveChanges = True)
    xl.Quit()
    
    print('Encryption is completed. \n')
    
except Exception as e:
    print(e);
    print('\n')