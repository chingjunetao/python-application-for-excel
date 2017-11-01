# -*- coding: utf-8 -*-
"""
Created on Mon Oct 9 10:28:26 2017
@author: June Tao Ching
"""

import openpyxl as xl
from win32com.client import Dispatch
import os

try:
    #Check config file
    if(os.path.exists('ContentConfig.txt') == False):
        raise Exception('Warning! Config file does not exist! \n')
        
    inputFilename = input("Please enter document name >>")
    outputFilename = os.path.splitext(inputFilename)[0] + '_lite' + os.path.splitext(inputFilename)[1]
    
    #Create an empty output file
    if (os.path.exists(outputFilename)):
        os.remove(outputFilename)    
    outputWorkbook = xl.Workbook()
    outputWorkbook.save(outputFilename)
    
    print('Copying... \n')
    #read Copy Content
    with open('ContentConfig.txt', 'r') as myfile:
        content = myfile.read().split('\n')
    
    xl = Dispatch("Excel.Application")
    xl.Visible = False  # You can remove this line if you don't want the Excel application to be visible
    
    inputPath = os.getcwd() + '\\' + inputFilename
    outputPath = os.getcwd() + '\\' + outputFilename
    
    inputWorkbook = xl.Workbooks.Open(inputPath)
    outputWorkbook = xl.Workbooks.Open(outputPath)
    
    #Copy contents to output file
    i = 1
    for files in reversed(content):
        ws = inputWorkbook.Worksheets(files)
        ws.Copy(Before = outputWorkbook.Worksheets(1))
        i = i + 1
    
    #Delete extra sheet 
    extraWorksheet = outputWorkbook.Worksheets('Sheet')
    extraWorksheet.Delete()
    
    print('Complete Copying! \n')
    
    inputWorkbook.Close(SaveChanges = False)
    outputWorkbook.Close(SaveChanges = True)
    xl.Quit()

except Exception as e:
    print(e);
    print('\n')
