# -*- coding: utf-8 -*-
"""
Created on Sun Mar  1 20:47:50 2020

@author: justi
"""

# Compiles all the fit scores into a single file

import xlrd # Library to read/write xlsx files
import xlsxwriter

# read_file_name Name of fully processed data file to read from (shortened, will be added to)
WRITE_FILE_NAME = "compiledFits.xlsx"
header = ["Trial","Node","chi^2","r^2"]
alum_fit_data = list() #[file][node][r_score/chi^2]
bras_fit_data = list()

def extractData(read_file_name): # Reads file and records data
    for i in range(11, 21): # Iterate through files
        file_data = list() # List of fit values for current file
        read_file_name = read_file_name + str(i)+".xlsx" # Contruct current filename
        file = xlrd.open_workbook(read_file_name) # Open file for reading
        sheet_names = file.sheet_names() # Get the names of the sheets
        for sheet in sheet_names: # Iterate through the sheets
            worksheet = file.sheet_by_name(sheet) # Get the current worksheet
            chi = float(worksheet.row(1)[10].value) # Get chi^2 value
            r = float(worksheet.row(1)[9].value) # Get R^2 value
            file_data.append([r,chi]) # Store data
        read_file_name = read_file_name[0:12] # Cut off ending of name
        if (i < 17): # Determine which list to append to
            alum_fit_data.append(file_data)
        else:
            bras_fit_data.append(file_data)
    return

def writeData(): # Writes the data to an xlsx file
    extractData("analysedData") # Get data
    xlsx_file = xlsxwriter.Workbook(WRITE_FILE_NAME) # Open xlsx file and create worksheets
    alum_worksheet = xlsx_file.add_worksheet("Aluminum")
    alum_worksheet.write_row(0,0,header) # Write header
    bras_worksheet = xlsx_file.add_worksheet("Brass")
    bras_worksheet.write_row(0,0,header)
    row = 1 # Counter for which row to write to
    for i in range(0,len(alum_fit_data)): # Iterate through aluminum files
        current_trial = alum_fit_data[i] # Get current trial data
        for j in range(0,len(current_trial)): # Iterate through nodes
            alum_worksheet.write_row(row,0,[i+1,j+1,current_trial[j][1],current_trial[j][0]]) # Write data
            row += 1 # Iterate the row counter
    row = 1 # Counter for which row to write to
    for i in range(0,len(bras_fit_data)): # Same process for brass
        current_trial = bras_fit_data[i] # Get current trial data
        for j in range(0,len(current_trial)): # Iterate through nodes
            bras_worksheet.write_row(row,0,[i+1,j+1,current_trial[j][1],current_trial[j][0]]) # Write data
            row += 1 # Iterate the row counter
    xlsx_file.close() # Close file
    return

writeData()
