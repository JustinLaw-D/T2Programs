# -*- coding: utf-8 -*-
"""
Created on Thu Feb 20 13:13:52 2020

@author: justi
"""

# Copies over every tenth row to a new file

import csv

CSV_READ_FILE_NAME = "rawdata19s.csv" # Name of raw csv document
CSV_WRITE_FILE_NAME = "processedData19.csv" # Name of processed csv document
DATE = "21/02/2020" # Date of data collection in dd/mm/yyyy format
MATERIAL = "Brass" # String containing the name of the material used

names = ["Time","TempOne","TempTwo","TempThree","TempFour","TempFive","TempSix","TempSeven","Atmospheric"] # Array of names for columns

def readWrite():
    csvWriterFile = open(CSV_WRITE_FILE_NAME, 'w', newline='') # Open csv file for writing
    csvWriter = csv.writer(csvWriterFile) # Open csv writer for constants
    csvWriter.writerow(["Date", DATE, "Material", MATERIAL]) # Write constant to file
    csvWriter = csv.DictWriter(csvWriterFile, fieldnames=names) # Open csv writer roper
    csvWriter.writeheader() # Self-explanatory
    csvReaderFile = open(CSV_READ_FILE_NAME, 'r', newline='')
    csvReader = csv.DictReader(csvReaderFile, fieldnames=names) # Open dictionary writer for rawdata
    write = 0 # Counter that determines which rows are copied
    start_time = 0 # Initial time
    for row in csvReader:
        if write < 2: # Skip the first two row
            pass
        elif ((write - 2) % 10) == 0: # Every tenth row, starting at row 2
            if write == 2: # First time around, grab the base time
                start_time = float(row["Time"])
            row["Time"] = (float(row["Time"]) - start_time)/1000 # Normalize to zero and convert to seconds
            csvWriter.writerow(row) # Write the processed row to the new file
        write += 1 # Update write counter
    csvReaderFile.close() # Close and save files
    csvWriterFile.close()

readWrite()
