# -*- coding: utf-8 -*-
"""
Created on Wed Jan 15 22:36:33 2020

@author: justin
"""

import serial
import csv

## Reads data being sent from arduino for single temperature sensor

CSV_FILE_NAME = "rawdata5q.csv" # Name of csv document
DATE = "25/02/2020" # Date of data collection in dd/mm/yyyy format
BASE_TEMPERATURE = "18.37/5.88" # String containing the inside and outside temperatures

temperatures = list() ## List of readings
currentReading = list() ## Current reading
names = ["Time","Temperature"] # Array of names for columns

ardPort = serial.Serial("COM3", 9600) ## Initializes serial port named com3, baud rate 9600

def terminationSequence():
    done = False # Done writing data?
    readingWrite = dict() # Dictionary for writing information to csv
    csvFile = open(CSV_FILE_NAME, 'w', newline='') # Open csv file for writing
    csvWriter = csv.writer(csvFile) # Open csv writer for constants
    csvWriter.writerow(["Date", DATE, "Inside/Outside", BASE_TEMPERATURE]) # Write constant to file
    csvWriter = csv.DictWriter(csvFile, fieldnames=names) # Open csv writer roper
    csvWriter.writeheader() # Self-explanatory
    for i in range (len(temperatures)): # Iterate through readings
        currentReading = temperatures[i] # Get reading
        counter = 0 # Counter for reading string index
        for j in range(0, len(names)): # Iterate through values
            temp = "" # New temp string
            if (currentReading[counter] == '{'): #If new reading
                counter += 1 # Go to next character
                while (currentReading[counter] != '}'): # While not the end of the reading
                    temp += currentReading[counter] # Add to string
                    counter += 1 # Update index
            if ((counter == len(currentReading) - 1) and j != len(names) - 1): # If we're done before filling the row, we're done
                done = True
                break
            counter += 1 # To get to the next thing
            temp = float(temp) # Convert to float
            if (j != 0):
                temp += 273.15 # If its a temperature convert to Kelivn
            readingWrite[names[j]] = temp
        if (not done):
            csvWriter.writerow(readingWrite) # Write information to new row
    csvFile.close() # Close file, save information

while True:
    try:
        while (ardPort.in_waiting > 0): # If there's data to read, read it
            temp = ardPort.read().decode() # Read character and convert to unicode
            if (temp == '<'): # If new reading
                currentReading.clear() # Clear list for new reading
            elif (temp == '>'): # If end of reading  
                temperatures.append(currentReading.copy()) # Add to list of temperature readings
            else:
                currentReading.append(temp)
    except (serial.SerialException): ## Exception thrown if port is closed
        terminationSequence() ## If the port is closed, call termination sequence
        ardPort.close()
        break # End program
        
