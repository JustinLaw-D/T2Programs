# -*- coding: utf-8 -*-
"""
Created on Sat Feb 22 16:19:51 2020

@author: justi
"""

## This program takes processed data and model data, and creates seven linear functions in excel
## one for each temperature, and calculates the R^2 value of each fit.
## Specifically looks at last five nodes.

import csv
import xlsxwriter
from math import log as ln
from math import pi

# Data in both files should be normalized to start at T = 0s, and run at the same intervals for the same amount of time
CSV_DATA_FILE_NAME = "processedData15.csv" # Name of csv document with processed data
CSV_MODEL_FILE_NAME = "modelData15.csv" # Name of the csv document with model data
XLSX_SHEETS_FILE_NAME = "linear_fit15.xlsx" # Name of the xlsx document with the linear fit data
names = ["Time","TempOne","TempTwo","TempThree","TempFour","TempFive","TempSix","TempSeven","Atmospheric"] # Array of names for columns
header = ["Time", "DataTemperature", "ModelTemperature", "DataTemperature - Atmospheric", "ModelTemperature - Atmospheric", "ln(Data)", "ln(Model)"]
t_alpha = 9.4*10**-5 # Theoretical alpha value for the data
L = 0.3048 # Length of the rod

def caclulateCutOff(): # Calculates the cut off time for a linear fit
    t = (2*(L**2)*ln(10))/(3*t_alpha*(pi**2)) # Calculate t value
    return t

def extractData(): # Pulls data from the csv files and stores them in lists
    data_file = open(CSV_DATA_FILE_NAME, 'r', newline='') # Open data file for reading
    data_reader = csv.DictReader(data_file, fieldnames=names) # Open data file reader
    measured_data = list() # List to hold experimental data
    skip = 0 # To skip first two rows
    for row in data_reader:
        if skip < 2: # Skip the first two rows
            skip += 1
            continue
        measured_data.append(row) # Otherwise, add the data to the list
    data_file.close() # Close file
    ## Same process, but with model data
    data_file = open(CSV_MODEL_FILE_NAME, 'r', newline='')
    data_reader = csv.DictReader(data_file, fieldnames=names)
    model_data = list()
    skip = 0
    for row in data_reader:
        if skip < 2:
            skip += 1
            continue
        model_data.append(row)
    data_file.close()
    return measured_data, model_data # Return the data lists

def getTime(measured_data): # Stores the time data in extractData as its own list
    ## This function just extists so that I only have to deal with time once
    times = list() # Create list
    for reading in measured_data: # Loops through the readings
        times.append(float(reading["Time"])) # Get time and add it to the list
    return times # Return the list of times

def processData(measured_data, model_data, times): # Processes the extracted data to be ready for writing for a given temperature
    ## This returns a 3d list. d1 = temperature number, d2 = data/model/data-atm/mode-atm/ln(data-atm)/ln(model-atm), d3 = datapoint
    ## First, build the list
    processed_data = list() # The 3d list
    for i in range (0,7): # Append the seven tempreature lists
        processed_data.append(list())
        for k in range(0,7): # Append the five data columns to the temperature list
            processed_data[i].append(list())
            
    ## Then, input the data
    atmospheric = float(measured_data[0]["Atmospheric"]) # Get atmospheric temperature
    for index in range(0, len(measured_data)): # Iterates through every row of the data
        for i in range(0,7): # Iterate through temperatures
            ## Get model and experimental data
            measured = float(measured_data[index][names[i+1]])
            model = float(model_data[index][names[i+1]])
            processed_data[i][0].append(measured) # Stores the data
            processed_data[i][1].append(model)
            processed_data[i][2].append(measured - atmospheric)
            processed_data[i][3].append(model - atmospheric)
            processed_data[i][4].append(ln(measured-atmospheric))
            processed_data[i][5].append(ln(model-atmospheric))
            
    ## Then, cut if off at the calculated point
    t_cut = caclulateCutOff()
    while (times[0] < t_cut): # Iterates through times while they need to be cut
        times.pop(0) # Remove the current time
        for k in range(0,7): # Iterate through temperatures
            for i in range(0,6): # Iterate through datapoints
                processed_data[k][i].pop(0) # Remove the reading at the current time
            
    return processed_data, times

def setupSheets(atmospheric): # Sets up the xlsx file
    xlsx_file = xlsxwriter.Workbook(XLSX_SHEETS_FILE_NAME) # Create file
    worksheets = list() # List of xlsx_file worksheets
    for i in range(1,8): # Loop to initialize the seven sheets, one for each temperature
        worksheets.append(xlsx_file.add_worksheet("Temperature " + str(i))) # Add worksheet with proper name
    for i in range(0,7): # Loop through sheets and add headers
        worksheets[i].write_row(0,0,header) # Write the header to the sheet
        worksheets[i].write('I1','Atmospheric')
        worksheets[i].write('I2',atmospheric) # Write atmospheric temperature
    return xlsx_file, worksheets

def makeNormalCharts(xlsx_file, worksheets, num): # Generates plots of the normal (non-ln) data
    ## Num is the number of datapoints for each temperature
    for i in range(0,7): # Iterate through worksheets/nodes
        worksheet = worksheets[i] # Get current worksheet
        chart = xlsx_file.add_chart({'type':'scatter'}) # Add scatter chart to the excel document
        chart.set_title({'name':'Temperature(K) vs Time(s)'}) # Add title and axis titles to chart
        chart.set_x_axis({'name':'Time(s)'})
        chart.set_y_axis({'name':'Temperature(K)'})
        chart.add_series({'name':'Data', # Name of the series
                          'values':['Temperature '+str(i+1),1,3,1+num,3], # y-values, sheet for temperature 1 from [row, column] to [row,column]
                          'categories':['Temperature '+str(i+1),1,0,1+num,0], # x-values, same format
                          'marker':{'type':'circle','size':2} # Set markers to small circles
        })
        chart.add_series({'name':'Model', # Name of the series
                          'values':['Temperature '+str(i+1),1,4,1+num,4], # y-values, sheet for temperature 1 from [row, column] to [row,column]
                          'categories':['Temperature '+str(i+1),1,0,1+num,0], # x-values, same format
                          'marker':{'type':'square','size':2,'color':'red'} # Set markers to squares
        })
        worksheet.insert_chart('J5', chart) # Insert chart into worksheet
    return

def makeLinearCharts(xlsx_file, worksheets, num): # Generate plots of the ln data
    for i in range(0,7): # Iterate through worksheets/nodes
        worksheet = worksheets[i] # Get current worksheet
        chart = xlsx_file.add_chart({'type':'scatter'}) # Add scatter chart to the excel document
        chart.set_title({'name':'ln(Temperature) vs Time(s)'}) # Add title and axis titles to chart
        chart.set_x_axis({'name':'Time(s)'})
        chart.set_y_axis({'name':'ln(Temperature)'})
        chart.add_series({'name':'Data', # Name of the series
                          'values':['Temperature '+str(i+1),1,5,1+num,5], # y-values, sheet for temperature 1 from [row, column] to [row,column]
                          'categories':['Temperature '+str(i+1),1,0,1+num,0], # x-values, same format
                          'marker':{'type':'circle','size':2}, # Set markers to small circles
                          'trendline':{'type':'linear',
                                       'display_equation':True,
                                       'display_r_squared':True,
                                       'forward':40,
                                       'backward':40,
                                       'line':{'dash_type':'long_dash'}
                                       } # Display trendline
        })
        chart.add_series({'name':'Model', # Name of the series
                          'values':['Temperature '+str(i+1),1,6,1+num,6], # y-values, sheet for temperature 1 from [row, column] to [row,column]
                          'categories':['Temperature '+str(i+1),1,0,1+num,0], # x-values, same format
                          'marker':{'type':'square','size':2,'color':'red'}, # Set remove markers
                          'trendline':{'type':'linear',
                                       'display_equation':True,
                                       'display_r_squared':True,
                                       'forward':40,
                                       'backward':40,
                                       'line':{'dash_type':'long_dash'}
                                       } # Display trendline
        })
        worksheet.insert_chart('J20', chart) # Insert chart into worksheet
    return

def fillSheets(): # Fills the xlsx file with the data
    measured_data, model_data = extractData() # Get all necessary data and file handles
    times = getTime(measured_data)
    processed_data, times = processData(measured_data, model_data, times)
    xlsx_file, worksheets = setupSheets(float(measured_data[0]["Atmospheric"]))
    for i in range(0,7): # Index for temperatures
        worksheet = worksheets[i] # Get current worksheet
        dataset = processed_data[i] # Get current node's data
        worksheet.write_column(1,0,times) # Write times to column 0, starting at row 1
        for k in range(0,7): # Iterate through data types in dataset
            worksheet.write_column(1,k+1,dataset[k]) # Write data
    makeNormalCharts(xlsx_file, worksheets, len(processed_data[0][0]))
    makeLinearCharts(xlsx_file, worksheets, len(processed_data[0][0]))
    xlsx_file.close() # For now, to save
    return

fillSheets()
