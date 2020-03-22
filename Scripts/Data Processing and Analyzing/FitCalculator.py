# -*- coding: utf-8 -*-
"""
Created on Sun Feb  2 21:49:30 2020

@author: justi
"""

## CANNOT FUNCTION INDEPENDENTLY, TO BE CALLED FROM DATAFITEXCEL
u_Temp = 0.125 # Temperature uncertainty
P = 2 # Number of perameters

## This part of the program calculates the chi^2 score of the data

def calcResiduals(processed_data):
    sum_residuals = [0,0,0,0,0,0,0] # Sum of all residuals squared
    num = len(processed_data[0][0]) # Number of data points
    for i in range(0,7): # Iterate through nodes
        dataset = processed_data[i] # Get data for that node
        for k in range(0, num): # Iterate through datapoints
            data = dataset[0][k] # Get data and model values
            model = dataset[1][k]
            residual = (data - model) # Calculate residual
            sum_residuals[i] += (residual/u_Temp)**2 # Add to total residuals after processing
    return sum_residuals, num

def calculateChiScore(processed_data):
    sumResiduals, num = calcResiduals(processed_data)
    chiScores = list() # List of x^2-scores
    for i in range(0,7): # Iterate through locations
        chiScores.append(sumResiduals[i]/(num-P)) # Calculate and add x^2 scores
    return chiScores
        