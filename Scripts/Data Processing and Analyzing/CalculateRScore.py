# -*- coding: utf-8 -*-
"""
Created on Sun Feb 23 19:53:45 2020

@author: justi
"""

def calculateMean(experimental_data): # Calculates the mean of the experimental data at node
    total = 0 # Temporary total variable 
    for k in range(0, len(experimental_data)): # Iterate through data points
        total += experimental_data[k] # Add ln(temperature)
    mean = total/len(experimental_data) # Compute mean
    return mean

def calcSumOfSquares(experimental_data, mean): # Calculates the sum of the variances
    sumSquares = 0 # Current sum of squares
    for k in range(0, len(experimental_data)): # Iterate through datapoints
        sumSquares += (experimental_data[k] - mean)**2 # Add variance to sum
    return sumSquares

def calcSumOfResiduals(residual_data): # Calculates the sum of the residuals squared
    sumResiduals = 0 # Residuals squared
    for k in range(0, len(residual_data)): # Iterate through datapoints
        sumResiduals += (residual_data[k])**2 # Add residual squared
    return sumResiduals

def calculateRScores(measured_data): # Calculates R score, is interacted with
    r_scores = list() # List of r scores
    for i in range(0,7): # Iterate through nodes    
        mean = calculateMean(measured_data[i][4]) # Get mean
        sum_squares = calcSumOfSquares(measured_data[i][4], mean) # Get sum of squares
        sum_residuals = calcSumOfResiduals(measured_data[i][6]) # Get sum of residuals
        R = 1 - (sum_residuals/sum_squares) # Calculate r_score for node
        r_scores.append(R) # Add to list
    return r_scores
