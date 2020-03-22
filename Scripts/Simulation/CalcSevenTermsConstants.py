# -*- coding: utf-8 -*-
"""
Created on Thu Feb 20 15:50:00 2020

@author: justi
"""

## A more complicated version of CalcOneTermConstants which attempts to use
## a seven term approximation. Currently it doesn't find the alpha value
## properly, but it could be built upon.

import math
import numpy as np

def constructMatrix():
    matrix = list() # Total matrix
    currentLinear = list() # Current row of matrix
    for i in range(1,8): # Iterate through for each temperature
        currentLinear.clear() # Reset list
        for k in range(1,8): # Iterate through each term
            currentLinear.append(math.sin((k*math.pi)*(i/8))) # (i)/8 = x/L, and k*math.pi is self-explanatory
        matrix.append(currentLinear.copy()) # Add to matrix
    print(matrix)
    return np.asarray(matrix) # Return the matrix as an array

def constructConstantsMatrix(temperatures):
    # Recieves a list of temperatures and times. Temperatures at first seven points are 0-6,
    # Atmospheric is 7, length of the rod is 8, final temperature 9, time to final temperature 10
    constants = list() # List of solutions (temperatures)
    atmospheric = temperatures[7] # Get the atmospheric temperature to subtract from solutions
    for i in range(0,7): # Iterate through the temperatures
        constants.append(temperatures[i] - atmospheric) # Add temperature to list
    return np.asarray(constants.copy()) # Convert to array and return

def findCVals(temperatures):
    matrix = constructMatrix()
    constants = constructConstantsMatrix(temperatures)
    c_values = np.linalg.solve(matrix,constants) # Find c solutions
    return c_values

def findAlpha(C, T_init, T_atm, T_final, dT, L):
    current = T_final # Current variable
    current -= T_atm # Subtract atmospheric pressure
    current /= C # Divide C
    current /= math.sin(math.pi/8) # Divide it by sin(pi/8)
    current = math.log(current) # Take natural log
    current *= -1 # Make positive
    current /= (math.pi)**2 # Divide by pi^2
    current *= L**2 # Multiply by L^2
    current /= dT # Divide by the time change
    return current # And you have alpha

def getFunction(temperatures, T_final, dt, L): # Get constants for a seven term function
    c_values = findCVals(temperatures)
    alpha = findAlpha(c_values[0], temperatures[0], temperatures[7], T_final, dt, L)
    return c_values, alpha

def T(c_values, alpha, x, t, L, T_atm): # Returns value of analytic function modelling temperature
    total = 0 # Total of all terms
    for i in range(1,8): # Iterate through terms
        term = 0 # Storage for current term
        term += c_values[i-1] * math.sin((i*math.pi)*(x/L)) # Sine and C parts
        term *= math.exp(-1*(alpha*(i**2)*(math.pi**2)*t)/(L**2)) # Exponential part
        total += term # Add to total
    total += T_atm # Add final constant term
    return total
