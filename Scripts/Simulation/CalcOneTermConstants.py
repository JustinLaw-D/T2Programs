# -*- coding: utf-8 -*-
"""
Created on Thu Feb 20 15:50:00 2020

@author: justi
"""

import math

def findC(T_init, T_atm):
    current = T_init # Current variable
    current -= T_atm # Subtract atmospheric temperature
    current /= math.sin(3*math.pi/8) # Divide by sin(pi/8)
    return current # And you have C

def findAlpha(T_init, T_atm, T_final, dT, L):
    C = findC(T_init, T_atm) # Get C
    current = T_final # Current variable
    current -= T_atm # Subtract atmospheric pressure
    current /= C # Divide C
    current /= math.sin(3*math.pi/8) # Divide it by sin(pi/8)
    current = math.log(current) # Take natural log
    current *= -1 # Make positive
    current /= (math.pi)**2 # Divide by pi^2
    current *= L**2 # Multiply by L^2
    current /= dT # Divide by the time change
    return current # And you have alpha

def getFunction(T_INIT, T_ATM, T_FINAL, dT, L): # Get constants for single term function
    C = findC(T_INIT, T_ATM)
    alpha = findAlpha(T_INIT, T_ATM, T_FINAL, dT, L)
    return C, alpha
