# -*- coding: utf-8 -*-
"""
Created on Sat Feb 29 09:10:42 2020

@author: justi
"""

## This is a short fancy program for one function, and one function only
## This function calculates the temperature of the sensor at a later time based on the current time
## Be in awe of its greatness
## ***Right now I can only garuntee this works for when the sensor temperature is above the node temperature

from math import exp

alph = -1.6563*10**-2 # Constant for exponential decay

def sensorUpdate(T_sensor, T_node, dt):
    ## Takes the current temperature of the sensor and node, and the time step
    C = T_sensor - T_node # Get the constant C
    dT_next = C*exp(alph*dt) # Get the temperature difference for the next time
    T_next = dT_next + T_node # Add back offset to get sensor temperature
    return T_next  