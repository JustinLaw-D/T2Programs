#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jan 13 19:20:05 2020
@author: raja
"""

import numpy as np
import matplotlib.pyplot as plt
import csv
import CalcOneTermConstants as Ct
import SensorCorrection as sc

data_file_name = 'processedData15.csv'
model_file_name = 'modelData15.csv'
L = 0.3048 #lenth of the rod (in meters)
n = 56 #number of nodes
dx = L/n #Thickness of the wall divided by the number of nodes
#---------------------------------------------------------------------------------
T0 = 312.71 #initial temperature at one end point
T1s = 281.9 #left boundary condition
T2s = T1s #right boundary condition
Temp_dataFinal = 299.52 #this is the temp value used for the calculation of alpha
dataTimeChange = 217.09 #this is the time value used for the calculation of alpha
#The following values are what the data collected for the nodes
T1 = 305.4
T2 = 299.9
T3 = 294.34
T4 = 291.21
T5 = 288.78
T6 = 287.4

#alpha = Ct.findAlpha(T0, T1s, Temp_dataFinal, dataTimeChange, L) #this is the thermal difusivity defined as thermal conductivity / (density (rho) * Cp)
alpha = 11.72*10**-6
#---------------------------------------------------------------------------------
t_final = 100000 #final time keep as in integer for proper data collection
print(alpha)
alpha_air = .0000020

dt = .01 #the time intervals keep as an integer for proper data collection
showEvolve = False
x = np.linspace(dx/2, L-dx/2, n) # We will be plotting the temperature of space between each node
T = np.ones(n)*T0 #This is the starting temperature
dTdt = np.empty(n) #defining the temperature derivative

t = np.arange(0, t_final, dt) #time vector

'''These are the initial reading that we will determine from our tests
we can imput the 7 sensors readings as soon as the first one reaches
the desired temp (in this case 200 degrees C) and imput them here and 
see how it compares to the real deal!'''


#Sets up all the nodes here to have a good guess of the temperature
for i in range(0,7):
    T[i],T[i+7],T[i+14],T[i+21],T[i+28],T[i+35],T[i+42],T[i+49] = T0+(i*(T1-T0)/7),T1+(i*(T2-T1)/7),T2+(i*(T3-T2)/7),T3+(i*(T4-T3)/7),T4+(i*(T5-T4)/7),T5+(i*(T6-T5)/7),T6,T6

Ts_next = T.copy() # Setup temperature sensor data

#Lists that are used in csv files
time = []

temp0 = []
temp1 = []
temp2 = []
temp3 = []
temp4 = []
temp5 = []
temp6 = []
plotTime = []


data_times = list()
skip = 0
first_time = 0
with open(data_file_name, 'r', newline = '') as data_file:
    csvReader = csv.DictReader(data_file, fieldnames=["Time","TempOne","TempTwo","TempThree","TempFour","TempFive","TempSix","TempSeven","Atmospheric"])
    for row in csvReader:
        if skip == 2:
            first_time = float(row["Time"])
        if not (skip < 2):
            data_times.append((float(row["Time"]))-first_time)
        skip += 1

k = 0
for j in range(0, len(t)): #Goes through each time and calculates the temperature at that node
    Ts = Ts_next
    for i in range (1, n-1): #This loops through each spatial node and finds the derivative at that node
        dTdt[i] = alpha *(-(T[i]-T[i-1])/dx**2+(T[i+1]-T[i])/dx**2) #This is the approximation for every node that isnt on the boundary
            
    dTdt[0] = alpha_air *(-(T[0]-T1s)/dx**2+(T[1]-T[0])/dx**2) #Left boundary approximation T1s is the left boundary
    dTdt[n-1] = alpha_air *(-(T[n-1]-T[n-2])/dx**2+(T2s-T[n-1])/dx**2) #right boundary approximation T2s is the right boundary
    if (j != len(t) - 1): # At every time except the last one
        for i in range(0,n): # Loop through the spacial nodes
            Ts_next[i] = sc.sensorUpdate(Ts[i],T[i],dt) # Update temperature sensors
    T = T + dTdt*dt
    time.append(j*dt)

    if round(time[j], 2) == round(data_times[k], 2):
        temp0.append(Ts[0])
        temp1.append(Ts[7])
        temp2.append(Ts[14])
        temp3.append(Ts[21])
        temp4.append(Ts[28])
        temp5.append(Ts[35])
        temp6.append(Ts[42])
        plotTime.append(j*dt)
        k += 1
    
    if k == len(data_times):
        break

    if showEvolve == True:
        plt.clf()
        plt.ion()
        plt.figure(1)
        plt.plot(x, T)
        plt.axis([0, L, 0, 400])
        plt.xlabel('Distance (m)')
        plt.ylabel('Temperature (C)')
        #plt.legend(str(j*dt))
        plt.show()
        plt.pause(0.1) 


with open(model_file_name, 'w', newline='') as csv_file:
    csv_writer = csv.writer(csv_file)
    csv_writer.writerow(['time (s)', 'temp0', 'temp1', 'temp2', 'temp3', 'temp4', 'temp5','temp6', 'alpha value:', alpha])
    csv_writer.writerow([''])
    for i in range(0, len(temp0)):
        print(i)
        csv_writer.writerow([plotTime[i],temp0[i],temp1[i],temp2[i],temp3[i],temp4[i],temp5[i],temp6[i]])


if showEvolve == True:
    plt.figure(1)
    plt.plot(x, T)
    plt.axis([0, L, 0, 400])
    plt.xlabel('Distance (m)')
    plt.ylabel('Temperature (C)')
    #plt.legend(str(j*dt))
    plt.show()
    plt.pause(0.1)



    

