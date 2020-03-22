
/**************************************************************************/
/*
This code was adapted from the example code for the Adafruit mcp9808
temperature sensor library. It recieves temperature readings from the sensors
and exports them to a connected computer via the COM port.
Modified by: Justin Lawrence
Date Last Modified: 24/1/2020
*/
/**************************************************************************/

#include <Wire.h>
#include "Adafruit_MCP9808.h"

// Create the MCP9808 temperature sensor objects
Adafruit_MCP9808 tempsensors[8] = {Adafruit_MCP9808(),Adafruit_MCP9808(),Adafruit_MCP9808(),Adafruit_MCP9808(),Adafruit_MCP9808(),Adafruit_MCP9808(),Adafruit_MCP9808(),Adafruit_MCP9808()};
unsigned long timeBase; // Base time for comparison

void setup() {
  Serial.begin(9600);
  while (!Serial); //waits for serial terminal to be open
  
  // Make sure the sensor is found, you can also pass in a different i2c
  // address with tempsensor.begin(0x19) for example, also can be left in blank for default address use
  // Also there is a table with all addres possible for this sensor, you can connect multiple sensors
  // to the same i2c bus, just configure each sensor with a different address and define multiple objects for that
  //  A2 A1 A0 address
  //  0  0  0   0x18  this is the default address
  //  0  0  1   0x19
  //  0  1  0   0x1A
  //  0  1  1   0x1B
  //  1  0  0   0x1C
  //  1  0  1   0x1D
  //  1  1  0   0x1E
  //  1  1  1   0x1F

  // Initialize all sensors

  for (int i{0}; i < 8; i++) {
    if(!tempsensors[i].begin(0x18+i)) { // Initializes sensors
      Serial.print("Sensor "); Serial.print(i + 1); Serial.println(" failed to initialize");
      while (true);
    }
    tempsensors[i].setResolution(3); // sets the resolution mode of reading, the modes are defined in the table bellow:
  }
  
  // Mode Resolution SampleTime
  //  0    0.5°C       30 ms
  //  1    0.25°C      65 ms
  //  2    0.125°C     130 ms
  //  3    0.0625°C    250 ms

  for (byte i{0}; i < 8; i++) {
    tempsensors[i].wake(); // Wake up all sensors
  }
  delay(10000);
  timeBase = millis();
  
}

void loop() {

  Serial.print("<"); // Open reading
  Serial.print("{"); // Open datapoint
  Serial.print(millis()); // Write out time
  Serial.print("}"); // Close datapoint
  for (byte i{0}; i < 8 ; i++) { // Read and write out temperatures
    float temp = tempsensors[i].readTempC();
    Serial.print("{");
    Serial.print(temp);
    Serial.print("}");
  }
  Serial.print(">"); // Close reading
  delay(130); // 130ms delay for the next reading
}
