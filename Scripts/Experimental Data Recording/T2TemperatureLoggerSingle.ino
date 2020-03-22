
/**************************************************************************/
/*!
This is a demo for the Adafruit MCP9808 breakout
----> http://www.adafruit.com/products/1782
Adafruit invests time and resources providing this open source code,
please support Adafruit and open-source hardware by purchasing
products from Adafruit!
*/
/**************************************************************************/
// Adapted from adafruit demo
// Records data from single temperature sensor

#include <Wire.h>
#include "Adafruit_MCP9808.h"

// Create the MCP9808 temperature sensor object
Adafruit_MCP9808 tempsensor = Adafruit_MCP9808();

void setup() {
  Serial.begin(9600);
  while (!Serial); //waits for serial terminal to be open, necessary in newer arduino boards.
  
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
  if (!tempsensor.begin(0x18)) {
    Serial.println("Temperature sensor could not be found!");
    while (1);
  }

  tempsensor.setResolution(3); // sets the resolution mode of reading, the modes are defined in the table bellow:
  tempsensor.wake(); // Wake the temperature sensor
  delay(4000); // Wait 4 seconds to begin measurements
  // Mode Resolution SampleTime
  //  0    0.5°C       30 ms
  //  1    0.25°C      65 ms
  //  2    0.125°C     130 ms
  //  3    0.0625°C    250 ms
}

void loop() {
  
  Serial.print("<"); // Open reading
  Serial.print("{"); // Open datapoint
  Serial.print(millis()); // Write out time
  Serial.print("}"); // Close datapoint
  // Read and write out temperature
  float temp = tempsensor.readTempC();
  Serial.print("{");
  Serial.print(temp);
  Serial.print("}");
  Serial.print(">"); // Close reading
  delay(250); // 130ms delay for the next reading
  
}
