#OLA Post Processing

##About
In order to accurately compare different test conditions, they must evaluated at similar conditions. Currently, the results include the portion of the test before
the actuator turns on and the part where the cleaning is done and the test taker has to shut off the camera. These data points are irrelevant in terms of cleaning. 
However, the energy consumption is calculated on a per time basis so energy consumption is not accurately calculated. This code looks for the time where the actuator turns 
on and determines the time it takes to get to 80% clean as determined by the OLA. Using this metric, the different conditions can be evaluated fairly for efficiency.
