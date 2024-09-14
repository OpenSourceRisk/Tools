# VarInput.py

A tool to create a scenario file for historic simulation VAR and simultaneously calculate a variance covariance matrix from these scenarios, which can be used in analytic VAR calculations.

Customize input files to the needed ones:

`historicmarketdataFile = "Input/histmarketdata.txt"`  
`fixingdataFile = "Input/Fixingdata.txt"`  
`curveconfigFile = "Input/curveconfig.xml"`  
`conventionsFile = "Input/conventions.xml"`  
`pricingengineFile = "Input/pricingengine.xml"`  
`todaysmarketFile = "Input/todaysmarket.xml"`  
`simulationFile = "Input/simulation.xml"`  
`sensitivityFile = "Input/sensitivity.xml"`  

And also the output files can be customized:  

`outputcovarianceFile = 'covariance.csv'`  
`outputscenariosFile = 'scenarios.csv'`  

These settings can also be configured in a separate file called `VarInput.config` in the same folder, which is compiled/executed if it exists.

Apart from the ORE config files (curveconfig, conventions, pricingengine, todaysmarket, simulation and sensitivity)  
- a historic market data file is required, which is essentially a concatenation of several ORE market data files for the required historic dates.
- a fixing data file is needed.

The sensitivity configuration is actually not required for the scenario report, it is used by the tool to derive the required differential calculation method for the input of the variance covariance calculation.

The results of the first historic data are regarded as a reference for the required columns (I don't check the simulation.xml), 
if there are any missing subsequent market data, the results of the scenario report are discarded with a warning.

Module requirements: numpy, pandas, lxml and ORE (tested with version 1.8.12.1)