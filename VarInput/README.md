# VarInput.py

A tool to put together scenarios for historic simulation and calculate a variance covariance matrix from the scenarios, which can be used in analytic VAR calculations.

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

These settings can also be configured in a separate file called `VarInput.config` in the same folder, which is executed if it exists.

Furthermore 
- a historic market data file is required, which is essentially a collection of ORE market data for various historic dates.
- a fixing data file is needed.

If there are any missing market data leading to incomplete calculations of scenarios, the user is prompted to either remove these missing data points or cancel the calculation to repair the missing market data.

Module requirements: numpy, pandas, lxml and ORE (version 1.8.12.1)