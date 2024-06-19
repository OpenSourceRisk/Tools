import numpy as np
import pandas as pd
import os
import ORE as ore
from lxml import etree
from warnings import simplefilter
simplefilter(action="ignore", category=pd.errors.PerformanceWarning)

# input files
historicmarketdataFile = "Input/histmarketdata.txt"
fixingdataFile = "Input/Fixingdata.txt"
curveconfigFile = "Input/curveconfig.xml"
conventionsFile = "Input/conventions.xml"
pricingengineFile = "Input/pricingengine.xml"
todaysmarketFile = "Input/todaysmarket.xml"
simulationFile = "Input/simulation.xml"
sensitivityFile = "Input/sensitivity.xml"
# output files
outputcovarianceFile = 'covariance.csv'
outputscenariosFile = 'scenarios.csv'

def decodeXML(filename):
	if not os.path.isfile(filename):
		print("no file " + filename + " found!")
		os._exit(1)
	return etree.tostring(etree.parse(filename)).decode('UTF-8')

def check(app):
	errors = app.getErrors()
	time = app.getRunTime()
	print ("%.2f sec, %d errors" % (time, len(errors)))       
	if len(errors) > 0:
		for e in errors:
			print(e)
		return False
	else:
		return True

def ql_period_in_years(periodStr):
	period = ore.Period(periodStr)
	units = period.units()
	if units == ore.Days:
		denominator = 365
	elif units == ore.Weeks:
		denominator = 52
	elif units == ore.Months:
		denominator = 12
	elif units == ore.Years:
		denominator = 1
	else:
		print("unhandled period unit " + str(units))
		os._exit(1)
	return period.length() / denominator

# prepare historic market data for iterative scenario analytic
if not os.path.isfile(historicmarketdataFile):
	print("no file " + historicmarketdataFile + " found!")
	os._exit(1)
df = pd.read_csv(historicmarketdataFile,sep='\t')
df["wholeLine"] = df["Date"].astype(str)+"\t"+df["Name"]+"\t"+df["Value"].astype(str)
df["Date"] = pd.to_datetime(df["Date"],format="%Y%m%d")

# set up ORE for running scenario analytic
inputs = ore.InputParameters()
inputs.setResultsPath(".")
inputs.setAllFixings(True)
inputs.setEntireMarket(True)
inputs.setCurveConfigs(decodeXML(curveconfigFile))
inputs.setConventions(decodeXML(conventionsFile))
inputs.setPricingEngine(decodeXML(pricingengineFile))
inputs.setTodaysMarketParams(decodeXML(todaysmarketFile))
inputs.insertAnalytic("SCENARIO")	
inputs.setScenarioSimMarketParams(decodeXML(simulationFile))

# get YieldCurve Tenors from simulation.xml for converting discount factors and fxspot rates
simmarketDef = etree.parse(simulationFile) 
tenors = simmarketDef.find("./Market/YieldCurves/Configuration/Tenors").text.split(",")
tenorsYrs = [ ql_period_in_years(period) for period in tenors]

# get information from sensitivity.xml for shifting 
if not os.path.isfile(sensitivityFile):
	print("no file " + sensitivityFile + " found!")
	os._exit(1)
sensitivityDef = etree.parse(sensitivityFile)
values = {}
for el in sensitivityDef.iterfind("./"):
	if el.tag == "CrossGammaFilter":
		continue
	for subEl in el.iterfind("./"):
		# need to change for FXSpot (from scenario report) <> FxSpot (from sensitivity.xml)
		startPattern = (subEl.tag if subEl.tag != "FxSpot" else "FXSpot") + "/" + subEl.attrib.values()[0]
		shiftSize = float(subEl.find("./ShiftSize").text)
		shiftType = subEl.find("./ShiftType").text
		if shiftType in values:
			if shiftSize in values[shiftType]:
				values[shiftType][shiftSize].append(startPattern)
			else:
				values[shiftType][shiftSize] = []
		else:
			values[shiftType] = {}
			values[shiftType][shiftSize] = []
			values[shiftType][shiftSize].append(startPattern)

if not os.path.isfile(fixingdataFile):
	print("no file " + fixingdataFile + " found!")
	os._exit(1)
with open(fixingdataFile) as f:
    fixingsdata = ore.StrVector(f.read().splitlines())

# create scenarios file as well
if os.path.isfile(outputscenariosFile):
	os.remove(outputscenariosFile)
file=open(outputscenariosFile, 'w')

headerRow = ""
dfriskfact = pd.DataFrame()
for scenDate in df["Date"].unique():
	# get marketdata block from history
	marketdata = df[df["Date"] == scenDate]["wholeLine"].tolist()
	inputs.setAsOfDate(scenDate.strftime("%Y-%m-%d"))
	
	# run scenario report
	oreapp = ore.OREApp(inputs, "log.txt", 63, True)
	oreapp.run(marketdata,fixingsdata)
	if not check(oreapp):
		os._exit(1)
	report = oreapp.getReport("scenario")
	
	# write historic scenarios, first create headerRow only once ...
	if headerRow == "":
		for i in range(report.columns()):
			headerRow += report.header(i)+("\t" if i<report.columns()-1 else "")
		file.write(headerRow + "\n")
	# ... and write data row for history date
	dataRow = report.dataAsString(0)[0]+"\t"+str(report.dataAsSize(1)[0])+"\t"
	for i in range(2,report.columns()):
		dataRow += str(report.dataAsReal(i)[0])+("\t" if i<report.columns()-1 else "")
	file.write(dataRow + "\n")

	# accumulate history for covariance calculation, converting curves to zero rates
	for i in range(3,report.columns()):
		riskfactor = report.header(i)
		riskfactorParts = riskfactor.split("/")
		if riskfactorParts[0] == "DiscountCurve" or riskfactorParts[0] == "IndexCurve":
			# convert curve discountfactor to zero rate before
			dfriskfact.at[scenDate,riskfactor] = -np.log(float(report.dataAsReal(i)[0]))/tenorsYrs[int(riskfactorParts[2])]
		else:
			# use riskfactor value directly
			dfriskfact.at[scenDate,riskfactor] = report.dataAsReal(i)[0]
file.close()

# dfriskfact.to_pickle("riskfactors")
# dfriskfact = pd.read_pickle("riskfactors")

# transpose the risk factors so the dates are horizontal and the risk factors vertical (for np.cov)
dfriskfactT = dfriskfact.T

# calculate differencials according to shiftType and shiftSize for each riskfactor
convLog = ""
for shiftType in ['Absolute','Relative']:
	values[shiftType]["data"] = pd.DataFrame()
	for shiftSize in values[shiftType].keys():
		if shiftSize == "data":
			continue
		for startPattern in values[shiftType][shiftSize]:
			dfPart = dfriskfactT[dfriskfactT.index.str.startswith(startPattern)]
			if dfPart.empty:
				continue
			dfPartDelayed = dfPart.shift(1, axis=1)
			if (shiftType == 'Absolute'):
				resDf = (dfPartDelayed - dfPart) * (1/shiftSize)
			else:
				resDf = (dfPartDelayed / dfPart - 1) * (1/shiftSize)
			if values[shiftType]["data"].empty:
				values[shiftType]["data"] = resDf
			else:
				values[shiftType]["data"] = pd.concat([values[shiftType]["data"], resDf], axis=0)

# put all differencials together, remove missing data and calculate the variance covariance matrix
finalDf = pd.concat([values['Absolute']["data"], values['Relative']["data"]], axis=0)
finalDf.drop(columns=finalDf.columns[0], axis=1, inplace=True) # remove first column -> difference to delayed
if (finalDf.isna().any().sum() > 0):
	print("following dates have missing data:")
	print(finalDf.loc[finalDf.isna().any(axis=1),finalDf.isnull().any()])
	ret = input("should these be removed (maybe business holiday) or [c]ancel to repair input data?")
	if ret == 'c':
		os._exit(1)
	finalDf.dropna(axis=1,inplace=True)
varcovar = np.cov(finalDf)

# output covariance file in ORE format
if os.path.isfile(outputcovarianceFile):
	os.remove(outputcovarianceFile)
file=open(outputcovarianceFile, 'w')
for x in range(0, varcovar.shape[0]):
	for y in range(x, varcovar.shape[1]):
		vcov = varcovar[x,y]
		file.write(str(finalDf.index[x]) + "\t" + str(finalDf.index[y]) + "\t%.2f" % vcov + "\n")
