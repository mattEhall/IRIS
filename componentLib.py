# Model Component Classes for IRIS - Island Renewables Integration Simulator
# (c) Copyright Matt Hall and Matthew McCarville, 2019
# To be cleaned up and released under BSD license.

import matplotlib.pyplot as plt
import numpy as np
import xlwings as xw
import datetime

import matplotlib as mpl
mpl.rcParams['lines.linewidth'] = 1

# regular expressions used
#wb.sheets\("([^"]+)"\).range\(("[^"]+")\).value
#wb.sheets\[([^\]]+)\].range\(("[^"]+")\).value
#wb.sheets\("([^"]+)"\).range\((\([^\)]+\))\).value
#wb.sheets\(([^\)]+)\).range\((\([^\)]+\))\).value
#getCellVal\(wb, "\1", \2\)
# good refences: https://www.rexegg.com/regex-quickstart.html
#                https://regexr.com



# a temporary function (to be moved) for convenience in loading columns of data with xlwings and returning as a np.array
def loadxlcol(wb, sheetName, topCellOfColumn):

	#print("  loadxlcol loading "+str(sheetName)+str(topCellOfColumn))

	try:
		# note: workbookObject must be an xlwings workbook object
		data = wb.sheets[sheetName].range(wb.sheets[sheetName][topCellOfColumn], 
													  wb.sheets[sheetName][topCellOfColumn].end("down")).options(np.array).value 

		return data #np.array(data, dtype=np.float)

	except:
		print("Error while trying to load column of data at "+str(sheetName)+str(topCellOfColumn))
		return []
		
	# TODO: <<<<<<<<<<< figure out number conversion

# a function to get the value of a cell, including built-in error checking and handling
def getCellVal(wb, sheet, range, type=np.float, condition=lambda x:x!=None, errormsg="", errorval=0):
	# sheet: reference to the sheet of the Excel file
	# range: parameter to be passed to xlwings range function - e.g. (1,1) or "A1"
	# for type provide function to make a type
	# for condition, set a lambda function that will return true if the value is valid, e.g. condition=lambda x: x > 0 and x <10
	# for errormsg, provide a string of what the error should read if condition isn't satisfied - an exception will only be thrown if this is not empty 
	# for errorval, what value to return if condition isn't satisfied
	
	value =  wb.sheets(sheet).range(range).value
	
	if condition(value):       # if condition is satisfied
		return type(value)     # make value of desired type, and return
	else:
		if errormsg == "":     # if no error text provided, don't throw an exception, just return the error value
			return errorval
		else:
			notice = "Error with sheet '"+sheet.name+"' cell "+str(range)+": "+errormsg
			raise NameError(notice)
			return errorval
			

# a function for getting the Range reference to a time series or dataset header, but checking for errors
def getCellRef(range, optional=False):

	if range.formula == "":
		if optional:
			return None
		else:
			raise Exception("Error: cell "+str(range.sheet)+str(range.address)+" is empty. It should point to a time series header.")
	
	else:
		try: 
			if "!" in range.formula:    # if it appears to reference a different sheet, things are simple
				ref = xw.Range(range.formula)
			else:                       # otherwise, need to ensure xlwings looks on the same sheet rather than the active sheet
				#print(range.formula)
				#print(range.sheet)
				#print(range.sheet.range("A1"))
				#print(range.sheet.range(range.formula))
				ref = range.sheet.range(range.formula)
			#print("getCellRef got destination of "+str(ref)+" from "+str(range.formula))
			return ref
		except:
			if optional:
				return None
			else:
				raise Exception("Error getting cell reference at cell "+str(range.sheet)+str(range.address)
			                +". Check that this cell's formula points to the title of a data set")
			
	
## scale and offset a time series to meet total energy and peak specs (useful for making generic data inputs)
def scaleAndOffset(data, dt=1, energy=0, peak=1):

	dataMax = np.max(data)
	dataMean = np.mean(data)
	dataN = len(data)
	if energy==0:
		return data*peak/dataMax
	else:
		# calculate coefficients for linear transformation y=mx+b where 
		# y is output power signal and x is input power signal
		
		# m = 
		m = (peak - energy/(dataN*dt))/(dataMax - dataMean)  # slope of linear transformation
		
		# b = y2 - m x2 
		b = peak - m*dataMax # offset of linear transformation
	
		return m*data + b
		
		



# load and interpolate TS data of any length - cellref is the header of the main time series column (the y data)
def getTS(cellref):
	
	row = cellref.row
	col = cellref.column
	sht = cellref.sheet
	wb = cellref.sheet.book
		
	# the general info
	name  = getCellVal(wb, sht, (row,col)  , type=str)
	units = getCellVal(wb, sht, (row+1,col), type=str, errorval="")
	TStype  = getCellVal(wb, sht, (row+2,col), type=str, errormsg="Type entry of column '"+name+"' must be chosen from the list.")	
	xdataref = getCellRef(wb.sheets(sht).range((row+3,col)), optional=True)
	
	# the y data	
	ydata = loadxlcol(wb, sht, wb.sheets(sht).range((row+5,col)).address)

	# set up datetime vector (this is currently hard-coded and redundant)
	DT0 = datetime.datetime(2020,1,1)                              # starting point of simulation
	DTs = [DT0 + datetime.timedelta(hours=i) for i in range(8760)]  # set up datetime array for each hour in 1 year

	hrs = np.arange(8760)                                           # time series in hours
	#ts = th*60*60                                                  # time series in seconds

	# get the x data if applicable	
	# note: if the time column is provided numerically, it's interpreted as hours, and mismatches will be stretched to fit 8760
	#       if it's provided as dates and times, these are used (currently adjusted to start from the first midnight, and tiled to fit)
	if xdataref != None:
		xrow = xdataref.row
		xcol = xdataref.column
		xsht = xdataref.sheet
		
		#print("xsht is "+str(xsht))
		
		xname = getCellVal(wb, xsht, (xrow,xcol)  , type=str)
		xunits= getCellVal(wb, xsht, (xrow+1,xcol), type=str)   # <<< do anything with this?
		xdata = loadxlcol(wb, xsht, wb.sheets(xsht).range((xrow+5,xcol)).address)
		
		if len(xdata) != len(ydata):
			raise ValueError("In sheet '"+sht.name+"': Length of time column '"+xname+"' is not the same as length of data column '"+name+"'.")
		
		#TODO: check for units of x data
		
		# if datetime format, do some conversions to get an hours vector
		if type(xdata[0]) == datetime.datetime:   
			
			xdelta = xdata - DT0    # get timedeltas of input times, relative to sim start time
			xdata = np.array([i.total_seconds() for i in xdelta])/60.0/60.0  # convert to hours, relative to sim start time
		# otherwise we assume array is already in hours
			
		#else:                                      # if numerical, treat as hours
		#	ydata2 = np.interp( np.arange(xdata[0], xdata[-1]+0.001, 1.0), xdata, ydata)
		
	# if no x data provided, the inputs should be stretched over the applicable period...
	else:
		xdata = np.linspace(0, 1, len(ydata)+1)[:-1]      # this is assuming the "stretch" mode rather than tiling


	
	# process the y data to form a year-long time series
	if TStype == "year":
	
		n = 8760
	
		if xdataref==None:                                      # no time data provided case
			ys = np.interp( hrs, xdata*n, ydata)                # stretch over the year
		else:
			#if xdata[-1]-xdata[0] > n:                         # ensure more than the required time isn't provided
			#	raise ValueError("In sheet '"+sht.name+"': Times provided in column '"+xname+"' cover too large a span (>"+n" hours).")
			# could also tile here if not enough provided...    ydata = np.tile(ydata, np.int(8760/len(y_an))+1) ???
	
			if type(xdata[0]) == datetime.datetime:             # datetime provided case
				ys = np.interp( hrs, xdata, ydata, period=n)    # apply by date in the year, wrapping if needed
				
			else:                                               # hourly data provided case
				ys = np.interp( hrs, xdata, ydata, period=n)    # apply by date in the year, wrapping if needed
			
	elif TStype == "week":
	
		n = 168
	
		if xdataref==None:                                      # no time data provided case
			ys = np.interp( hrs[:n], xdata*n, ydata)                # stretch over the week
		else:
			if type(xdata[0]) == datetime.datetime:             # datetime provided case
				ys = np.interp( hrs[:n], xdata, ydata, period=n)    # apply by date in the week, wrapping if needed
				
			else:                                               # hourly data provided case
				ys = np.interp( hrs[:n], xdata, ydata, period=n)    # apply by date in the week, wrapping if needed

		# tile to fill up the year
		ys = np.tile(ys, int(8760/n)+1)[:8760]
		
	elif TStype == "day":
	
		n = 24
	
		if xdataref==None:                                      # no time data provided case
			ys = np.interp( hrs[:n], xdata*n, ydata)                # stretch over the day
		else:
			if type(xdata[0]) == datetime.datetime:             # datetime provided case
				ys = np.interp( hrs[:n], xdata, ydata, period=n)    # apply by hour in the day, wrapping if needed
				
			else:                                               # hourly data provided case
				ys = np.interp( hrs[:n], xdata, ydata, period=n)    # apply by hour in the day, wrapping if needed

		# tile to fill up the year
		ys = np.tile(ys, 365)[:8760]
		
	
	# if this is a weekday-weekend scenario, also load the weekday/weekendday data
	elif TStype == "weekday (wknd->)":
		
		# additional ydata for weekend day
		name2    = getCellVal(wb, sht, (row  ,col+1)  , type=str)
		units2   = getCellVal(wb, sht, (row+1,col+1), type=str, errorval="")
		TStype2  = getCellVal(wb, sht, (row+2,col+1), type=str)	
		ydata2   = loadxlcol( wb, sht, wb.sheets(sht).range((row+5,col+1)).address)
		
		if TStype2 != "weekend day":
			raise ValueError("In sheet '"+sht.name+"': Expecting weekday column '"+name+"' to be followed by a 'weekend day' type column.")
		if len(ydata2) != len(ydata):
			raise ValueError("In sheet '"+sht.name+"': Length of weekend day column '"+name2+"' is not the same as length of weekday column '"+name+"'.")
		
		n = 24
	
		if xdataref==None:                                      # no time data provided case
			ys1= np.interp( hrs[:n], xdata*n, ydata)                # stretch over the day
			ys2= np.interp( hrs[:n], xdata*n, ydata2)
		else:
			if type(xdata[0]) == datetime.datetime:             # datetime provided case
				ys1= np.interp( hrs[:n], xdata, ydata, period=n)    # apply by hour in the day, wrapping if needed
				ys2= np.interp( hrs[:n], xdata, ydata2, period=n)
				
			else:                                               # hourly data provided case
				ys1= np.interp( hrs[:n], xdata, ydata, period=n)    # apply by hour in the day, wrapping if needed
				ys2= np.interp( hrs[:n], xdata, ydata2, period=n)

		# tile to fill up the year (5 weekdays and 2 weekend days)
		ys = np.tile(np.hstack([ys1, ys1, ys1, ys1, ys1, ys2, ys2]),53)[:8760]
			
			
	else:
		raise ValueError("In sheet '"+sht.name+"': Type entry for column '"+name+"' is not from the list of supported options.")
		
		#print("Error: no time series provided from sheet "+str(sht)+":"+str(row)+" "+str(col))
		#return -1
		
	return ys, name, units
	
	

# load and interpolate BEV arrival/departure/SOC TS data of any length
def getBEVTS(wb, cellref):
	
	row = cellref.row
	col = cellref.column
	sht = cellref.sheet
	
	name  = getCellVal(wb, sht, (row,col)  , type=str)
	units = getCellVal(wb, sht, (row+1,col), type=str) 
	
	# the y data	(four adjacent columns)
	ydata1 = loadxlcol(wb, sht, wb.sheets(sht).range((row+4,col  )).address)
	ydata2 = loadxlcol(wb, sht, wb.sheets(sht).range((row+4,col+1)).address)
	ydata3 = loadxlcol(wb, sht, wb.sheets(sht).range((row+4,col+2)).address)
	ydata4 = loadxlcol(wb, sht, wb.sheets(sht).range((row+4,col+3)).address)
	
	
	#ydata = wb.sheets[sheetName].range(wb.sheets[sheetName].range((row,col)), 
	#                                              wb.sheets[sheetName].range((row,col+3)).end("down")).options(np.array).value

	#print(ydata1)
	#print(ydata2)
	#print(ydata3)
	#print(ydata4)
	

	# get the x data if applicable
	xdataref = getCellRef(wb.sheets(sht).range((row+2,col)), optional=True)
	
	if xdataref != None:
		xrow = xdataref.row
		xcol = xdataref.column
		xsht = xdataref.sheet
		
		#print("xsht is "+str(xsht))
		
		xname = getCellVal(wb, xsht, (xrow,xcol)  , type=str)
		xunits= getCellVal(wb, xsht, (xrow+1,xcol), type=str)
		xdata = loadxlcol(wb, xsht, wb.sheets(xsht).range((xrow+4,xcol)).address)
		
		#print(xdata)
	
		#TODO: check for units of x data
		
		# if xdata is provided, interpolate to hourly intervals (for now)
		ydata1 = np.interp( np.arange(xdata[0], xdata[-1]+0.001, 1.0), xdata, ydata1)
		ydata2 = np.interp( np.arange(xdata[0], xdata[-1]+0.001, 1.0), xdata, ydata2)
		ydata3 = np.interp( np.arange(xdata[0], xdata[-1]+0.001, 1.0), xdata, ydata3)
		ydata4 = np.interp( np.arange(xdata[0], xdata[-1]+0.001, 1.0), xdata, ydata4)
	
	ydata = np.vstack([ydata1, ydata2, ydata3, ydata4])
	
	return ydata, name, units




def gatherTS(wb, cellref):
	'''# gets/creates annaul TS data from annual,weekly, or daily inputs'''
	
	row = cellref.row
	col = cellref.column
	sht = cellref.sheet
		
	form_an = getCellRef(wb.sheets(sht).range((row  , col)), optional=True)
	form_wk = getCellRef(wb.sheets(sht).range((row+1, col)), optional=True)
	form_wd = getCellRef(wb.sheets(sht).range((row+2, col)), optional=True)
	form_we = getCellRef(wb.sheets(sht).range((row+3, col)), optional=True)
	
	# try annual time series
	if form_an != None:
		#print("line 190 about to getTS from")
		#print(form_an)
		y_an, name, units = getTS(wb, form_an)
		
		if len(y_an) < 8760:
			y_an = np.tile(y_an, np.int(8760/len(y_an))+1)
			
	# otherwise try weekly time series
	elif form_wk != None:
		y_wk, name, units = getTS(wb, form_wk)
		
		#TODO: ensure time series is one week long exactly!
		
		y_an = np.tile(y_wk, np.int(8760/len(y_wk))+1)
		
	# otherwise try daily time series
	elif form_wd != None:
		y_wd, name, units = getTS(wb, form_wd) # weekday
		
		if form_we != None:  # if weekends are different
			y_we, name, units = getTS(wb, form_we) # weekend day
			y_an = np.tile(np.hstack([y_wd, y_wd, y_wd, y_wd, y_wd, y_we, y_we]), np.int(365/7)+1)  # make the year
			
			
		else:   # one daily pattern only case (no weekend distinction) 
			y_an = np.tile(y_wd, 365)
	else:
		print("Error: no time series provided from sheet "+str(sht)+":"+str(row)+" "+str(col))
		return -1
		
	return y_an[:8760], name, units
		

# gets performance curve from cell reference to y data header
def gatherCurve(wb, cellref):
	
	row = cellref.row
	col = cellref.column
	sht = cellref.sheet
	
	name  = getCellVal(wb, sht, (row,col)  , type=str)
	units = getCellVal(wb, sht, (row+1,col), type=str) 
	
	print("Getting curve "+str(name))
	
	# the y data	
	ydata = loadxlcol(wb, sht, wb.sheets(sht).range((row+4,col)).address)

	# get the x data if applicable
	xdataref = getCellRef(wb.sheets(sht).range((row+2,col)), optional=True)
	
	if xdataref != None:
		xrow = xdataref.row
		xcol = xdataref.column
		xsht = xdataref.sheet
		
		xname = getCellVal(wb, xsht, (xrow  ,xcol), type=str)
		xunits= getCellVal(wb, xsht, (xrow+1,xcol), type=str)
		xdata = loadxlcol(wb, xsht, wb.sheets(xsht).range((xrow+4,xcol)).address)
		
	else:
		print("error: xdata must be specified for "+name)
		return

	return ydata, xdata, name, xname, xunits


	
# ================ end of functions. Classes start here. =======================



class load:
	'''Flexible load class 1 - this provides a generic model of flexible (shiftable) loads'''
	
	def __init__(self):
		
		self.name = "generic flexible load"
		
		self.loadTS         = []  # total load time series, numpy array [MW]
		self.adjustedLoadTS = []  # total load time series after load shifting/flexing, numpy array [MW]
		
		self.flexLoadTS = []  # time series of flexible portion of loadTS [MW] 
		
		self.cap_energy      = 0.0  # [MWh] energy capacity   <<<<< this isn't actually needed...?
		
		# the below can be populated with numpy arrays of how the available capacity varies over time
		self.availability_time     = [] # time instances (in hours)
		self.availability_fraction = [] # available capacity (relative to full capacity, ranging from 0 to 1)
		
		
		self.lastTime = 0.0
		
		self.flexLoadMod = "none"
		
		# flexload type 2 quantities
		self.time_shift_fwd = [] # this is the maximum number of hours in one direction (+/- extents) the load at one instant can be spread out. 
		self.time_shift_bck = []		
	#	self.loadShiftMode = 1   # <<<<<<<<<<<<<<<
		#I.e. a value of 4 will yield a total spread of 9 hours. (or 8 hours? need to compensate for other dts still? <<<)
		
		# flexload type 1 quantities
		self.cap_discharge   = 0.0  # [MW] max discharge rate
		self.self_discharge  = 0.0  # [%/dikay] self discharge rate in % of full capacity per day
		self.eta_charge      = 1.0  # [%] efficiency of charging
		self.eta_discharge   = 1.0  # [%] efficiency of discharging
		self.SOC             = 0.0  # [] state of charge
		self.SOCTS           = np.zeros(8760)  # state of charge time series (length hardcoded for now)
		
	
	# read in required data from excel and create the load 
	def create(self, wb, col, preview=0):
		
		loadcellref = wb.sheets("load").range((4,col))  # the range of the title of the time series

		self.name = loadcellref.value
		
		print("--------------------------")
		print("Creating load "+self.name)

		# get initial time series
		inputTS, inputTSname, inputTSunits = getTS(getCellRef(wb.sheets("load").range((6,col))))   # starting time series

		# apply seasonal scaling (multiply)
		form_scale = getCellRef(wb.sheets("load").range((7,col)), optional=True)
		if form_scale != None:
			scaleTS, scaleTSname, scaleTSunits = getTS(form_scale)
			inputTS = inputTS*scaleTS 
			
		# apply seasonal shifting (add)
		form_shift = getCellRef(wb.sheets("load").range((8,col)), optional=True)
		if form_shift != None:
			shiftTS, shiftTSname, shiftTSunits = getTS(form_shift)
			inputTS = inputTS + shiftTS		
		
		# apply floor of zero - we don't want to deal with negative loads here <<< should make a warning or something
		np.clip(inputTS, 0.0, None, out=inputTS)

		# apply performance curve conversion if applicable
		form_perf = getCellRef(wb.sheets("load").range((10,col)), optional=True)
		if form_perf != None:
			ycurve, xcurve, curvename, _, _ = gatherCurve(wb, form_perf)  # get the performance curve x and y data
						
			xdata, _, _ = getTS(getCellRef(wb.sheets("load").range((11,col))))  # this is the x(t) data used to feed the performance curve
					
			perf_scaler = np.interp(xdata, xcurve, ycurve)
			
			loadTS = inputTS/perf_scaler   # scale load by INVERSE of performance curve! (opposite of generation -- this treats it like a COP number eg. for heat pumps)
		
			meanPerf = np.sum(perf_scaler*loadTS)/np.sum(loadTS)
		
		else:
			loadTS = inputTS
			
			
		# load shaping/scaling
		energy = getCellVal(wb, "load", (13,col), errorval=None)
		peakP  = getCellVal(wb, "load", (14,col), errorval=None)
		if (energy != None) and (peakP != None):
			loadTS = scaleAndOffset(loadTS, energy=energy, peak=peakP)
		elif energy != None:
			loadTS = scaleAndOffset(loadTS, energy=energy)
		elif peakP != None:
			loadTS = scaleAndOffset(loadTS, peak=peakP)
				
		self.loadTS = loadTS
		self.adjustedLoadTS = np.array(self.loadTS) # make a copy of the original load for future record keeping
		
			
		# ---- demand response parameters ----
		
		self.flexLoadMod = getCellVal(wb, "load", (16,col), type=str, errorval="none")  # string describing flexible load type

		if self.flexLoadMod != "none":                                    # if flexible load, read in the magnitude inputs
		
			flexLoadFixed    = getCellVal(wb, "load", (17,col), errorval=0)   # absolute quantity of flexible load		
			flexLoadFraction = getCellVal(wb, "load", (18,col))               # dynamic fractional quantity of flexible load (% of fixed load)   
			if flexLoadFraction == None:
				flexLoad = flexLoadFixed + np.zeros(8760)
			else:
				flexLoad = flexLoadFixed + flexLoadFraction*loadTS
				
			form_flexLoadTS  = getCellRef(wb.sheets("load").range((19,col)), optional=True)
			if form_flexLoadTS != None:
				flexLoadVariableTS, _, _ = getTS(form_flexLoadTS)   # the option to read in a load avialability time series
				flexLoad = flexLoad + flexLoadVariableTS                # gets added to flexible load portion if some already exists
				
			self.flexLoadTS   = flexLoad                                  # save the final flexible load time series
			
		
		if self.flexLoadMod == "storage based":
			self.cap_energy    = getCellVal(wb, "load", (21,col))
			self.cap_charge    = getCellVal(wb, "load", (22,col))
			self.cap_discharge = getCellVal(wb, "load", (23,col))
			self.self_discharge= getCellVal(wb, "load", (24,col))
			self.eta_charge    = getCellVal(wb, "load", (25,col))
			self.eta_discharge = getCellVal(wb, "load", (26,col))
			
		elif "shift" in self.flexLoadMod:		
		#	self.loadShiftMode  = getCellVal(wb, "load", (36,col))
			self.time_shift_fwd = getCellVal(wb, "load", (28,col))
			self.time_shift_bck = getCellVal(wb, "load", (29,col))
				

		if preview > 0:  
		
			# TODO: demo currently only works for peak shaving - should add note or adjust to clarify
		
			ts= np.arange(len(inputTS))
		
			if "shift" in self.flexLoadMod:
				demosupply = np.mean(self.loadTS)    # for illustration, assume supply is constant, equal to mean load
				demo_adjustedLoadTS = self.applyLoadShift(demosupply, self.loadTS, preview=1)
			
			if form_perf != None:      # more detailed plots if a performance curve is used
				
				fig, ax = plt.subplots(3,1,sharex=True)
				ax[0].step(ts, inputTS)
				ax[0].set_ylabel(inputTSname+" ("+inputTSunits+")")
				ax[1].step(ts, perf_scaler)
				ax[1].axhline(meanPerf, label="mean performance value = {:5.2f}".format(meanPerf), dashes=(1,1))
				ax[1].legend()
				ax[1].set_ylabel(curvename)  #("+curveYunits+"/"+curveXunits+")"
				ax[2].step(ts, self.loadTS, label = "original load")
				if "shift" in self.flexLoadMod:
					ax[2].step(ts, demo_adjustedLoadTS, "--r", label="demo shifted load")
					ax[2].legend()
				ax[2].set_ylabel("load (MW)")
				ax[2].set_xlabel("hours")
				ax[0].set_title("Load: "+self.name)
				fig.tight_layout()
			
			else:
				plt.figure()
				plt.step(ts, self.loadTS, label = "original load")
				if "shift" in self.flexLoadMod:
					plt.step(ts, demo_adjustedLoadTS, "--r", label="demo shifted load")
					plt.legend()
				plt.ylabel("load (MW)")
				plt.xlabel("hours")
				plt.title(f"Load: {self.name}")
				plt.legend()
				plt.tight_layout()
										
			#plt.show()
		
		return 1
	
	
	
	# calculate spreading distance (hours) based on theoretical energy capacity (either call this or set things manually)
	def setCapacity(self, nominalEnergyCapacity, nominalPowerCapacity):
	
		self.cap_energy = nominalEnergyCapacity
		self.flexload = nominalPowerCapacity
		
		
		# formula for starge capacity based on DR power and time spread for a step function is 
		#   EnergyStorageCapacity = 0.25*(halfTspread + 1)*DRcapacity
		#   --> halfTspread  =  EnergyStorageCapacity*4/nominalPowerCapacity - 1
		self.shift_span = nominalEnergyCapacity*4/nominalPowerCapacity - 1
		
	 
	
	
	# for load-shifting type flexLoad, this applies load spreading on the time series in one pass
	def applyLoadShift(self, supply, totalLoad, preview=0):
		# paramaters: supply time series, original load time series, load time series out after load shift
		
		if "shift" in self.flexLoadMod:
			
			print("Performing load shifting for "+self.name)
			
			# -----
			def shiftLoadInstant(shiftMag, nbck, nfwd, i, totalLoad, localLoad):
				'''sub-function to do load shifting for one load instant'''
				
				
				# first spread the load change over adjacent hours
				inds = list(range(-nbck,0))+list(range(1,nfwd+1))
				'''
				shiftMagActual = 0
				
				for j in inds:
					delta = np.min([localLoad[i+j], shiftMag/(nbck+nfwd)]) # ensure we won't shift a load negative
				
					totalLoad[i+j] -= delta
					localLoad[i+j] -= delta
					
					shiftMagActual += delta
				
				
				totalLoad[i] += shiftMagActual                   # how much to increase load at current hour
				localLoad[i] += shiftMagActual      
					
				'''
				# first spread the load change over adjacent hours
				inds = list(range(-nbck,0))+list(range(1,nfwd+1))
				
				target = totalLoad[i]+shiftMag   # this is what we're aiming for (and will also be used as a limit on how much to shift adjacent loads)
				
				#print(f"\n i={i} - shiftMag is {shiftMag}")
				
				shiftMagActual = 1.0*shiftMag                                       # how much we shift may get reduced, so track that here
				delta0 = -shiftMag/(nbck+nfwd)                                       # how much to distribute the load to each pont (start by assuming uniform)
				
				
				for j in inds:
					# at one point I used target = targetLoad[i+j], a precomputed time series, but it had little advantage		
										
					delta = np.clip(delta0, -localLoad[i+j], None)             # ensure we won't shift this particular load negative
				
					if shiftMag < 0:   # if trying to spread a load peak	
						if totalLoad[i+j] >= target:                           # if the load at this instant is already above the target value, do nothing
							delta = 0
						elif totalLoad[i+j]+delta > target:                    # if load+delta at this instant is above what we're shifting time i's load to
							delta = target - totalLoad[i+j]                    # limit it   
					
					elif shiftMag > 0:   # if trying to spread a load dip	
						if totalLoad[i+j] <= target:                           # if the load at this instant is already above the target value, do nothing
							delta = 0
						elif totalLoad[i+j]+delta < target:                    # if load+delta at this instant is below what we're shifting time i's load to
							delta = target - totalLoad[i+j]                    # limit it   
					
					totalLoad[i+j] += delta
					localLoad[i+j] += delta
					
					shiftMagActual -= (delta - delta0)    # reduce how much the load at i will be shifted to be consistent	
					#print(f" j={j} - shiftMagActual is {shiftMagActual}")
				
				totalLoad[i] += shiftMagActual                   # how much to increase load at current hour
				localLoad[i] += shiftMagActual      
								
				return
			# -----	
			
			nfwd = np.int(self.time_shift_fwd)
			nbck = np.int(self.time_shift_bck)
			
			n = len(totalLoad)
			
			totalLoad2          = np.array(totalLoad)  # make the output array be a copy of the input total load array
			
			targetTotalLoad          = np.array(totalLoad)  # this array will contain the target load time series
			
			
			# peak shaving mode   (load peak, not net load peak!)
			if self.flexLoadMod == "shift: min peak":
				print("peak shaving mode")
				
				peakload = np.max(self.loadTS)   # peak of this specific load (used for scaling)
				
				for i in range(nbck, n-nfwd):					# go through hours except at ends	

					i1 = np.max([0,i-2*nbck])   # scaling the comparision window to 4X the time shift range
					i2 = np.min([n,i+2*nfwd])
					
					meanLoad = np.mean(totalLoad[i1:i2])
					#if totalLoad[i] > meanLoad:   # if above the average load for this time span
					#targetTotalLoad[i] = np.max([0.5*(meanLoad+totalLoad[i]), totalLoad[i]-self.flexLoadTS[i]])     # lower toward the mean load by up to [flexload] amount
					targetTotalLoad[i] = np.max([0.5*(meanLoad+np.max(totalLoad[i1:i2])), totalLoad[i]-self.flexLoadTS[i]])     # set a target value within the range of the load flexibility
					#targetTotalLoad[i] = np.max([meanLoad, totalLoad[i]-self.flexLoadTS[i]])     # lower toward the mean load by up to [flexload] amount

				for i in range(nbck, n-nfwd):	# go through hours except at ends again, and this time apply the shift	
						
					#threshold = np.mean(totalLoad[i1:i2]) + np.std(totalLoad[i1:i2])*np.max([0,(12-mspread)/24])
					shiftMag = targetTotalLoad[i] - totalLoad[i]
					
					if shiftMag < 0:
						shiftLoadInstant(shiftMag, nbck, nfwd, i, totalLoad2, self.adjustedLoadTS)
						
				
			# export minimization mode
			elif self.flexLoadMod == "shift: min export":
				print("export minimization mode")
					
				netpower = supply - totalLoad    # net power supply time series
				
				for i in range(nbck, n-nfwd):
				
					# how much to shift the load for anything above the threshold (from 0 to 1)
					if self.flexLoadTS[i] > 0:
						shiftMag = np.clip(netpower[i], 0, self.flexLoadTS[i])  # get positive net power values up to max of flexload
						targetTotalLoad[i] = totalLoad[i] + shiftMag
					
				for i in range(nbck, n-nfwd):	# go through hours except at ends again, and this time apply the shift	
						
					shiftMag = targetTotalLoad[i] - totalLoad[i]
					
					if shiftMag > 0:
						shiftLoadInstant(shiftMag, nbck, nfwd, i, totalLoad2, self.adjustedLoadTS)
							
							
			# shortage minimization mode
			elif self.flexLoadMod == "shift: min shortage":
				print("shortage minimization mode")

				netpower = supply - totalLoad    # net power supply time series

				for i in range(nbck, n-nfwd):						# go through hours except at ends
				
					if self.flexLoadTS[i] > 0:
						shiftMag = np.clip(netpower[i], -self.flexLoadTS[i], 0)  # get negative net power values up to max of -flexload
						targetTotalLoad[i] = totalLoad[i] + shiftMag
				
				for i in range(nbck, n-nfwd):	# go through hours except at ends again, and this time apply the shift	
						
					shiftMag = targetTotalLoad[i] - totalLoad[i]
					
					if shiftMag < 0:             
						shiftLoadInstant(shiftMag, nbck, nfwd, i, totalLoad2, self.adjustedLoadTS)
		
			if preview > 1:  # only show in-depth plots if preview level 2 or above is specified
				ts= np.arange(len(self.loadTS))
				plt.figure()
				plt.step(ts, self.loadTS, label = "original load")
				plt.step(ts, self.adjustedLoadTS, "--r", label="shifted load")
				plt.step(ts, totalLoad, ":k", label="total load")
				plt.step(ts, targetTotalLoad, ":g", label="target load")
				plt.title("Load: "+self.name)
				plt.legend()
		
			return totalLoad2
			
		else:
			return totalLoad  # if this isn't a shifting type of flexible load, don't do anything
	
			
	# time step for storage-like behaviour of flexible loads of type 1	
	def timeStep(self, t, dt, powerRequested):
	
		# t - current time (in hours)  <<<<<<<< watch out - gonna use this as an index for now...
		i = np.int(t)
		# dt - time step size to move forward by (hours)
		# powerRequested (how much power is desired to discharge (or charge if negative) from the storage  (MW)

		#dt = time - self.lastTime

		#TODO incorporate flexloadTS below


		if self.flexLoadMod=="storage based":
		
			# if time-varying availability data has been provided, use it to calculate the current available capacity
			if len(self.availability_time) > 0:
				av = np.interp(t, self.availability_time, self.availability_fraction)  # get what fraction of capacity is currently avialable
				self.SOC = np.min([self.SOC, self.cap_energy*av])                      # if current capacity is below SOC, reduce the SOC (and where does this energy go?)
			else:
				av = 1 # full capacity available
			
		
			
			P2Echarge    = dt*self.eta_charge     # factor to go from charge rate to delta SOC
			P2Edischarge = dt/self.eta_discharge  # factor to go from discharge rate to delta -SOC
			
			if powerRequested < 0.0:              # if charging
				power = np.max([-self.cap_charge*av, powerRequested])      # limit charge rate
				if self.SOC < self.cap_energy*av:       # if room to charge battery
					if self.SOC - power*P2Echarge >= self.cap_energy*av: # if going to fill
						power = (self.SOC - self.cap_energy*av)/P2Echarge
						self.SOC = self.cap_energy*av
					else:
						self.SOC = self.SOC - power*P2Echarge
						
				else:   # battery full 
					self.SOC = self.SOC
					power = 0
			else:                         	 # if discharging
				power = np.min([self.cap_discharge*av, powerRequested, self.loadTS[i]])      # limit discharge rate
				if self.SOC > 0:          # if room to discharge battery
					if self.SOC - power*P2Edischarge<= 0:  # if going to empty
						power = self.SOC/P2Edischarge
						self.SOC = 0
						
					else:
						self.SOC = self.SOC - power*P2Edischarge
						
				else:   # battery already empty
					self.SOC = self.SOC
					power = 0
			
			
			self.SOC = np.max([0, self.SOC - self.self_discharge*self.cap_energy*dt])   # apply self discharge!
					
			self.SOCTS[         i] = self.SOC   # record state of charge
			self.adjustedLoadTS[i] -= power   # record adjusted load
			
			self.lastTime = t
			
		
			return power   # generated electricity positive, consumed electricity negative <<<<<<<<<<<<<<<<<
		
		else:
			return 0
	
	
	
	#TODO: make some plots showing the DR behaviour for generic cases when a user hits preview
	def demo(self):
		# makes plots showing behaviour in example cases
		
		t = np.arange(32)
		l = np.zeros(32)+10
		l[10:20]=20

		l2 = l+0

		for i in range(len(l)):
			if l[i] > 10:
				l2[i] -= 8
				for j in range(1,5):
					l2[i+j] += 1
					l2[i-j] += 1

		plt.plot(l)
		plt.plot(l2)
		plt.show()
		 
# A case of 6 MW DR spread out +/-3 dt gives a storage of 1+2+3 = 6dt
# A case of 6 MW DR spread out +/-6 dt gives a storage of 10.5dt

# formula for starge capacity based on DR power and time spread for a step function is 
#   EnergyStorageCapacity = 0.25*(halfTspread + 1)*DRcapacity



## Storage class - this provides a generic model of energy storage technologies
class generator:
	
	def __init__(self):
		
		self.name = "generic generator"
		
		self.costcap = 0 # capital cost coefficient ($/W)
		self.costfix = 0 # fixed OpEx coefficient ($/W/yr)
		self.costvar = 0 # variable OpEx coefficient ($/MWh generated)
		self.life  = 0
		self.disrate = 0  # discount rate
		self.GHGcap = 0 # <<<<< unused (t/W)
		self.GHGfix = 0 # <<<<< unused (t/W/yr)
		self.GHGvar = 0 # (kg/MWh = g/kWh generated)
		
		self.capacity      = 0.0  # [MW] rated power

		self.resourceTS = []     # time series of resource availability [MW]
		self.genTS      = []     # time series of generation [MW]
		
		# the below can be populated with numpy arrays of how the available capacity varies over time
		self.availability_time     = [] # time instances (in hours)
		self.availability_fraction = [] # available capacity (relative to full capacity, ranging from 0 to 1)
		
		self.genTSused      = [] 
		self.genTScurtailed = [] 
		
		
		self.lastTime = 0.0
		
		
	def create(self, wb, col, preview=0):
	
		gencellref = wb.sheets("generation").range((4,col))  # the range of the title of the time series

		self.name = gencellref.value
		
		print("--------------------------")
		print("Creating generator "+self.name)

		# get initial time series
		self.resourceTS, resName, resUnit = getTS(getCellRef(wb.sheets("generation").range((7,col))))   # starting time series

		# apply seasonal scaling (multiply)
		form_scale = getCellRef(wb.sheets("generation").range((8,col)), optional=True)
		if form_scale != None:
			scaleTS, scaleTSname, scaleTSunits = getTS(form_scale)
			self.resourceTS = self.resourceTS*scaleTS 
			
		# apply seasonal shifting (add)
		form_shift = getCellRef(wb.sheets("generation").range((9,col)), optional=True)
		if form_shift != None:
			shiftTS, shiftTSname, shiftTSunits = getTS(form_shift)
			self.resourceTS = self.resourceTS + shiftTS	

		# apply performance curve conversion if applicable
		form_perf = getCellRef(wb.sheets("generation").range((11,col)), optional=True)
		if form_perf != None:
			ycurve, xcurve, curvename, _, _ = gatherCurve(wb, form_perf)  # get the performance curve x and y data
			
			#form_xperf = getCellRef(wb.sheets("generation").range((16,col)))
			#xdata, _, _ = gatherTS(wb, form_xperf)                        # this is the x(t) data used to feed the performance curve
					
			#perf_scalar = np.interp(xdata, xcurve, ycurve)
			
			print("interpolating curve")
			#print(xcurve)
			#print(ycurve)
			
			self.genTS = np.interp(self.resourceTS, xcurve, ycurve)
		else:
			self.genTS = np.array(self.resourceTS)
				
		
		# production shaping/scaling
		energy = getCellVal(wb, "generation", (13,col), errorval=None)
		peakP  = getCellVal(wb, "generation", (14,col), errorval=None)
		if (energy != None) and (peakP != None):
			self.genTS = scaleAndOffset(self.genTS, energy=energy, peak=peakP)
		elif energy != None:
			self.genTS = scaleAndOffset(self.genTS, energy=energy)
		elif peakP != None:
			self.genTS = scaleAndOffset(self.genTS, peak=peakP)
			
		self.capacity = np.max(self.genTS)
		
		# apply efficiency scalar AFTER identifying capacity (this reduces the power output, but not the definition of rated capacity)
		efficiency  = getCellVal(wb, "generation", (15,col))
		if efficiency != None:
			self.genTS = self.genTS*efficiency
			
			
		print(" - capacity:  {:6.1f} MW".format(self.capacity))
		print(" - peak gen:  {:6.1f} MW".format(np.max(self.genTS)))
		print(" - cap. fac.: {:6.1f} %".format(np.sum(self.genTS)/self.capacity*100/8760))
			
		# financial specs
		self.life    = getCellVal(wb, "generation", (17,col))
		self.disrate = getCellVal(wb, "generation", (18,col))
		self.costcap = getCellVal(wb, "generation", (19,col))
		self.costfix = getCellVal(wb, "generation", (20,col))
		self.costvar = getCellVal(wb, "generation", (21,col))		
		#self.GHGcap = getCellVal(wb, "generation", (31,col))
		#self.GHGfix = getCellVal(wb, "generation", (32,col))
		self.GHGvar = getCellVal(wb, "generation", (22,col))
			
		# initialize additional time series for later
		self.genTSused      = np.array(self.genTS)	
		self.genTScurtailed = np.zeros(self.genTS.shape)

		if preview > 0:
			fig,ax = plt.subplots(2,1,sharex=True)
			ax[0].plot(self.resourceTS)
			ax[1].plot(self.genTS)
			ax[0].set_ylabel(resName+" ("+resUnit+")")
			ax[1].set_ylabel("output (MW)")
			ax[0].set_title("Generation: "+self.name)
			fig.tight_layout()
		
		#plt.show()
		
		
		return 1

	def getCost(self):

		# amortize to annual basis and add in operating costs (in M$/year)
		annual_cost = -np.pmt(self.disrate, self.life, self.costcap*self.capacity, 
		                      fv=0) + self.costfix*self.capacity + self.costvar*np.sum(self.genTS)

		# emissions on an annual basis (kg CO2e/year)
		annual_emissions = self.GHGcap*self.capacity/self.life + self.GHGfix*self.capacity + self.GHGvar*np.sum(self.genTS)
	
		print(self.name+" - generation COE:  {:6.1f} $/MWh".format(annual_cost/np.sum(self.genTSused)*1e6))
		#print(self.name+" - generation GHGs: {:6.1f} kg CO2e/MWh".format(annual_emissions/np.sum(self.genTSused)))
	
		return annual_cost, annual_emissions
		

## Storage class - this provides a generic model of energy storage technologies
class storage:
	
	def __init__(self):
		
		self.name = "generic storage"
		
		
		self.costcap = 0 # capital cost coefficient ($/W)
		self.costfix = 0 # fixed OpEx coefficient ($/W/yr)
		self.costvar = 0 # variable OpEx coefficient ($/Wh generated)
		self.maxCycles=0   # max throughput for lifetime, normalized by energy capacity (MWh total charging / MWh capacity)
		self.disrate = 0  # discount rate
		self.GHGcap  = 0 # unused <<<< (t/W)
		self.GHGfix  = 0 # unused <<<< (t/W/yr)
		self.GHGvar  = 0 # (kg/MWh = g/kWh throughput)
		
		self.cap_energy      = 0.0  # [MWh] energy capacity
		self.cap_charge      = 0.0  # [MW] max charge rate
		self.cap_discharge   = 0.0  # [MW] max discharge rate
		self.self_discharge  = 0.0  # [%/hr] self discharge rate in % of full capacity per day
		self.eta_charge      = 1.0  # [%] efficiency of charging
		self.eta_discharge   = 1.0  # [%] efficiency of discharging
		
		self.SOC             = 0.0  # [] state of charge
		self.SOCTS           = np.zeros(8760)  # state of charge time series (length hardcoded for now)
		
		# the below can be populated with numpy arrays of how the available capacity varies over time
		self.availability_time     = [] # time instances (in hours)
		self.availability_fraction = [] # available capacity (relative to full capacity, ranging from 0 to 1)
		
		
		self.lastTime = 0.0
		
	
	def create(self, wb, col, preview=0):
	
		cellref = wb.sheets("storage").range((4,col))  # the range of the title of the time series

		self.name = cellref.value		
		
		print("--------------------------")
		print("Creating storage "+self.name)

		self.cap_energy    = getCellVal(wb, "storage", (5 ,col))
		self.cap_charge    = getCellVal(wb, "storage", (6 ,col))
		self.cap_discharge = getCellVal(wb, "storage", (7 ,col))
		self.self_discharge= getCellVal(wb, "storage", (8 ,col))
		self.eta_charge    = getCellVal(wb, "storage", (9 ,col))
		self.eta_discharge = getCellVal(wb, "storage", (10,col))		
		
		# financial specs
		self.maxCycles = getCellVal(wb, "storage", (17,col))
		self.life      = getCellVal(wb, "storage", (18,col))
		self.disrate   = getCellVal(wb, "storage", (19,col))
		self.costcap   = getCellVal(wb, "storage", (20,col))
		self.costfix   = getCellVal(wb, "storage", (21,col))
		self.costvar   = getCellVal(wb, "storage", (22,col))
		
		#self.GHGcap = getCellVal(wb, "storage", (24,col))
		#self.GHGfix = getCellVal(wb, "storage", (25,col))
		self.GHGvar = getCellVal(wb, "storage", (23,col))
		
		
		# previewing feature
		if preview > 0:
			n = 8760
			pwr = np.zeros(n)
			soc = np.zeros(n)
			self.SOC = self.cap_energy
			for i in range(8760):
				if i < 1000:
					pwr_requested = 0
				elif i < 2000:
					pwr_requested =  0.1*self.cap_energy
				elif i < 3000:
					pwr_requested = -0.1*self.cap_energy
				elif i == 3000:
					pwr_requested = 0.5*self.cap_energy
				elif i == 3002:
					pwr_requested = -0.5*self.cap_energy
				else:
					pwr_requested = 0
					
				pwr[i] = self.timeStep(i, 1, pwr_requested)
				soc[i] = self.SOC
				
			fig, ax = plt.subplots(2,1,sharex=True)
			ax[0].plot(pwr)
			ax[1].plot(soc)
			
		

		return 1
		
		# TODO: add in temporal variability
		
		
	def timeStep(self, t, dt, powerRequested):
	
		# t - current time (in hours)
		# dt - time step size to move forward by (hours)
		# powerRequested (how much power is desired to discharge (or charge if negative) from the storage  (MW)

		#dt = time - self.lastTime
		
		# if time-varying availability data has been provided, use it to calculate the current available capacity
		if len(self.availability_time) > 0:
			av = np.interp(t, self.availability_time, self.availability_fraction)  # get what fraction of capacity is currently avialable
			self.SOC = np.min([self.SOC, self.cap_energy*av])                      # if current capacity is below SOC, reduce the SOC (and where does this energy go?)
		else:
			av = 1 # full capacity available
		
		
		P2Echarge    = dt*self.eta_charge     # factor to go from charge rate to delta SOC
		P2Edischarge = dt/self.eta_discharge  # factor to go from discharge rate to delta -SOC
		
		if powerRequested < 0.0:              # if charging
			power = np.max([-self.cap_charge*av, powerRequested])      # limit charge rate
			if self.SOC < self.cap_energy*av:       # if room to charge battery
				if self.SOC - power*P2Echarge >= self.cap_energy*av: # if going to fill
					power = (self.SOC - self.cap_energy*av)/P2Echarge
					self.SOC = self.cap_energy*av
				else:
					self.SOC = self.SOC - power*P2Echarge
					
			else:   # battery full 
				self.SOC = self.SOC
				power = 0
		else:                         	 # if discharging
			power = np.min([self.cap_discharge*av, powerRequested])      # limit discharge rate
			if self.SOC > 0:          # if room to discharge battery
				if self.SOC - power*P2Edischarge<= 0:  # if going to empty
					power = self.SOC/P2Edischarge
					self.SOC = 0
					
				else:
					self.SOC = self.SOC - power*P2Edischarge
					
			else:   # battery already empty
				self.SOC = self.SOC
				power = 0
	
		self.lastTime = t
		
		self.SOC = np.max([0, self.SOC - self.self_discharge*self.cap_energy*dt])   # apply self discharge!
		
		self.SOCTS[np.int(t)] = self.SOC   # record state of charge
	
		return power   # generated electricity positive, consumed electricity negative
	
	
	
	def getCost(self):

		dSOCs = np.diff(self.SOCTS)
		throughput = np.sum( dSOCs*(dSOCs>0))  # total energy throughput of battery (measured by INPUT only)

		if self.life == None:  # if no lifetime specified, use 100 years as safe upper limit
			self.life = 100
		
		if throughput > 0 and self.maxCycles != None:  #  lifetime is when energy throughput hits N times capacity, or when year limit is surpassed
			lifetime = np.min([self.maxCycles*self.cap_energy/throughput, self.life])    
		else:
			lifetime = self.life

		self.life = lifetime		

		# amortize to annual basis and add in operating costs (in M$/year)
		annual_cost = -np.pmt(self.disrate, self.life, self.costcap*self.cap_energy, fv=0) + self.costfix*self.cap_energy + self.costvar*throughput

		# emissions on an annual basis (kg CO2e/year)
		annual_emissions = self.GHGcap*self.cap_energy/self.life + self.GHGfix*self.cap_energy + self.GHGvar*throughput
	
		print(self.name+" - storage COE: {:6.1f} $/MWh".format(annual_cost/throughput*1e6))
	
		return annual_cost, annual_emissions
	

## BEV class - this provides a generic model of BEVs with V2G capacility
class BEV:
	
	def __init__(self):
		
		self.name = "generic storage"
		self.cap_energy      = 0.0  # [MWh] energy capacity
		self.cap_charge      = 0.0  # [MW] max charge rate
		self.cap_discharge   = 0.0  # [MW] max discharge rate
		self.self_discharge  = 0.0  # [%/day] self discharge rate in % of full capacity per day
		self.eta_charge      = 1.0  # [%] efficiency of charging
		self.eta_discharge   = 1.0  # [%] efficiency of discharging
		
		self.SOC             = 0.0  # [] state of charge
		
		self.SOCTS           = np.zeros(8760)  # state of charge time series (length hardcoded for now)
		
		
		# the below can be populated with numpy arrays of how the available capacity varies over time
		self.availability_time     = [] # time instances (in hours)
		self.availability_fraction = [] # available capacity (relative to full capacity, ranging from 0 to 1)
		self.dSOC_in               = [] # relative energy being added to SOC by vehicles plugging in minus vehicles leaving [MWh]
		self.dSOC_out              = []
		
		self.lastTime = 0.0
		
		
		
	def create(self, wb, col, preview=0):
			
		cellref = wb.sheets("BEVs").range((4,col))  # the range of the title of the time series

		self.name = cellref.value

		print("--------------------------")
		print("Creating BEV fleet "+self.name)

		self.cap_energy    = getCellVal(wb, "BEVs", (5 ,col))
		self.min_SOC_fix   = getCellVal(wb, "BEVs", (6 ,col))
		self.cap_charge    = getCellVal(wb, "BEVs", (7 ,col))
		self.cap_discharge = getCellVal(wb, "BEVs", (8 ,col))
		self.self_discharge= getCellVal(wb, "BEVs", (9 ,col))
		self.eta_charge    = getCellVal(wb, "BEVs", (10,col))
		self.eta_discharge = getCellVal(wb, "BEVs", (11,col))
		#self.annualEnergy = getCellVal(wb, "BEVs", (12,col))
		#self.eta_discharge = getCellVal(wb, "BEVs", (13,col))  # <<<< serasonal variation time serires to be dded later
		self.BEVmodel      = getCellVal(wb, "BEVs", (14,col), errormsg="BEV model 1 or 2 must be provided")

		# ---------------------- BEV model 1 setup -------------------------

		if self.BEVmodel==1:
		
			# get initial time series of fleet use for the year
			#inputTS, inputTSname, inputTSunits = gatherTS(wb, wb.sheets("BEVs").range((17,col)))   # starting time series
			inputTS, inputTSname, inputTSunits = getTS(getCellRef(wb.sheets("BEVs").range((17,col))))   # starting time series

			inputTS = inputTS/np.sum(inputTS)   # scale to sum to 1
			
			self.unavail_frac = getCellVal(wb, "BEVs", (18,col), errorval=0)  # what fraction not particpiating in grid
			
			self.availability_fraction = 1- self.unavail_frac - inputTS               # this is availability of EVs for grid interaction
			self.availability_fraction = self.availability_fraction*(self.availability_fraction > 0)  # ensure it doesn't go negative

			self.annualEnergy = 1000*getCellVal(wb, "BEVs", (12,col), errormsg="annual energy consumption must be provided")  # this is annual energy need accounting for charge inefficiency

			self.loadTS = self.annualEnergy*inputTS*self.eta_charge # represent EV charging needs as load proporational to usage time series (note that charge efficiency effect is removed here so this is what comes from the battery)
			
			self.availability_time=np.arange(len(inputTS))
			
			if preview > 0:
				fig, ax = plt.subplots(2,1,sharex=True)
				ax[0].plot(self.availability_time, self.loadTS)
				ax[1].plot(self.availability_time, self.availability_fraction)
				
				ax[1].set_xlabel("time (hrs)")
				ax[0].set_ylabel("net load (MWh)")
				ax[1].set_ylabel("availability (%)")
				ax[0].set_title("BEV fleet: "+self.name+" (total load {:.3f} GWh/year)".format(self.annualEnergy/1000))
				fig.tight_layout()
				

		# ---------------------- BEV model 2 setup -------------------------
		elif self.BEVmodel==2:
			# get initial time series
			#loadTS, _, _ = gatherTS(wb, wb.sheets("BEVs").range((7,col)))   # starting time series

			# >> the following is like calling gatherTS except it's for the four-column arrival-departure data for EVs, so a custom form <<

			# gets/creates annaul TS data from annual,weekly, or daily inputs
			#def gatherTS(wb, cellref):

			row = cellref.row
			col = cellref.column  # redundant <<<
			sht = cellref.sheet
				
			departure_fracs, _, _ = getTS(getCellRef(wb.sheets("BEVs").range((21,col))))   # fraction departing time series
			departure_SOCs , _, _ = getTS(getCellRef(wb.sheets("BEVs").range((22,col))))   # average departing SOC time series
			arrival_fracs  , _, _ = getTS(getCellRef(wb.sheets("BEVs").range((23,col))))   # fraction arriving time series
			arrival_SOCs   , _, _ = getTS(getCellRef(wb.sheets("BEVs").range((24,col))))   # average arriving SOC time series
	
			'''
			form_an = getCellRef(wb.sheets("BEVs").range((22, col)))
			form_wk = getCellRef(wb.sheets("BEVs").range((26, col)))
			form_wd = getCellRef(wb.sheets("BEVs").range((27, col)))
			form_we = getCellRef(wb.sheets("BEVs").range((28, col)))

			# try annual time series
			if form_an != None:
				#print("line 109 about to getTS from")
				#print(form_an)
				y_an, name, units = getBEVTS(wb, form_an)   # note that getBEVTS returns an Nx4 array (four columns)
				
				if len(y_an) < 8760:
					y_an = np.tile(y_an, np.int(8760/len(y_an))+1)
					
			# otherwise try weekly time series
			elif form_wk != None:
				y_wk, name, units = getBEVTS(wb, form_wk)
				
				#TODO: ensure time series is one week long exactly!
				
				y_an = np.tile(y_wk, np.int(8760/len(y_wk))+1)
				
			# otherwise try daily time series
			elif form_wd != None:
				y_wd, name, units = getBEVTS(wb, form_wd) # weekday
				
				if form_we != None:  # if weekends are different
					y_we, name, units = getBEVTS(wb, form_we) # weekend day
					y_an = np.tile(np.hstack([y_wd, y_wd, y_wd, y_wd, y_wd, y_we, y_we]), np.int(365/7)+1)  # make the year
				
				else:   # one daily pattern only case (no weekend distinction) 
					y_an = np.tile(y_wd, 365)
			else:
				print("Error loading BEV time series for "+self.name)

			#plt.plot(y_an.transpose())
			#plt.show()
			
			arrival_fracs   = y_an[2,:]
			arrival_SOCs    = y_an[3,:]
			departure_fracs = y_an[0,:]
			departure_SOCs  = y_an[1,:]
			'''
				
			initial_fraction=1
			dt=1
			
			# set up time series for this object, do preprocessing
			#def initialize(self, arrival_fracs, arrival_SOCs, departure_fracs, departure_SOCs, initial_fraction=1, dt=1):
			# initial_fraction - fraction of fleet plugged in at t0
			
			#<<<<<<<<<<<<< nee
			
			# go through departure and arrival numbers and compute three things:
			# 1. capacity time series		
			# 2. SOC delta time series based on arrivals and departures
			# 3. minimum SOC to satisfy departure demands (this is done from end to start, based on instantaneous charge limit)
			
			self.availability_time=np.arange(len(arrival_fracs))  # <<< likely redundant, and just 0:8760
			
			N = len(self.availability_time)
			self.dSOC_in  = np.zeros(N)
			self.dSOC_out = np.zeros(N)
			self.min_SOC  = np.zeros(N)   # [MWh] - I think this is the minimum SOC (in total stored energy) of the plugged in EVs only
			
			self.availability_fraction = np.zeros(N)   # fraction of fleet plugged in
			self.available_SOC = np.zeros(N)           # fraction of full-fleet full SOC that is plugged in SOC
			
			self.availability_fraction[0] = initial_fraction
			self.available_SOC[0] = self.availability_fraction[0]
			
			for i in range(1,N):
				self.availability_fraction[i] = self.availability_fraction[i-1] + arrival_fracs[i] - departure_fracs[i]
							
				self.dSOC_in[i] = arrival_fracs[i]*arrival_SOCs[i] 
				self.dSOC_out[i]= departure_fracs[i]*departure_SOCs[i]
				
				#self.available_SOC[i] = self.available_SOC[i-1] + arrival_fracs*arrival_SOCs - departure_fracs*departure_SOCs
				
			# figure out minimum SOC needed for all the departures	
			for i in range(N-1,0,-1):	
				if departure_fracs[i] > 0:      # if cars departing at this time step
					SOC_out = departure_fracs[i]*departure_SOCs[i]*self.cap_energy   # how much energy is leaving
					for j in range(24):         # go back in time and specify minimum SOC required to support this
						max_dSOC = np.min([SOC_out, self.cap_charge*self.availability_fraction[i-j]])*dt + arrival_fracs[i-j]*arrival_SOCs[i-j]*self.cap_energy    # max SOC contribution possible in this step (based on max charge rate plus any returning vehicles)
						
						#print(str(i)+" "+str(i-j)+" "+str(SOC_out)+" "+str(max_dSOC))
						
						if SOC_out - max_dSOC > 0:
							self.min_SOC[i-j] += SOC_out - max_dSOC   # required SOC at this step given charge rate limit needed to meet departure SOC
							SOC_out -= max_dSOC                       # remaining amount of SOC needed at previous steps
						else:
							break
			
			# calculate hypothetical load time series if all charging/discharging was on demand (load is positive
			self.loadTS = (self.dSOC_out - self.dSOC_in)*self.cap_energy
			
			
			if preview > 0:
				fig, ax = plt.subplots(3,1,sharex=True)
				ax[0].step(self.availability_time, -self.dSOC_out, color="r")
				ax[0].step(self.availability_time, self.dSOC_in, color="g")
				#ax[0].bar(self.availability_time, -self.dSOC_out, fc="r")  
				#ax[0].bar(self.availability_time, self.dSOC_in, fc="g")

				ax[1].plot(self.availability_time, 100*self.availability_fraction, "k"  , label="fraction plugged in")
				ax[1].plot(self.availability_time, 100*self.min_SOC/self.cap_energy, "b", label="minimum fleet SOC")
				
				ax[2].plot(self.availability_time, self.dSOC_out-self.dSOC_in)
				ax[2].set_xlabel("time (hrs)")
				ax[0].set_ylabel("stored energy\narrival (MWh)")
				ax[1].set_ylabel("Fractions (%)")
				ax[1].legend()
				ax[2].set_ylabel("net load before\nsmoothing (MW)")
				ax[0].set_title("BEV fleet: "+self.name+" (total load {:.3f} GWh/year)".format(np.sum(self.dSOC_out-self.dSOC_in)*dt/1000))
				fig.tight_layout()
				
			
			# enforce fixed minimum permissible state of charge for the vehicles
			for i in range(len(self.min_SOC)):
				self.min_SOC[i] = np.max([self.min_SOC[i], self.min_SOC_fix*self.availability_fraction[i]*self.cap_energy])
		
		else:
			raise Exception("ERROR: BEV model 1 or 2 must be specified for BEV fleet "+self.name)
			return 0
		
		return 1
		
		
	def timeStep(self, t, dt, powerRequested):
	
		# adjust below to constrain to instantaneous minimum SOC <<<<
		# also add in SOC delta time series due to departures and arrivals <<<<
	
		# t - current time (in hours)
		i = np.int(t)  # <<< not valid for non unity dt <<<
		# dt - time step size to move forward by (hours)
		# powerRequested (how much power is desired to discharge (or charge if negative) from the storage  (MW)

		#dt = time - self.lastTime
		
		
		if self.BEVmodel == 1:
		
			power_to_wheels = self.loadTS[i]   # this is how much power is needed for charging for fleet demand (in use...)
					
			av = np.interp(t, self.availability_time, self.availability_fraction)  # get what fraction of capacity is currently avialable
			
			# if vehicle departures cause available plugged-in capacity to drop below the state of charge, then cap SOC and subtract the excess from the demand
			if self.cap_energy*av < self.SOC:
				#print(f"At time {t} the capacity is saturated so putting excess toward charging demand. Excess is {(self.SOC - self.cap_energy*av)}")
				power_to_wheels -= (self.SOC - self.cap_energy*av)/dt
				self.SOC = self.cap_energy*av
			
			
			
			# three cases:
			if -powerRequested*self.eta_charge > power_to_wheels:              # if charging storage capacity (grid supplies more than cars demand)
				power = np.max([-self.cap_charge*av, powerRequested])*self.eta_charge + power_to_wheels  # power measured at battery (limit charge rate from grid)
				
				if self.SOC < self.cap_energy*av:       # if room to charge battery
					if self.SOC - power*dt >= self.cap_energy*av: # if going to fill
						power = (self.SOC - self.cap_energy*av)/dt
						self.SOC = self.cap_energy*av
					else:
						self.SOC = self.SOC - power*dt
						
				else:   # battery full 
					self.SOC = self.SOC
					power = 0
					
				power_out = (power - power_to_wheels)/self.eta_charge  # this will be negative (it's actually power in from grid)
				
				
			elif powerRequested < 0:                                      # if discharging storage capacity but getting some supply from grid (not V2G)
				power = np.max([-self.cap_charge*av, powerRequested])*self.eta_charge + power_to_wheels  # power measured at battery (limit charge rate from grid)
				
				if self.SOC > 0:          # if room to discharge battery
					if self.SOC - power*dt<= 0:  # if going to empty
						power = self.SOC/dt
						self.SOC = 0
						
					else:
						self.SOC = self.SOC - power*dt
						
				else:   # battery already empty
					self.SOC = self.SOC
					power = 0			
				
				power_out = (power - power_to_wheels)/self.eta_charge  # this will be negative (it's actually power in from grid)
				
				
			else:                         	 # if discharging storage capacity (to grid + car demand) (net V2G)			
				power = np.min([self.cap_discharge*av, powerRequested])/self.eta_discharge + power_to_wheels     # power measured at battery (limit discharge rate to grid only)
				
				if self.SOC > 0:          # if room to discharge battery
					if self.SOC - power*dt<= 0:  # if going to empty
						power = self.SOC/dt
						self.SOC = 0
						
					else:
						self.SOC = self.SOC - power*dt
						
				else:   # battery already empty
					self.SOC = self.SOC
					power = 0
		
				power_out = (power - power_to_wheels)*self.eta_discharge  # this will be positive
				
				#if t>2192:
				#	breakpoint()
			
			
			
			power = power_out
			
			
		
		elif self.BEVmodel == 2:
			#TODO: check for proper application of efficiency/losses in charge/discharge calcs driven by pReqd
			
			# if time-varying availability data has been provided, use it to calculate the current available capacity
			if len(self.availability_time) > 0:
				av = np.interp(t, self.availability_time, self.availability_fraction)  # get what fraction of capacity is currently avialable
				self.SOC = np.min([self.SOC, self.cap_energy*av])                      # if current capacity is below SOC, reduce the SOC (and where does this energy go?)
			else:
				av = 1 # full capacity available
			
			
			dSOC_in = np.interp(t, self.availability_time, self.dSOC_in)
			dSOC_out= np.interp(t, self.availability_time, self.dSOC_out)
			
			
			P2Echarge    = dt*self.eta_charge     # factor to go from charge rate to delta SOC
			P2Edischarge = dt/self.eta_discharge  # factor to go from discharge rate to delta -SOC
			
			# check for charging needs to meet imminent EV departure SOC requirements
			min_SOC_target = np.interp(t+dt, self.availability_time, self.min_SOC)  # minimum required SOC at next time step
			dSOC_net = dSOC_in - dSOC_out  	#  energy change from arriving and departing vehicles
			pReqd = (min_SOC_target - (self.SOC + dSOC_net))/dt # required *charging* power at this time step given above
			
			if powerRequested < 0.0 or pReqd > 0:              # if charging
				power = -np.min([self.cap_charge*av, np.max([-powerRequested, pReqd])])      # choose larger of required EV charging or grid-desired charging, then limit charge rate
				if self.SOC < self.cap_energy*av:       # if room to charge battery
					if self.SOC - power*P2Echarge >= self.cap_energy*av: # if going to fill
						power = (self.SOC - self.cap_energy*av)/P2Echarge
						self.SOC = self.cap_energy*av
					else:
						self.SOC = self.SOC - power*P2Echarge
						
				else:   # battery full 
					self.SOC = self.SOC
					power = 0
			else:                         	 # if discharging
				power = np.min([self.cap_discharge*av, np.min([powerRequested, -pReqd])])      # limit discharge rate (and ensure not less than needed for future EV departures)
				if self.SOC > min_SOC_target:          # if room to discharge battery
					if self.SOC - power*P2Edischarge<= min_SOC_target:  # if going to hit minimum, limit discharge to not go below it
						power = (self.SOC - min_SOC_target)/P2Edischarge
						self.SOC = min_SOC_target
						
					else:
						self.SOC = self.SOC - power*P2Edischarge
						
				else:   # battery already at minimum
					self.SOC = self.SOC
					power = 0
		
			self.SOC += dSOC_net  # account for energy from arriving and departing vehicles
		
		
		# in all cases:
		
		
		self.SOC = np.max([0, self.SOC - self.self_discharge*self.cap_energy*dt])   # apply self discharge!
		
		self.SOCTS[i] = self.SOC   # record state of charge
		
		self.lastTime = t
	
		return power   # generated electricity positive, consumed electricity negative

	
# load time series data if available, and synthesize into year-long data
def getTimeSeries(wb, sh, col, row):	

	ref = wb.sh(sh).range((col,row)).address
	
	



	
