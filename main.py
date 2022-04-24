# IRIS - Island Renewables Integration Simulator
# (c) Copyright Matt Hall and Matthew McCarville, 2019
# To be cleaned up and released under BSD license.


import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
from matplotlib import cm
from matplotlib.colors import ListedColormap, LinearSegmentedColormap


import numpy as np
import matplotlib.dates as dates

#from scipy import stats
from matplotlib import rc
import matplotlib.mlab as mlab       # for mlab's psd functionality

import xlwings as xw   # @mhall: xlwings seems to be a very powerful library for both reading and writing excel files

# load the libraries made for the model
import componentLib as components

# make sure we reload the libraries in case there are any changes
import importlib
components = importlib.reload(components)	

	
def ExcelHyperlink(cell, sheet):
	
	wb = xw.Book.caller()
	
	ButtonRange=wb.sheets(sheet).range(cell)
	
	i = ButtonRange.column
			
	if sheet=="overview":
	
		# for now we assume the user must have hit the run model button
		main()
	
	else:  # this is the preview case
	
		
		if wb.sheets(sheet).range((4,i)).value == None:
			raise NameError(f"This column ({i}) does not have a name so will not be loaded.")
			return 0
			
		else:
			
			if sheet=="load":			
				load1 = components.load()
				success = load1.create(wb, i, preview=1)
				
			elif sheet=="generation":
		
				gen1 = components.generator()
				success = gen1.create(wb,i, preview=1)
				
			elif sheet=="storage":
				stor1 = components.storage()
				success = stor1.create(wb,i, preview=1)
				
			elif sheet=="BEVs":
				BEV1 = components.BEV()
				success = BEV1.create(wb,i, preview=1)
				
			else:
				raise NameError("The sheet named "+sheet+ " does not have preview capabilities.")
				return
				
			
		if success >= 1:
			plt.show()
			print("it worked!")
			
		else:
			raise NameError("error creating the "+sheet+ " at "+sheet+" column "+str(i))
			
	return

# >>>>>>>>>>>>>>>>>>>>>>>>>> SCRIPT STARTS HERE <<<<<<<<<<<<<<<<<<<<<<<<<<<<<

# for frozen python using pyinstaller, see https://stackoverflow.com/questions/56006773/error-using-xlwings-runfrozenpython-arguments

def main():
#if __name__ == "__main__":

	# ==================== Set Up Time Info for Simulation ======================

	# make time array to be used in simulation for lining things up
	Tstart = np.datetime64("2016-01-01T00:00:00")       # start time in numpy datatime64 format
	Tstop  = np.datetime64("2016-12-31T23:59:59")       # end time

	dts = 60*60                                          # desired time step size (in seconds)

	time_d = np.arange(Tstart,Tstop,dts)                   # make the time array (in numpy datetime64 format)

	time_s = time_d.astype("timedelta64").astype(int)    # make a simple version of the array (in seconds)


		
	# =================== Set Up Input Data for Simulation =======================


	if __name__ == "__main__":
		wb = xw.Book('ModelSheetSample2.xlsm')  # connect xlwings to the main input Excel file
	else:
		wb = xw.Book.caller()  # connect xlwings to the main input Excel file


	# load a few general things centrally here
	costImports = components.getCellVal(wb, "overview", "K9" , errormsg="", errorval=0)
	costCapacity= components.getCellVal(wb, "overview", "K10", errormsg="", errorval=0)
	GHGImports  = components.getCellVal(wb, "overview", "K11", errormsg="", errorval=0)
	max_export  = components.getCellVal(wb, "overview", "K12", errormsg="", errorval=0)


	# ================= Aggregate load, supply, and storage for simulation ===============	

	dt = 1

	'''	
	dtHrs = dt/60/60
	print("time step in hrs :"+str(dtHrs))
	timeHrs = np.arange(0, len(WIND)*dtHrs, dtHrs) # make time vector in units of hours
	'''


	# ------------------------- Sources/Supply ----------------------------
		
	# simply put the relevant supply time series (which have already been interpolated to "time_s" into a list
	#SUPPLIES = [WIND, PV]

	SUPPLIES = []

	for i in range(3,20):	
		if wb.sheets("generation").range((4,i)).value == None:
			break
		else:
			gen1 = components.generator()
			success = gen1.create(wb,i, preview=0)
		if success >= 1:
			SUPPLIES.append(gen1)
		else:
			print("error creating generator "+wb.sheets("generation").range((4,i)).value)
			break
			
	print("Loaded "+str(len(SUPPLIES))+" generators.")


	# ---------------------------- Loads/Demand -----------------------------

	LOADS = []

	for i in range(3,20):	
		if wb.sheets("load").range((4,i)).value == None:
			break
		else:
			load1 = components.load()
			success = load1.create(wb, i, preview=0)	
		if success >= 1:
			LOADS.append(load1)
		else:
			print("error creating load")
			break

	print("Loaded "+str(len(LOADS))+" loads.")
		

	# ---------------------------- EV Resources ------------------------------

	BEVS = []

	for i in range(3,20):	
		if wb.sheets("BEVs").range((4,i)).value == None:
			break
		else:
			bev1 = components.BEV()
			success = bev1.create(wb, i, preview=0)
		if success >= 1:
			BEVS.append(bev1)
		else:
			print("error creating BEV fleet")
			break

	print("Loaded "+str(len(BEVS))+" BEV fleets.")

		
	# ------------------------- Storage Resources ----------------------------

	STORAGES = []   # initialize an empty list that will be populated with storage objects

	for i in range(3,20):	
		if wb.sheets("storage").range((4,i)).value == None:
			break
		else:
			stor1 = components.storage()
			success = stor1.create(wb, i)
		if success >= 1:
			STORAGES.append(stor1)
		else:
			print("error creating storage")
			break
			
	print("Loaded "+str(len(STORAGES))+" storage objects.")

	# =========================== Run the Model ============================
			
	# run the simulation
	#crunch(LOADS, SUPPLIES, STORAGES, BEVS, plots=1)

	loads     = LOADS
	supplies  = SUPPLIES
	storages  = STORAGES
	BEVs      = BEVS
	plots = 1
	time=[]


	n = len(loads[0].loadTS)  # this is hopefully the length of all time series....

	if len(time)==0:
		time = np.arange(n)

	nLoads    = len(loads   )
	nSupplies = len(supplies)
	nStorages = len(storages)
	nBEVs     = len(BEVs    )

	totalSupply = np.zeros(n)   # make a sum of the supplies
	totalLoad   = np.zeros(n)	
		
	for i in range(nLoads):
		totalLoad += loads[i].loadTS
		
	for i in range(nSupplies):
		totalSupply += supplies[i].genTS
		
	adjustedLoad = np.array(totalLoad) # duplicate the load (before potentially changing it)


	#@mhall: The following sections provide an order of operation for integrating the renewables and meeting demand.
	#        We could compare this order with what's described in Jacobson's textbook and adjust as we see fit.


	# -------------------------curtailment pass --------------------------

	# probably no curtailment necessary till later...


	# ---------------------- demand response passes --------------------------
	# Flexible loads may be modelled using a kernel that spreads the load from one time instant
	# to a number of adjacent time instants. This provides some approximation of both latent
	# storage within the load (the same overall load is met) and planning ahead (load can
	# be shifted to both later and earlier times). To model this in the system, each flexible
	# load over the full time series is modelled in turn.

	for i in range(nLoads):
		
		adjustedLoad = loads[i].applyLoadShift(totalSupply, adjustedLoad)


	# ------------------------- storage and BEV passes --------------------------

	# could apply filters to decompose net load and selectively apply storage tech...

	# initialize BEV and storage time series arrays (these record/store the storage behaviour over the simulation)

	load_supplies= np.zeros([nLoads,n])   # for type 1 flexible loads only (storage based), this is how 
	load_SOC     = np.zeros([nLoads,n])

	BEV_supplies = np.zeros([nBEVs,n])     # storage power in (MW) time series for each BEV fleet
	BEV_SOC      = np.zeros([nBEVs,n])     # state of charge time series for each BEV fleet

	storage_supplies = np.zeros([nStorages,n])     # storage power in (MW) time series for each storage object
	storage_SOC      = np.zeros([nStorages,n])     # state of charge time series for each storage object

	curtail = np.zeros(n) # time series of power curtailment

	# initially start each storage tech at half full (by setting end value of what would normally be previous run)
	for j in range(nStorages):
		storage_SOC[j,-1] = 0.5*storages[j].cap_energy   # they're started from this last value

	# run storage simulation
	for ibp in range(1):	  # loop through analysis multiple times starting with previous end SOCs to ensure energy balance

		supplies_used      = np.zeros([nSupplies, n])
		supplies_curtailed = np.zeros([nSupplies, n])
		
		# fill in supply used matrix with total generation to start with
		for j in range(nSupplies):
			supplies_used[j,:] = supplies[j].genTS
			
		
		# start each storage tech SOC to whatever the previous run finished at
		for j in range(nStorages):
			storages[j].SOC = storage_SOC[j,-1]

		# loop through time steps
		for i in range(0,n):
		

			netload = adjustedLoad[i] - totalSupply[i]   # get net load (@mhall: need to check all the logic around here)
			
			# adjust storage merit order depending on situation (@mhall: not used yet)
			if netload > 0:       # if discharging storage
				storage_order = range(nStorages)
			else:
				storage_order = range(nStorages,0,-1)
			
			
			# use flexible load storages
			for j in range(nLoads):
				load_supplies[j,i] = loads[j].timeStep(time[i], dt, netload)  # NOT planning to consider this a charging/discharging
				netload      -= load_supplies[j,i]
				adjustedLoad[i] -= load_supplies[j,i]  # apply this change to the adjusted load (since it's techincally a load change) 
				
				load_SOC[j,i] = loads[j].SOC # save state of charge (not necessary)
			
			
			# use BEVs			
			for j in range(nBEVs):
				BEV_supplies[j,i] = BEVs[j].timeStep(time[i], dt, netload)
				netload -= BEV_supplies[j,i]
				
				BEV_SOC[j,i] = BEVs[j].SOC # save state of charge (not necessary)
			
			
			# use storage	
			for j in range(nStorages):
				storage_supplies[j,i] = storages[j].timeStep(time[i], dt, netload)
				netload -= storage_supplies[j,i]
				
				storage_SOC[j,i] = storages[j].SOC # save state of charge (not necessary)
			
			# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
			# apply some curtailment?
			if netload < -max_export:
				curtail[i] = -netload - max_export
								
				# distribute curtailment proportionally among supplies
				for j in range(nSupplies):
					supplies_curtailed[j,i] = curtail[i]*supplies_used[j,i]/np.sum(supplies_used[:,i])
					
				for j in range(nSupplies):
					supplies_used[j,i] = supplies_used[j,i] - supplies_curtailed[j,i]   # subtract from supply time series   @mhall: check this
			
			

		# save used generation time series to each supply object
		for j in range(nSupplies):
			supplies[j].genTSused = supplies_used[j,:]

		'''	
		# check if energy balances, otherwise adjust starting battery capacity to equal final capacity and iterate
		if battState[-1] < 0.99*battState[0]-1:
			print("battery state inbalance. going from "+str(battState[0])+" to "+str(battState[-1]))
			battState[0] = battState[-1]
		elif nocurtail > 0:
			i_inbalance = np.argmax(np.abs(netenergy))  # find first case of curtailment or shortage

			if i_inbalance > 0 and np.abs(netenergy[i_inbalance]) > 1:   # if there is a signifcant inbalance
				battState[0] -= 0.95*netenergy[i_inbalance]                   # adjust starting battery state accordingly
				print("curtailment issue. adjusting starting cap by "+str(-netenergy[i_inbalance]))
			else:
				print("good, no curtailment")
				break
		else:
			print("good")
			break
		'''

	# load and demand response plot
	'''
	plt.figure()
	plt.plot(totalLoad)
	plt.plot(adjustedLoad, "--r")
	plt.title("total load after adjustment")
	plt.show()
	'''

	# calculating some last numbers

	#total_charge    = np.sum( -storage_supplies*(storage_supplies<0), axis=1)*dt # energy charging throughput <<<<<< not including BEVs yet
	#total_discharge = np.sum(  storage_supplies*(storage_supplies>0), axis=1)*dt # energy discharging throughput
		
		
	# -------------------------------------- load shedding? ------------------------------------

	#print(This_is_the_end_of_what_works_so_far__)

	# ------------------------- implied last step: power import/export -----------------------------
	#        (any mismatch of supply and demand is assumed to be met by import/export with NB)

	# power exchange with NB (positive = import)
	imported = adjustedLoad - np.sum(supplies_used, axis=0) - np.sum(storage_supplies, axis=0)  - np.sum(BEV_supplies, axis=0)   # [MW] 


	load_with_charging = adjustedLoad + np.sum(-storage_supplies*(storage_supplies<0), axis=0) + np.sum( -BEV_supplies*(BEV_supplies<0), axis=0)

	final_load_with_storage = adjustedLoad + np.sum(-storage_supplies, axis=0) + np.sum( -BEV_supplies, axis=0)

	# --------------------------- some final calculations of key numbers ---------------------------

	total_import = np.sum( np.maximum( imported, 0))*dt       # sum energy from imports only [MWh]
	total_export = np.sum( np.maximum(-imported, 0))*dt       # sum energy from exports only [MWh]

	supply_int = np.sum(totalSupply)*dt - total_export - np.sum(curtail)*dt        # the portion of on-Island generation actually used on Island [MWh]

	percent_curtailed = np.sum(curtail)/ np.sum(totalSupply)
	percent_exported  = total_export    /((np.sum(totalSupply))*dt)
	percent_integrated= supply_int       /((np.sum(totalSupply))*dt)

	# costs
	if (costImports == None) or (GHGImports == None):
		LCOE = 0.0
		GHGi = 0.0
	else:
		cost_an = np.zeros(nSupplies+nStorages)
		GHG_an  = np.zeros(nSupplies+nStorages)   # in kg CO2e
		for j in range(nSupplies):
			cost_an[j], GHG_an[j] = supplies[j].getCost()
		for j in range(nStorages):
			cost_an[nSupplies+j], GHG_an[nSupplies+j] = storages[j].getCost()
			
		cost_import = total_import*costImports/1e6 + np.max(imported)*costCapacity/1e6
		GHG_import  = total_import*GHGImports  # units are kg/MWh * MWh = kg

		LCOE = 1e6 * (np.sum(cost_an) + cost_import) /( np.sum(adjustedLoad)*dt )
		GHGi = (np.sum(GHG_an) + GHG_import)/( np.sum(adjustedLoad)*dt )

	#TODO: write output to both console and text file, or excel sheet.

	print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
	print("Total generation:      {:8.1f} GWh".format(0.001*np.sum(totalSupply)*dt))
	print("Total annual load:     {:8.1f} GWh".format(0.001*np.sum(totalLoad)*dt))
	print("Load factor (pre-DR):  {:8.1f}%   ".format(100*np.mean(totalLoad)/np.max(totalLoad)))
	print("Load factor (post-DR): {:8.1f}%   ".format(100*np.mean(adjustedLoad)/np.max(adjustedLoad)))
	print("Integrated generation: {:8.1f} GWh".format(0.001*supply_int))
	print("                       {:8.1f}%   ".format(100*supply_int/(np.sum(supplies_used)*dt)))
	print("Peak export (MW):      {:8.1f} MW ".format(-np.min(imported)))
	print("Peak import (MW):      {:8.1f} MW ".format(np.max(imported)))
	print("Net import (GWh):      {:8.1f} GWh".format(0.001*np.sum(imported)*dt))
	print("Total (exports only):  {:8.1f} MWh".format(total_export))
	print("Total (imports only):  {:8.1f} MWh".format(total_import))
	print("Imports energy cost:   ${:8.1f}M".format(total_import*costImports/1e6))  
	print("Imports capacity cost: ${:8.1f}M".format(np.max(imported)*costCapacity/1e6))  
	print("Elec. import GHGs:     {:8.1f} tCO2e".format(GHG_import/1e3))  
	print("Local renewable energy:{:8.1f}%   ".format(100*supply_int/(np.sum(final_load_with_storage)*dt)))
	print("LCOE:                  {:8.1f} $/MWh".format(LCOE))
	print("overall GHG intensity: {:8.1f} kgCO2e/MWh".format(GHGi))
	print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
	
	for j in range(nSupplies):
		print("CF {:8.1f}% ".format(100*np.mean(supplies_used[j,:])/np.max(supplies_used[j,:])))

	#TODO: make GHG inputs just in g/kWh

	# ------------------------------- write output file -------------------------------------

	ofile = open("ModelOutputInMW.csv", "w")

	# write headers

	ofile.write("Time")

	for j in range(nSupplies):
		ofile.write(","+supplies[j].name)

	for j in range(nLoads):
		ofile.write(","+loads[j].name)

	for j in range(nBEVs):
		ofile.write(","+BEVs[j].name)

	for j in range(nStorages):
		ofile.write(","+storages[j].name)
		
	ofile.write("\n")
		
	# write data

	for i in range(len(time)):

		ofile.write("{:8.2f}".format(time[i]))

		for j in range(nSupplies):
			ofile.write(",{:8.2f}".format(supplies[j].genTS[i]))

		for j in range(nLoads):
			ofile.write(",{:8.2f}".format(-loads[j].adjustedLoadTS[i]))

		for j in range(nBEVs):
			ofile.write(",{:8.2f}".format(BEV_supplies[j,i]))

		for j in range(nStorages):
			ofile.write(",{:8.2f}".format(storage_supplies[j,i]))  #SOCTS

		ofile.write("\n")

	ofile.close()

	# ---------------------------------- plotting, if enabled --------------------------------------

	if plots==1:	
		
		# Generation and Load Plots
		fig, ax = plt.subplots(4,1,sharex=True, figsize=[9,5.5], gridspec_kw={'height_ratios':[1.8,1.3,1,1]})
		
		# generation plot
		l1 = ax[0].plot(time, totalLoad, 'k',                           lw=1, label="original load")
		l2 = ax[0].plot(time, adjustedLoad, '--r',                      lw=1, label="adjusted load")
		l2 = ax[0].plot(time, load_with_charging, ':', color=(0.5,0,0), lw=1, label="load + charging")
		handles, labels = ax[0].get_legend_handles_labels()
		
		stacks = [supplies_used[j,:] for j in range(nSupplies)]+[
				  storage_supplies[k,:]*(storage_supplies[k,:]>0) for k in range(nStorages)]+[
				  BEV_supplies[k,:]*(BEV_supplies[k,:]>0) for k in range(nBEVs)]+[
				 imported*(imported > 0)]
				  
				  
		cmap1 = cm.get_cmap('summer', 12)
		cmap2 = cm.get_cmap('cool'  , 12)
				  
		colorlist = [cmap1(j/nSupplies) for j in range(nSupplies)]+[
					 cmap2(k/(nStorages+nBEVs)) for k in range(nStorages+nBEVs)]+[(0.7,0.7,0.7,1)]		  
				  
		stack_coll = ax[0].stackplot(time, stacks, lw=0, colors=colorlist)
		
		# make legend entries
		handles = handles+[Rectangle((0, 0), 1, 1, fc=pc.get_facecolor()[0]) for pc in stack_coll] # proxy rectangles
		labels = labels+[gen.name+" supply" for gen in supplies]+[
						 stor.name+" discharge" for stor in storages+BEVs]+["Imported power"]    # supply names
				
		ax[0].legend(handles, labels, loc='upper right')
		
		# load plot
		l2 = ax[1].plot(time, load_with_charging, ':', color=(0.5,0,0), lw=1, label="load + charging")
		handles, labels = ax[1].get_legend_handles_labels()
		
		stacks = [load.adjustedLoadTS for load in loads]+[
				  -storage_supplies[k,:]*(storage_supplies[k,:]<0) for k in range(nStorages)]+[
				  -BEV_supplies[k,:]*(BEV_supplies[k,:]<0) for k in range(nBEVs)]+[
				 -imported*(imported < 0)]
				  
				  
		cmap1 = cm.get_cmap('summer', 12)
		cmap2 = cm.get_cmap('cool'  , 12)
				  
		colorlist = [cmap1(j/nLoads) for j in range(nSupplies)]+[
					 cmap2(k/(nStorages+nBEVs)) for k in range(nStorages+nBEVs)]+[(0.7,0.7,0.7,1)]		  
				  
		stack_coll = ax[1].stackplot(time, stacks, lw=0, colors=colorlist)
		
		# make legend entries
		handles = handles+[Rectangle((0, 0), 1, 1, fc=pc.get_facecolor()[0]) for pc in stack_coll] # proxy rectangles
		labels = labels+[load.name+" load" for load in loads]+[
						 stor.name+" charging" for stor in storages+BEVs]+["Exported power"]    # supply names
				
		ax[1].legend(handles, labels, loc='upper right')
		
		
		# export plot
		balance_out = -imported #???
		
		ax[2].axhline(0,color=[0.5,0.5,0.5])
		ax[2].stackplot(time,balance_out*(balance_out<0), colors=[(0.7,0.7,0.7,1)], lw=0)
		ax[2].plot(time, balance_out, "k", lw=1, label="with storage")
		
		for j in range(nStorages):
			#ax[2].plot(storage_supplies[j,:], label=storages[j].name, lw=1.5, color=cmap2(j/(nStorages+nBEVs)))
			ax[3].plot(storage_SOC[j,:], label=storages[j].name, lw=1.2, color=cmap2(j/(nStorages+nBEVs)))
					
		for j in range(nBEVs):
			#ax[2].plot(BEV_supplies[j,:], label="EV: "+BEVs[j].name, lw=1.5, color=cmap2((j+nStorages)/(nStorages+nBEVs)))
			ax[3].plot(BEV_SOC[     j,:], label="EV: "+BEVs[j].name, lw=1.2, color=cmap2((j+nStorages)/(nStorages+nBEVs)))
			
		# note that SOC for variable storage is only the stored energy that is currently available/plugged-in
			
			
		ax[0].set_title("Generation, Demand, Import/Export, and Storage Over the Year")
		ax[0].set_ylabel("generation\n(MW)")
		ax[1].set_ylabel("load\n(MW)")
		ax[2].set_ylabel("export\n(MW)")
		ax[3].set_ylabel("storage\nSOC (MWh)")
		ax[3].set_xlabel("hours")
		ax[3].set_xlim([0,8760])
		ax[3].legend(loc='upper right')
		
		
		fig.tight_layout()
		fig.subplots_adjust(hspace=0.2)
		#fig.savefig("latest.png", bbox_inches=0, dpi=300)
		
		
		# BEV Plots
		if len(BEVS) > 0:
			fig, ax = plt.subplots(3,1,sharex=True, figsize=[9,5.5])#, gridspec_kw={'height_ratios':[1.8,1.3,1]})
			
			for j in range(nBEVs):
				ax[0].plot(BEVs[j].availability_fraction,lw=1.2, color=cmap2((j+nStorages)/(nStorages+nBEVs)),  label="EV: "+BEVs[j].name)
				ax[1].plot(BEVs[j].SOCTS,                lw=1.2, color=cmap2((j+nStorages)/(nStorages+nBEVs)),  label="EV: "+BEVs[j].name)
				ax[2].plot(BEVs[j].loadTS,               lw=1.2, color=cmap2((j+nStorages)/(nStorages+nBEVs)))
				ax[2].plot(-BEV_supplies[j,:],           lw=1.2, color=cmap2((j+nStorages)/(nStorages+nBEVs)), dashes=[1,2], 
						   label="EV: "+BEVs[j].name+" Consumption: {:6.0f} GWh".format(np.sum(-0.001*BEV_supplies[j,:])))
				
			ax[0].set_title("EV Fleet Behaviour Over the Year")
			ax[0].set_ylabel("Fraction of EVs \nplugged in")
			ax[1].set_ylabel("Stored energy \nplugged in (MWh)")
			ax[2].set_ylabel("EV fleet charging \npower (MW)")
			ax[2].set_xlabel("hours")
			ax[2].legend(loc='upper right')
			fig.tight_layout()


	#TODO: write output data somewhere




	plt.show()  # show whatever plots have been made

	
	
	
	
	