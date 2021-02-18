# TODO: Update to include an option to save the case
# TODO: Faster importing / data cleansing options
# TODO: Further work on work instruction required

import os
import sys
import load_est
# load_est = reload(load_est)
import load_est.psse as psse
# psse = reload(psse)
import load_est.constants as constants
import load_est.common_functions as common_functions
# constants = reload(constants)
import logging
import collections
import time
import numpy as np
import math
import dill
import pandas as pd
pd.options.display.width = 0

sys_path_PSSE = r'C:\Program Files (x86)\PTI\PSSE33\PSSBIN'  # or where else you find the psspy.pyc
sys.path.append(sys_path_PSSE)
os_path_PSSE = r'C:\Program Files (x86)\PTI\PSSE33\PSSBIN'  # or where else you find the psse.exe
os.environ['PATH'] += ';' + os_path_PSSE
os.environ['PATH'] += ';' + sys_path_PSSE
# noinspection PyUnresolvedReferences
import psspy  # noqa

# enables correct logging when functions called from GUI
logger = logging.getLogger(constants.Logging.logger_name)


class Station:

	def __init__(self, df, st_type):
		"""
		Station class init function to initialise station object properties
		:param pd.Dataframe() df: dataframe containing one station information from excel spreadsheet
		:param st_type: either 'BSP' or 'Primary'
		"""

		# reset index to ensure first row is row zero to be able to use iloc later
		df.reset_index(drop=True, inplace=True)
		# create a copy of the full data frame
		self.df = df.copy()
		# create a dataframe of the first row only
		df_fr = df.iloc[0].copy()

		# Initialise station type
		self.st_type = st_type

		# TODO: Initialise season load estimates
		self.spring_load=float # these are based on percentile values of provided values and are used to fill in missing season load values
		self.summer_load=float
		self.min_load=float

		# Initialise gsp col dictionary
		self.gsp = df_fr.iat[constants.XlFileConstants.gsp_col_no]# n: its 0, which gives the gsp name
		self.gsp_col = {df_fr.index[constants.XlFileConstants.gsp_col_no]: df_fr.iat[constants.XlFileConstants.gsp_col_no]}

		# Initialise name col dictionary
		self.name_val = df_fr.iat[constants.XlFileConstants.name_col_no]
		self.name = {df_fr.index[constants.XlFileConstants.name_col_no]: df_fr.iat[constants.XlFileConstants.name_col_no]}

		# Initialise nrn col dictionary
		self.nrn = {df_fr.index[constants.XlFileConstants.nrn_col_no]: df_fr.iat[constants.XlFileConstants.nrn_col_no]}

		# # Initialise growth_rate col dictionary
		# self.growth_rate_key_val = df_fr.iat[Constants.growth_rate_col]
		# self.growth_rate_key = {df_fr.index[Constants.growth_rate_col]: df_fr.iat[Constants.growth_rate_col]}

		# Initialise peak_mw col dictionary
		self.peak_mw_val = df_fr.iat[constants.XlFileConstants.peak_mw_col]
		self.peak_mw_dict = {df_fr.index[constants.XlFileConstants.peak_mw_col]: df_fr.iat[constants.XlFileConstants.peak_mw_col]}

		# Initialise power factor dictionary - default power factor of 1
		self.set_pf(1)

		self.gsp_diverse_forecast_dict = dict() # n: why is this defined here?
		self.gsp_aggregate_forecast_dict = dict()
		self.load_forecast_dict = dict()

		if self.st_type == constants.XlFileConstants.gsp_type: # is it's a gsp
			# extract gsp power factor
			gsp_pf = df.iat[constants.XlFileConstants.pf_cell_tuple]
			if not np.isnan(gsp_pf):
				self.set_pf(gsp_pf)
			self.gsp_scalable = True #

			# gsp_div_agg_df = df.loc[0:1, df.columns[Constants.load_forecast_col_range]]
			# gsp_div_agg_idx = ['Diverse', 'Aggregate']
			# gsp_div_agg_df.index = gsp_div_agg_idx
			# self.gsp_diverse_forecast_dict = gsp_div_agg_df.loc['Diverse'].to_dict()
			# self.gsp_aggregate_forecast_dict = gsp_div_agg_df.loc['Aggregate'].to_dict()

		else:

			self.idv_scalable = True #
			self.load_percentage = float #

		# Initialise name of upstream station as empty dictionary
		self.name_up = dict()

		self.load_forecast_dict = df_fr.iloc[constants.XlFileConstants.load_forecast_col_range].to_dict()# what is the key for this dictionary years?or column number sor??
		self.seasonal_percent_dict = df_fr.iloc[constants.XlFileConstants.seasonal_percent_col_range].to_dict()
		self.seasonal_percent_dict['Maximum Demand'] = 1 # 1 means 100%

		psse_busses_df = df.loc[0:1, df.columns[constants.XlFileConstants.psse_buses_col_range]] # intresting way of using.loc
		psse_idx = ['bus_no', 'pc']
		psse_busses_df.index = psse_idx
		self.psse_buses_dict = psse_busses_df.to_dict()

		# Initialise substation dictionary station as empty dictionary and number of substations to zero
		self.sub_stations_dict = dict()
		self.no_sub_stations = len(self.sub_stations_dict)

		self.load_forecast_diverse_fac = float

	def set_pf(self, pf):
		"""
		Function to set power factor of a station object
		:param pf: power factor value to set
		:return:
		"""
		self.pf = {constants.XlFileConstants.pf_str: pf}

	def add_sub_station(self, station_obj):
		"""
		Function to add a stations as a substation to a station object
		:param station_obj: station object to be added a sub station
		:return:
		"""
		# inherit gsp col and power factor from the upstream station
		station_obj.gsp_col = self.gsp_col
		station_obj.pf = self.pf

		# set the name of upstream station as the upstream station name
		station_obj.name_up = self.name

		# add the station object to the substation dictionary and update the number of substations
		self.sub_stations_dict.update({len(self.sub_stations_dict.keys()): station_obj})
		self.no_sub_stations = len(self.sub_stations_dict)

		return

	def calc_load_percentages(self):
		# this function calulates the percentage of loads each substation has
		year_list = self.df.columns[constants.XlFileConstants.load_forecast_col_range].to_list()
		year_list.sort()

		peak_mw_df = pd.DataFrame(columns=year_list)  #what does this do make an empty dataframe with years as column?

		for key, prim_stat in self.sub_stations_dict.iteritems():
			# for counter, yr in enumerate(year_list):
			# 	prim_stat.load_forecast_dict[yr] = \
			# 		prim_stat.peak_mw_val * Constants.growth_rate_dict[prim_stat.growth_rate_key_val] ** counter

			temp_df = pd.DataFrame([prim_stat.load_forecast_dict]) # what is prime stat laod forecast dict in terms of the excel sheet???
			temp_df.index = [prim_stat.name_val]
			peak_mw_df = peak_mw_df.append(temp_df)

		peak_mw_df.loc['Column_Total'] = peak_mw_df.sum(numeric_only=True, axis=0)

		self.load_forecast_dict = peak_mw_df.loc['Column_Total'].to_dict()
		self.load_forecast_diverse_fac = self.peak_mw_val / peak_mw_df.loc['Column_Total', year_list[0]].item() # what does item do here? also isnt this calsulated in the spreadsheet by dividing L5/L6?

		for key, prim_stat in self.sub_stations_dict.iteritems():

			prim_stat.load_percentage = peak_mw_df.loc[prim_stat.name_val, year_list[0]].item() / \
				peak_mw_df.loc['Column_Total', year_list[0]].item()

		return None

	def scalable_indv_check(self):

		# create temp dict to have just bus numbers not percentages also
		psse_buses_check_dict = dict()
		for key, sub_dict in self.psse_buses_dict.iteritems():
			psse_buses_check_dict[key] = sub_dict['bus_no']

		# concatenate relevant station properties, the following line first convert all dictionaries to dataframes then put them netx to each other to make a row dataframe
		df = pd.concat([
			pd.DataFrame([self.gsp_col]),
			pd.DataFrame([self.nrn]),
			pd.DataFrame([self.name]),
			pd.DataFrame([self.peak_mw_dict]),
			pd.DataFrame([self.pf]),
			pd.DataFrame([self.load_forecast_dict]),
			pd.DataFrame([psse_buses_check_dict])],
			axis=1,
		)# put all station data ina row dataframe with the name of the variables as headers and zero index

		check_cols = list()
		#todo:
		#estimate_cols=list()

		# check that MW Peak is a number greater than zero
		col_select = self.peak_mw_dict.keys()
		col_name = 'peak_mw' + '_pass'

		#col_est_name='peak_mw' + '_est'

		check_cols.append(col_name)
		#etimate_check_cols.append(col_est_name)

		df.loc[:, col_name] = True
		#df.loc[:, col_est_name] = False
		#this checks if the peak MW value is missing or if it's smaller than 0
		df.loc[df[df[col_select].le(0).any(1)].index, col_name] = False
		df.loc[df[df[col_select].isnull().any(1)].index, col_name] = False

		# todo: change the false for MW check back to true if it's a GSP as the MW of the GSP does not matter for the SSE load setting
		# if self.st_type == constants.XlFileConstants.gsp_type:
		# 	df.loc[:, col_name] = True
		# 	df.loc[:, col_est_name] = True

		# check load_forecast_dict is not null or negative
		col_select = self.load_forecast_dict.keys()
		col_name = 'load_forecast' + '_pass'
		#col_est_name = 'load_forecast' + '_est'

		check_cols.append(col_name)
		#etimate_check_cols.append(col_est_name)

		df.loc[:, col_name] = True
		# df.loc[:, col_est_name] = False

		df.loc[df[df[col_select].le(0).any(1)].index, col_name] = False
		df.loc[df[df[col_select].isnull().any(1)].index, col_name] = False
		# todo: if a new amend class in constant file is having its attribute of estimate as true, then change \
		# todo: the false to true (or jus the st column to true) and run a function which gets the df.loc[col_select] and estimate the missing values by interpolation then returns it
		# todo:


		# check all psse_buses_dict are not null
		col_select = psse_buses_check_dict.keys()
		col_name = 'psse_buses' + '_pass'
		check_cols.append(col_name)
		df.loc[:, col_name] = True
		df.loc[df[df[col_select].isnull().all(1)].index, col_name] = False #not sure if this looks at the bus names only or also their percentage values (its probably just the names)

		# final check
		col_select = check_cols
		col_name = 'Station_data_pass'
		df.loc[:, col_name] = False
		df.loc[df[df[col_select].all(1)].index, col_name] = True

		# final est check
		# col_select = check_est_cols
		# col_name = 'Station_data_est'
		# df.loc[:, col_name] = False
		# df.loc[df[df[col_select].all(1)].index, col_name] = True

		# output
		good_data = df['Station_data_pass'].item()
		#amend_data=df['Station_data_est'].item() # what does .item() do? is good_data a dataf frame?
		# todo: what does the line above do exactly? what is the good data form?
		# return good_data,amend_data
		return good_data

	def station_check(self):
		"""
		Function to return a row dataframe for a station object with new columns to check:
		- MW Peak is a number greater than zero
		- Each years load forecast is not null or negative
		- Each seasonal percentage is not null, negative or greater than 100%
		- That at least one PSSE bus exists for the load
		A final check column is added to check all above checks have been met
		:return pd.Dataframe: row dataframe of station object
		"""

		# create check dict to have just bus numbers not percentages
		# todo (percentage values are checked when applying load scaling)
		psse_buses_check_dict = dict()
		for key, sub_dict in self.psse_buses_dict.iteritems():
			psse_buses_check_dict[key] = sub_dict['bus_no']

		# concatenate relevant station properties
		df = pd.concat([
			pd.DataFrame([self.gsp_col]),
			pd.DataFrame([self.nrn]),
			pd.DataFrame([self.name]),
			pd.DataFrame([self.peak_mw_dict]),
			pd.DataFrame([self.pf]),
			pd.DataFrame([self.load_forecast_dict]),
			pd.DataFrame([self.seasonal_percent_dict]),
			pd.DataFrame([psse_buses_check_dict])],
			axis=1,
		)

		# initialise a list to store the new check columns names
		check_cols = list()
		#estimate_cols = list()

		# check that MW Peak is a number greater than zero
		col_select = self.peak_mw_dict.keys()
		col_name = 'peak_mw' + '_pass'
		check_cols.append(col_name)
		df.loc[:, col_name] = True
		df.loc[df[df[col_select].le(0).any(1)].index, col_name] = False
		df.loc[df[df[col_select].isnull().any(1)].index, col_name] = False


		# col_est_name = 'peak_mw' + '_est'
		# etimate_check_cols.append(col_est_name)
		# df.loc[:, col_est_name] = False
		# todo: change the false for MW check back to true if it's a GSP as the MW of the GSP does not matter for the SSE load setting and ew amend class in constant file is having its attribute of estimate as true
		# if self.st_type == constants.XlFileConstants.gsp_type and constants.amend.estimate==True:
		# 	df.loc[:, col_name] = True
		# 	df.loc[:, col_est_name] = True


		# check load_forecast_dict is not null or negative
		col_select = self.load_forecast_dict.keys()
		col_name = 'load_forecast' + '_pass'
		check_cols.append(col_name)
		df.loc[:, col_name] = True
		df.loc[df[df[col_select].le(0).any(1)].index, col_name] = False
		df.loc[df[df[col_select].isnull().any(1)].index, col_name] = False

		# col_est_name = 'load_forecast' + '_est'
		# etimate_check_cols.append(col_est_name)
		# df.loc[:, col_est_name] = False
		# todo: if a new amend class in constant file is having its attribute of estimate as true, then change \
		# todo: the check from false to true (or just the est column to true) and run a function which gets the df.loc[:,col_select] and estimate the missing values by interpolation then returns it
		# todo: and then the self.load_forecast_dict needs to be updated with the new values from the output dataframe .

		# check seasonal_percent is not null, negative or greater than 1
		col_select = self.seasonal_percent_dict.keys()
		col_name = 'seasonal_percent' + '_pass'
		# col_est_name = 'seasonal_forecast' + '_est'
		check_cols.append(col_name)
		df.loc[:, col_name] = True
		df.loc[df[df[col_select].le(0).any(1)].index, col_name] = False
		df.loc[df[df[col_select].gt(1).any(1)].index, col_name] = False
		df.loc[df[df[col_select].isnull().any(1)].index, col_name] = False


		# col_est_name = 'seasonal_percent' + '_est'
		# etimate_check_cols.append(col_est_name)
		# df.loc[:, col_est_name] = False
		#if constants.amend.estimate==True and df.loc[0,col_name]==False:
		# todo:





		# check all psse_buses_dict are not null
		col_select = psse_buses_check_dict.keys()
		col_name = 'psse_buses' + '_pass'
		check_cols.append(col_name)
		df.loc[:, col_name] = True
		df.loc[df[df[col_select].isnull().all(1)].index, col_name] = False

		# final check
		col_select = check_cols
		col_name = 'Station_data_pass'
		df.loc[:, col_name] = False
		df.loc[df[df[col_select].all(1)].index, col_name] = True

		# output
		good_data = df['Station_data_pass'].item()

		# populate good and bad dataframes
		if good_data:
			constants.XlFileConstants.good_data = constants.XlFileConstants.good_data.append(df)
		else:
			constants.XlFileConstants.bad_data = constants.XlFileConstants.bad_data.append(df)

		return good_data


def sse_load_xl_to_df(xl_filename, xl_ws_name, headers=True):
	"""
	Function to open and perform initial formatting on spreadsheet
	:param str() xl_filename: name of excel file 'name.xlsx'
	:param str() xl_ws_name: name of excel worksheet
	:param headers: where there is any data in row 0 of spreadsheet
	:return pd.Dataframe(): dataframe of worksheet specified
	"""

	if headers:
		h = 0
	else:
		h = None

	# import as dataframe and force to use xlsxwriter
	df = pd.read_excel(
		io=xl_filename,
		sheet_name=xl_ws_name,
		header=h,
	)
	# remove empty rows (i.e with all NaNs)
	df.dropna(
			axis=0,
			how='all',
			inplace=True
		)
	# remove empty columns (i.e with all NaNs)
	df.dropna(
		axis=1,
		how='all',
		inplace=True
	)
	# reset index
	df.reset_index(drop=True, inplace=True)

	return df


def extract_gsp_dfs(raw_df):
	"""
	Function to extract individual BSP dataframes
	:param pd.Dataframe() raw_df:
	:return dict(): network_df_dict Dictionary of dataframes with BSP name as key
	"""
	# extract the GSP row index to a list
	gsp_row_list = list(raw_df[raw_df.iloc[:, constants.XlFileConstants.gsp_col_no].str.contains(constants.XlFileConstants.gsp_type) == True].index)
	# add the last row index of df so last GSP is captured later on
	gsp_row_list.append(len(raw_df.index) + constants.XlFileConstants.row_separation)

	# extract the headers from the 1st GSP row and format
	headers = raw_df.iloc[gsp_row_list[0]].to_list()
	headers = map(lambda s: s.encode('ascii', 'replace'), headers)#replace the weird charachters by ?
	headers = map(lambda s: s.strip(), headers) # removes white space before and after the string
	headers = map(lambda s: s.replace('\n', ''), headers) # remove new line

	# set df columns
	raw_df.columns = headers

	net_df_dict = collections.OrderedDict()  #a a dictionary which iterates over the items while keeoing the sequence they were entered in
	# extract each BSP as a dataframe
	for i in xrange(0, len(gsp_row_list)-1):

		# extract BSP name
		name_row = gsp_row_list[i] + 1
		temp_name = raw_df.iloc[name_row, constants.XlFileConstants.gsp_col_no]
		# print temp_name

		# extract the rows of the df dataframe and reset index
		temp_df = raw_df.iloc[name_row:gsp_row_list[i+1] - constants.XlFileConstants.row_separation].copy()
		temp_df.reset_index(drop=True, inplace=True)

		# add temp_df to dictionary
		net_df_dict[temp_name] = temp_df

	return net_df_dict


def create_stations(df_dict):
	"""
	Function to create station objects from dataframe
	:param dict() df_dict: Dictionary of dataframes
	:return dict(): Dictionary of station objects
	"""
	st_dict = collections.OrderedDict()

	for name, net in df_dict.iteritems():
		# todo why is the
		logger.info('Processing: ' + name)
		# print('Processing: ' + name)

		# extract GSP rows as individual dataframe
		gsp_idx = net[net[constants.XlFileConstants.gsp_type] == name].index.item()
		gsp_df = net.iloc[gsp_idx:gsp_idx + constants.XlFileConstants.bsp_no_rows]

		# create station object
		gsp_station = Station(gsp_df, constants.XlFileConstants.gsp_type)

		# check gsp_station row
		# only add GSP if passes row check
		if gsp_station.station_check():  #here another station check can be done to add substations even if it does not pass with the suggested values.

			# create new dataframe without the bsp rows
			# todo bit hard coded
			prim_temp_df = net.loc[net.index[4:]]
			#test=net.iloc[4:]
			# step through prim_temp_df in 3 rows at a time
			for b in xrange(0, len(prim_temp_df.index), constants.XlFileConstants.prim_no_rows):
				prim_df = prim_temp_df.iloc[b: b + constants.XlFileConstants.prim_no_rows]
				prim_station = Station(prim_df, constants.XlFileConstants.primary_type)
				# if primary passes station check add to GSP
				if prim_station.station_check():
					gsp_station.add_sub_station(prim_station)
				# if primary passes gsp scalable check add to GSP
				elif prim_station.scalable_indv_check():
					gsp_station.add_sub_station(prim_station)
					prim_station.idv_scalable = False
				# else do not add the primary station to the gsp
				else:
					gsp_station.gsp_scalable = False
					del prim_station

			if gsp_station.gsp_scalable:
				# If the GSP is scalable ie all primaries pass checks - calculate forecast loads
				gsp_station.calc_load_percentages()

			# finally add station to station dictionary
			st_dict.update({len(st_dict.keys()): gsp_station})
		else:
			del gsp_station

	# for num, gsp in st_dict.iteritems():
	# 	print (gsp.gsp_col.values())

	return st_dict


def exp_stations_to_excel(st_dict):
	"""
	Function to loop through a dictionary of station objects and create dataframe
	:param dict() st_dict: dictionary of station class objects
	:return: dataframe of all stations and substations
	"""

	# initialise dataframe df
	df = pd.DataFrame()

	# loop through each station object in dictionary
	for key, station in st_dict.iteritems():

		# call station_df_row() to create station dataframe row and append to df
		df = df.append(station.station_df_row())  # where is station_df_row defined?

		# if the station has substations loop through and append station dataframe ro to df
		if station.no_sub_stations > 0:
			for idx, sub_station in station.sub_stations_dict.iteritems():
				df = df.append(sub_station.station_df_row())

	# reset dataframe index
	df.reset_index(inplace=True)

	return df


def scale_all_loads(df_load_values, year=str(), season=str(), diverse=False, zone=tuple(), gsp=tuple()):
	"""
	Function to update station loads
	TODO: @NS - This function should be moved within load_est module

	:param pd.DataFrame df_load_values:  DataFrame with all of the load values needed for scaling
	:param str() year:  Year loads should be scaled for
	:param str() season:  Season years should be scaled for
	TODO: @NS - I've added a new input below for selecting whether to use aggregate or diverse values
	:param bool diverse:  Whether to scale for the diversified or aggregate load values
	TODO: @NS - I've added new inputs that allow the user to select specific GSPs or Zones
	:param tuple zone:  Populated with a list of zones if the user selects specific zones to consider
	:param tuple gsp:  Populated with a list of gsps if the user selects specific zones to consider
	:return:
	"""

	# # Check if PSSE is running and if so retrieve list of selected busbars, else return empty list
	# psse_con = load_est.psse.PsseControl()
	# psse_con.load_data_case(pth_sav=psse_case)
	# psse_con.change_output(destination=False)

	loads = psse.LoadData()
	loads_df = loads.df.set_index('NUMBER')
	loads_list = map(int, list(loads_df.index))

	# TODO: @NS - The following code needs to be changed as follows
	# TODO: @NS... 1. Apply filter to select only Primary substations since not PSSe loads modelled explicitly for GSPs
	# TODO: @NS... 2. Apply filter to select only GSPs or Zones that have been provided as an input
	# TODO: @NS... 3. Loop through remaining DataFrame and set load values for the appropriate year (psuedo code below)

	# TODO: @NS - Note, the following is just a suggested example and it maybe worth breaking up into smaller functions
	for idx, substation in df_filtered.iterrows():
		# TODO: @NS - Calculate substation P and Q
		p = substation[year] * substation[season] * substation[pf]
		q =

		# Loop through each busbar in substation
		# TODO: @NS - Need some way to be able to loop through all of the busbars associated with a substation
		for bus in substations[buses]:

			p_bus = p * bus[proportion]
			q_bus = q * bus[proportion]

			# Confirm busbar exists in model and if so update, otherwise alert user
			if bus[busbar_number] in loads_list:
				# Set the load value for this busbar
				# TODO: @NS - Should make use of the psse.LoadData() class rather than this function, will come back to that
				ierr = psspy.load_chng_5(
					i=bus[busbar_number],
					id=constants.Loads.default_id,
					realar1=p_bus,  # P load MW
					realar2=q_bus	# Q load Mvar
				)
				logger.debug((
								 'PSSE busbar {} associated with {} updated with a new P/Q value of {:.2f}/{:.2f}'
							 ).format(bus[busbar_number], substation[name], p_bus, q_bus)
							 )
			else:
				logger.error((
								 'PSSE busbar {} associated with {} not found in PSSe model'
							 ).format(bus[busbar_number], substation[name])
							 )

			# TODO: @NS - We will also need to disconnect any other loads modelled at this busbar already in the model
			# TODO: @NS... to ensure the total load numbers are reaonsable

	# for station_no, station in constants.General.station_dict.iteritems():
	#
	# 	for i in xrange(0, station.no_sub_stations):
	#
	# 		sub_station = station.sub_stations_dict[i]
	#
	# 		if sub_station.idv_scalable:
	#
	# 			for sub_num, sub_psse_bus in sub_station.psse_buses_dict.iteritems():
	#
	# 				if math.isnan(sub_psse_bus['bus_no']):
	# 					continue
	# 				if sub_psse_bus['bus_no'] in loads_list:
	# 					p = sub_station.load_forecast_dict[year] * \
	# 						sub_station.seasonal_percent_dict[season] * \
	# 						sub_station.pf['p.f'] * \
	# 						sub_psse_bus['pc']
	#
	# 					q = p * math.tan(math.acos(sub_station.pf['p.f']))
	# 					# loads at the substations buses
	#
	# 					ierr = psspy.load_chng_5(
	# 						i=sub_psse_bus['bus_no'],
	# 						id=loads_df.loc[sub_psse_bus['bus_no'], 'ID'],
	# 						realar1=p,  # P load MW
	# 						realar2=q)  # Q load MW
	# 					break
	# 				else:
	# 					logger.info('Bus number ' + str(sub_psse_bus) + ' not in PSSE sav case')

	return None


def scale_all_gens(pc=float()):
	"""
	Function to Scale all generation in case by scaling percentage
	TODO: @NS - This function should be moved within load_est module
	:param pc:
	:return:
	"""

	machine_data = psse.MachineData()

	for gen in machine_data.df.itertuples():

		if gen.RPOS != 999:

			p = gen.PQGEN.real * float(pc) / float(100)
			q = gen.PQGEN.imag * float(pc) / float(100)

			ierr = psspy.machine_chng_2(
				i=gen.NUMBER,
				id=gen.ID,
				realar1=p,  # MW set point
				realar2=q,  # MVAr set point
				)

	return None


def set_growth_const(df):
	"""
	Function to set growth rate constant (dict)
	:param pd.Dataframe() df:
	:return:
	"""
	# todo way to hardcoded - can we suggest to SSE to create a better worksheet here???
	new_df = df[[1, 3]].copy()
	new_df = new_df.iloc[2:6]
	new_df.columns = ['key', 'growth_rate']
	new_df.set_index('key', drop=True, inplace=True)
	temp_dict = new_df.to_dict()

	constants.XlFileConstants.growth_rate_dict = temp_dict['growth_rate']

	return None


def process_load_estimates_xl(xl_path):
	"""
	Function to load SSE load estimates excel file, perform data checks and create dictionary of station objects
	:param xl_path: The path of the SSE load estimates file
	:return:
	"""

	# load worksheet into dataframe
	raw_dataframe = sse_load_xl_to_df(xl_path, constants.XlFileConstants.excel_ws_name)
	# todo: add a data_frame_approach function (xl_path) which gives out df_mofidied and saves the excel file witgh all data comparison/good/bad data
	# todo: function to load from dill
	# todo: add another function which gets raw_dataframe and modified data frame plus a variable to choose between aggregate or
	#  real powers and gives out a df with the same format as raw_dataframe but with modifed values.

	# extract individual GSPs
	network_df_dict = extract_gsp_dfs(raw_dataframe)

	# create station dictionary
	station_dict = create_stations(network_df_dict)


	# Check params folder exists to store log files in and if not create appropriate folders
	params_folder = os.path.join(constants.General.cur_path, constants.XlFileConstants.params_folder) # what is cur_path??
	if not os.path.exists(params_folder):
		os.mkdir(params_folder)

	# create new excel file name in the example folder and write the good data and bad data to separate sheets


	# Create a dictionary to save the params after reading in excel file
	## params_dict = create_params_pkl(station_dict, xl_path)
	# todo: write a function to set the parameter from modified df
	#set_params_constants(params_dict)


def create_params_pkl(station_dict, xl_path):
	"""
	Function to create params dict and save it to a pkl/dill file
	:param station_dict: dictionary of station objects
	:param xl_path: path of the SSE load estimates excel file
	:return: params_dict
	"""
	params_dict = dict()
	params_dict[constants.SavedParamsStrings.station_dict_str] = station_dict
	params_dict[constants.SavedParamsStrings.xl_file_name] = os.path.basename(xl_path)
	params_dict[constants.SavedParamsStrings.loads_complete_str] = constants.XlFileConstants.bad_data.empty # this is to check whether all data are available (as in all checks passes)

	# use the keys from the first station object for years list and demand scaling list
	params_dict[constants.SavedParamsStrings.years_list_str] = \
		sorted(station_dict[0].load_forecast_dict.keys())
	params_dict[constants.SavedParamsStrings.demand_scaling_list_str] = \
		sorted(station_dict[0].seasonal_percent_dict.keys())

	# generate a list of scalable GSP
	temp_list = list()
	for key, gsp in station_dict.iteritems():
		if gsp.gsp_scalable:
			temp_list.append(gsp.gsp)
	params_dict[constants.SavedParamsStrings.scalable_GSP_list_str] = sorted(temp_list)

	# save a pickle/dill file of the params dict to speed up processing later
	with open(os.path.join(
			constants.General.cur_path,
			constants.XlFileConstants.params_folder,
			constants.SavedParamsStrings.params_file_name),
			'wb') as f:
		dill.dump(params_dict, f)

	return params_dict


def set_params_constants(params_dict):
	"""
	FUunction to set the General constatns from a params dictionary
	:param params_dict: dictionary containing the params needed for the GUI
	params_dict = {
	scalable_GSP_list: list of GSPs that are scalable,
	years_list: list of years for forecast loads (from loaded excel file),
	demand_scaling_list: list of options for demand scaling (from excel file =  Maximum Demand,
	Minimum Demand, Spring/Autumn, Summer),
	station_dict: dictionary of station class objects,
	loads_complete: boolean if all loads are error free or not
	xl_file_name:  the excel file name used to create the station_dict

	:return:
	"""
	constants.General.params_dict = params_dict
	constants.General.scalable_GSP_list = \
		constants.General.params_dict[constants.SavedParamsStrings.scalable_GSP_list_str]
	constants.General.years_list = \
		constants.General.params_dict[constants.SavedParamsStrings.years_list_str]
	constants.General.demand_scaling_list = \
		constants.General.params_dict[constants.SavedParamsStrings.demand_scaling_list_str]
	constants.General.station_dict = \
		constants.General.params_dict[constants.SavedParamsStrings.station_dict_str]
	constants.General.loads_complete = \
		constants.General.params_dict[constants.SavedParamsStrings.loads_complete_str]
	constants.General.xl_file_name = \
		constants.General.params_dict[constants.SavedParamsStrings.xl_file_name]


if __name__ == '__main__':

	"""
		This is the main block of code that will be run if this script is run directly
	"""

	# Time stamp for performance checking
	t0 = time.time()

	# todo does BKDY represent anything here?
	# Produce unique identifier for logger
	uid = 'BKDY_{}'.format(time.strftime('%Y%m%d_%H%M%S'))

	# Check temp folder exists to store log files in and if not create appropriate folders
	script_path = os.path.realpath(__file__)
	script_folder = os.path.dirname(script_path)
	temp_folder = os.path.join(script_folder, 'temp')
	if not os.path.exists(temp_folder):
		os.mkdir(temp_folder)
	logger = load_est.Logger(pth_logs=temp_folder, uid=uid, debug=constants.DEBUG_MODE)

	# Produce initial log messages and decorate appropriately
	psse_con = load_est.psse.PsseControl()
	logger.log_colouring(run_in_psse=psse_con.run_in_psse)

	constants.General.cur_path = os.path.dirname(__file__)

	# if there is a params file load this and set constants
	params_file = os.path.join(
		constants.General.cur_path,
		constants.XlFileConstants.params_folder,
		constants.SavedParamsStrings.params_file_name
	)
	if os.path.exists(params_file):
		with open(params_file, 'rb') as f:
			constants.General.params_dict = dill.load(f)

		# update constants from params_dict
		set_params_constants(constants.General.params_dict)

	# todo this is probably no needed when running form PSSE
	# init_psse = psse.InitialisePsspy()
	# init_psse.initialise_psse()

	gui = load_est.gui.MainGUI()

	logging.shutdown()

	print('finished')




