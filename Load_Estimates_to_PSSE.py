import os
import sys
import load_est
import load_est.psse as psse
import load_est.constants as constants
import collections
import time
import numpy as np
import pandas as pd
import math
pd.options.display.width = 0

sys_path_PSSE = r'C:\Program Files (x86)\PTI\PSSE34\PSSPY27'  # or where else you find the psspy.pyc
sys.path.append(sys_path_PSSE)
os_path_PSSE = r'C:\Program Files (x86)\PTI\PSSE34\PSSBIN'  # or where else you find the psse.exe
os.environ['PATH'] += ';' + os_path_PSSE
os.environ['PATH'] += ';' + sys_path_PSSE
import psspy


class Station:

	# todo write function to change PSSE load

	def __init__(self, df, st_type):
		"""
		Station class init function to initialise station object properties
		:param pd.Dataframe() df: dataframe containing one station information from excel spreadsheet
		:param st_type: either 'BSP' or 'Primary'
		"""

		# Initialise station type
		self.st_type = st_type

		# Initialise gsp col dictionary
		self.gsp_col = {df.index[Constants.gsp_col_no]: df.iat[Constants.gsp_col_no]}

		# Initialise name col dictionary
		self.name = {df.index[Constants.name_col_no]: df.iat[Constants.name_col_no]}

		# Initialise nrn col dictionary
		self.nrn = {df.index[Constants.nrn_col_no]: df.iat[Constants.nrn_col_no]}

		# Initialise growth_rate col dictionary
		self.growth_rate = {df.index[Constants.growth_rate_col]: df.iat[Constants.growth_rate_col]}

		# Initialise peak_me col dictionary
		self.peak_mw = {df.index[Constants.peak_mw_col]: df.iat[Constants.peak_mw_col]}

		# Initialise power factor dictionary - default power factor of 1
		self.pf = {Constants.pf_str: 1}

		# Initialise name of upstream station as empty dictionary
		self.name_up = dict()

		# Initialise name and forecasting variables dependent on the station type (st_type)
		self.load_forecast_dict = df.iloc[Constants.load_forecast_col_range].to_dict()
		self.seasonal_percent_dict = df.iloc[Constants.seasonal_percent_col_range].to_dict()
		self.psse_buses_dict = df.iloc[Constants.psse_buses_col_range].to_dict()

		# Initialise substation dictionary station as empty dictionary and number of substations to zero
		self.sub_stations_dict = dict()
		self.no_sub_stations = len(self.sub_stations_dict)

	def set_pf(self, pf):
		"""
		Function to set power factor of a station object
		:param pf: power factor value to set
		:return:
		"""
		self.pf = {Constants.pf_str: pf}

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

		st_added = False
		if station_obj.station_check_df_row():
			# add the station object to the substation dictionary and update the number of substations
			self.sub_stations_dict.update({len(self.sub_stations_dict.keys()): station_obj})
			self.no_sub_stations = len(self.sub_stations_dict)
			st_added = True

		return st_added

	def station_check_df_row(self):
		"""
		Function to return a row dataframe for a station object
		:return pd.Dataframe: row dataframe of station object
		"""

		# concatenate relevant station properties
		df = pd.concat([
			pd.DataFrame([self.gsp_col]),
			pd.DataFrame([self.nrn]),
			pd.DataFrame([self.name]),
			pd.DataFrame([self.peak_mw]),
			pd.DataFrame([self.pf]),
			pd.DataFrame([self.growth_rate]),
			pd.DataFrame([self.load_forecast_dict]),
			pd.DataFrame([self.seasonal_percent_dict]),
			pd.DataFrame([self.psse_buses_dict])],
			axis=1,
		)

		check_cols = list()

		# check load_forecast_dict is not null or zero
		col_select = self.load_forecast_dict.keys()
		col_name = 'load_forecast' + '_pass'
		check_cols.append(col_name)
		df.loc[:, col_name] = True
		df.loc[df[df[col_select].le(0).any(1)].index, col_name] = False
		df.loc[df[df[col_select].isnull().any(1)].index, col_name] = False

		# check seasonal_percent_dict is not null or zero
		col_select = self.seasonal_percent_dict.keys()
		col_name = 'seasonal_percent' + '_pass'
		check_cols.append(col_name)
		df.loc[:, col_name] = True
		df.loc[df[df[col_select].le(0).any(1)].index, col_name] = False
		df.loc[df[df[col_select].isnull().any(1)].index, col_name] = False

		# check all psse_buses_dict are not null
		col_select = self.psse_buses_dict.keys()
		col_name = 'psse_buses' + '_pass'
		check_cols.append(col_name)
		df.loc[:, col_name] = True
		df.loc[df[df[col_select].isnull().all(1)].index, col_name] = False

		# final check
		col_select = check_cols
		col_name = 'Load data_pass'
		df.loc[:, col_name] = False
		df.loc[df[df[col_select].all(1)].index, col_name] = True

		# output
		good_data = df['Load data_pass'].item()

		# if in debug mode add to dataframes
		if Constants.DEBUG:
			if good_data:
				Constants.good_data = Constants.good_data.append(df)
			else:
				Constants.bad_data = Constants.bad_data.append(df)

		return good_data


class Constants:

	DEBUG = 0

	# define good data and bad data dataframes
	good_data = pd.DataFrame()
	bad_data = pd.DataFrame()

	# define columns from spreadsheet
	gsp_col_no = 0
	nrn_col_no = 1
	name_col_no = 2

	# define BPS and primary number of rows
	bsp_no_rows = 4
	prim_no_rows = 3

	# the number of rows between the end of a GSP and the next GSP (not including row with 'Average Cold Spell (ACS))
	row_separation = 1

	# define station type string
	gsp_str = 'GSP'
	bsp_type = 'BSP'
	primary_type = 'PRIMARY'
	pf_str = 'p.f'

	# define the column ranges of interest
	peak_mw_col = 7
	growth_rate_col = 10
	load_forecast_col_range = range(11, 25)
	seasonal_percent_col_range = range(26, 29)
	psse_buses_col_range = range(29, 37)

	# define cell on interest for pf
	pf_cell_tuple = (3, 7)

	def __init__(self):
		pass


def sse_load_xl_to_df(xl_filename, xl_ws_name):
	"""
	Function to open and perform initial formatting on spreadsheet
	:param str() xl_filename: name of excel file 'name.xlsx'
	:param str() xl_ws_name: name of excel worksheet
	:return pd.Dataframe(): dataframe of worksheet specified
	"""

	# import as dataframe
	df = pd.read_excel(
		io=xl_filename,
		sheet_name=xl_ws_name
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


def extract_bsp_dfs(raw_df):
	"""
	Function to extract individual BSP dataframes
	:param pd.Dataframe() raw_df:
	:return dict(): network_df_dict Dictionary of dataframes with BSP name as key
	"""
	# extract the GSP row index to a list
	gsp_row_list = list(raw_df[raw_df.iloc[:, Constants.gsp_col_no].str.contains(Constants.gsp_str) == True].index)
	# add the last row index of df so last GSP is captured later on
	gsp_row_list.append(len(raw_df.index) + Constants.row_separation)

	# extract the headers from the 1st GSP row and format
	headers = raw_df.iloc[gsp_row_list[0]].to_list()
	headers = map(lambda s: s.encode('ascii', 'replace'), headers)
	headers = map(lambda s: s.strip(), headers)
	headers = map(lambda s: s.replace('\n', ''), headers)

	# set df columns
	raw_df.columns = headers

	net_df_dict = collections.OrderedDict()
	# extract each BSP as a dataframe
	for i in xrange(0, len(gsp_row_list)-1):

		# extract BSP name
		name_row = gsp_row_list[i] + 1
		temp_name = raw_df.iloc[name_row, Constants.gsp_col_no]
		# print temp_name

		# extract the rows of the df dataframe and reset index
		temp_df = raw_df.iloc[name_row:gsp_row_list[i+1] - Constants.row_separation].copy()
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
		logger.info('Processing: ' + name)

		bsp_idx = net[net[Constants.gsp_str] == name].index.item()
		bsp_df = net.iloc[bsp_idx:bsp_idx + Constants.bsp_no_rows]

		# create station object
		bsp_station = Station(bsp_df.iloc[bsp_idx], Constants.bsp_type)
		# extract bsp power factor
		bsp_pf = bsp_df.iat[Constants.pf_cell_tuple]
		if not np.isnan(bsp_pf):
			bsp_station.set_pf(bsp_pf)
			# check bsp_station row
		bsb_station_pass = bsp_station.station_check_df_row()

		# create new dataframe without the bsp rows
		# todo bit hard coded
		prim_temp_df = net.loc[net.index[4:]]

		# step through prim_temp_df in 3 rows at a time
		for b in xrange(0, len(prim_temp_df.index), Constants.prim_no_rows):
			prim_station = Station(prim_temp_df.iloc[b], Constants.primary_type)
			# only add primary if passes row check
			station_added = bsp_station.add_sub_station(prim_station)
			if not station_added:
				del prim_station

		# only add BSP if passes row check
		if bsb_station_pass:
			st_dict.update({len(st_dict.keys()): bsp_station})
		else:
			del bsp_station

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
		df = df.append(station.station_df_row())

		# if the station has substations loop through and append station dataframe ro to df
		if station.no_sub_stations > 0:
			for idx, sub_station in station.sub_stations_dict.iteritems():
				df = df.append(sub_station.station_df_row())

	# reset dataframe index
	df.reset_index(inplace=True)

	return df


def update_loads(psse_case, station_dict, year=str(), season=str()):
	"""
	Function to update station loads
	:param dict() station_dict: dictionary of station objects
	:param str() year:
	:param str() season:
	:return:
	"""

	# Check if PSSE is running and if so retrieve list of selected busbars, else return empty list
	psse_con = load_est.psse.PsseControl()
	psse_con.load_data_case(pth_sav=psse_case)

	loads = psse.LoadData()
	loads_df = loads.df.set_index('NUMBER')

	loads_list = map(int, list(loads_df.index))

	for station_no, station in station_dict.iteritems():

		for num, psse_bus in station.psse_buses_dict.iteritems():

			p = station.load_forecast_dict[year] * station.seasonal_percent_dict[season]
			q = p * math.tan(math.acos(float(station.pf['p.f'])))
			# if there is load at the station bus
			if psse_bus in loads_list:
				ierr = psspy.load_chng_5(
					ibus=psse_bus,
					id=loads_df.loc[psse_bus, 'ID'],
					realar1=p,  # P load MW
					realar2=q)  # Q load MW

		for i in xrange(0, station.no_sub_stations):

			sub_station = station.sub_stations_dict[i]

			for sub_num, sub_psse_bus in sub_station.psse_buses_dict.iteritems():

				if sub_psse_bus is np.nan:
					continue
				if sub_psse_bus in loads_list:
					p = sub_station.load_forecast_dict[year] * sub_station.seasonal_percent_dict[season]
					q = p * math.tan(math.acos(float(station.pf['p.f'])))
					# loads at the substations buses

					ierr = psspy.load_chng_5(
						ibus=sub_psse_bus,
						id=loads_df.loc[sub_psse_bus, 'ID'],
						realar1=p,  # P load MW
						realar2=q)  # Q load MW
					break

				else:
					logger.info('Bus number ' + str(sub_psse_bus) + ' not in PSSE sav case')

	return None


if __name__ == '__main__':

	"""
		This is the main block of code that will be run if this script is run directly
	"""
	Constants.DEBUG = 1

	# Time stamp for performance checking
	t0 = time.time()

	# Produce unique identifier for logger
	uid = 'BKDY_{}'.format(time.strftime('%Y%m%d_%H%M%S'))

	# Check temp folder exists to store log files in and if not create appropriate folders
	script_path = os.path.realpath(__file__)
	script_folder = os.path.dirname(script_path)
	temp_folder = os.path.join(script_folder, 'temp')
	if not os.path.exists(temp_folder):
		os.mkdir(temp_folder)
	logger = load_est.Logger(pth_logs=temp_folder, uid=uid, debug=constants.DEBUG_MODE)

	# current project path
	cur_path = os.path.dirname(__file__)
	example_folder = r'load_est\test_files'
	file_name = r'SHEPD 2018 LTDS Winter Peak.sav'
	file_path = os.path.join(cur_path, example_folder, file_name)

	init_psse = psse.InitialisePsspy()
	init_psse.initialise_psse()

	psse_con = load_est.psse.PsseControl()

	# Produce initial log messages and decorate appropriately
	logger.log_colouring(run_in_psse=psse_con.run_in_psse)

	# Run main study
	# logger.info('Study started')

	# ------------------------------------------------------------------------------------------------------------------
	# workbook to open
	excel_filename = r'2019-20 SHEPD Load Estimates - v4.xlsx'
	# worksheet to open
	excel_ws_name = 'MASTER Based on SubstationLoad'
	file_path = os.path.join(cur_path, example_folder, excel_filename)

	# load worksheet into dataframe
	raw_dataframe = sse_load_xl_to_df(file_path, excel_ws_name)

	# extract individual BSPs
	network_df_dict = extract_bsp_dfs(raw_dataframe)

	# create station dictionary
	station_dict = create_stations(network_df_dict)

	if Constants.DEBUG:
		file_name = r'raw.xlsx'
		file_path = os.path.join(cur_path, example_folder, file_name)

		with pd.ExcelWriter(file_path) as writer:
			Constants.good_data.to_excel(writer, sheet_name='Complete Load Data')
			Constants.bad_data.to_excel(writer, sheet_name='Missing Load Data')
			worksheet1 = writer.sheets['Complete Load Data']
			worksheet1.set_tab_color('green')
			worksheet2 = writer.sheets['Missing Load Data']
			worksheet2.set_tab_color('red')

	constants.GUI.year_list = sorted(station_dict[0].load_forecast_dict.keys())
	constants.GUI.season_list = sorted(station_dict[0].seasonal_percent_dict.keys())
	constants.GUI.station_dict = station_dict

	gui = load_est.gui.MainGUI()

