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

xfilename='2019-20 SHEPD Load Estimates - v6.xlsx'
foldername='C:\Users\NegarShams\PycharmProjects\JK7938_SHEPD_LoadEstimates\load_est\test_files'
Sheetname='MASTER Based on SubstationLoad'
#fullname=os.path.join(
			# foldername,
			# xfilename,
			# # )
curpath = os.path.dirname(__file__)

fullname = os.path.join(curpath,'test_files',xfilename)

#DF = pd.read_csv('T2.csv')


DfX = pd.read_excel(
		io=fullname,
		sheet_name=Sheetname,
		header=0)  # by changing the header number the row from which the excel file is read changes if it's set as None then the whole excel is exported to df with the first row of excel being first row of data not the headr

# remove empty rows (i.e with all NaNs)
DfX.dropna(axis=0,how='all',inplace=True)
	# remove empty columns (i.e with all NaNs)
DfX.dropna(
		axis=1,
		how='all',
		inplace=True)
	# reset index
DfX.reset_index(drop=True, inplace=True)




def data_clean(raw_df):
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
	Merged=raw_df.iloc[0:2]
	#Merged3=raw_df[(raw_df['NRN'].notnull()) & raw_df['NRN']!='NRN']
	Merged3 = raw_df[raw_df['NRN'].notnull()]
	Merged3=Merged3[~Merged3['NRN'].isin(['NRN'])]


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
		# Merged=[Merged,temp_df]
		# Merged=pd.DataFrame
		Merged=Merged.append(temp_df)

	Merged=Merged.drop(Merged.index[0:1])
	# Merged.reset_index
	Merged1 = Merged[Merged['NRN'].notnull()]
	#return [net_df_dict,Merged3]
	return Merged3
# if __name__ == '__main__':

Cleaned_df = data_clean(DfX)
# reset index to ensure first row is row zero to be able to use iloc later
Cleaned_df.reset_index(drop=True, inplace=True)

Q=Cleaned_df[['Spring/Autumn']]
Q1=Cleaned_df[['Spring/Autumn','Summer']]
X=Cleaned_df.loc[Cleaned_df['Spring/Autumn']!=0,'Spring/Autumn']

Percen=np.percentile(X,75)
dicttest2=Cleaned_df