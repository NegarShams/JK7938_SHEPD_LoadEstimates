"""
#######################################################################################################################
###											Example DataFrame Comparison											###
###																													###
###		Code developed by David Mills (david.mills@PSCconsulting.com, +44 7899 984158) as part of PSC 		 		###
###		project JK7938 - SHEPD - studies and automation																###
###																													###
#######################################################################################################################
"""

# Generic Imports
import pandas as pd
import numpy as np
import unittest

# Unique imports
import common_functions as common

# General Constants
FILE_NAME_INPUT_1 = 'Processed Load Estimates_p_non_modified.xlsx'
FILE_NAME_INPUT_2 = 'Processed Load Estimates_p_modified.xlsx'
FILE_PTH_INPUT_1 = common.get_local_file_path(file_name=FILE_NAME_INPUT_1)
FILE_PTH_INPUT_2 = common.get_local_file_path(file_name=FILE_NAME_INPUT_2)

#
# FILE_NAME_OUTPUT = 'Example Comparison.xlsx'
FILE_NAME_OUTPUT = 'Example Comparison2.xlsx'
# FILE_PTH_OUTPUT = common.get_local_file_path(file_name=FILE_NAME_OUTPUT)
FILE_PTH_OUTPUT = common.get_local_file_path(file_name=FILE_NAME_OUTPUT)

# Engine to use when writing excel workbooks (XlsxWriter needed for formatting of tabs)
excel_engine = 'xlsxwriter'


# Functions
def produce_dataframe(dimensions, row_num):
	"""
		Produces a DataFrame where the row number is doubled
	:param int dimensions: Number of rows and columns
	:param int row_num: Row number that will be doubled
	:return pd.DataFrame df_raw:
	"""
	# Confirm that number of rows matches with dimensions and if not raise error message
	if row_num > dimensions:
		raise ValueError('Target row {} greater than dimensions {}'.format(row_num, dimensions))

	# Produces a list of lists of integers as generic data population
	generic_data = [range(0, dimensions) for _ in range(0, dimensions)]
	# Edit data to double a specific row
	generic_data[row_num] = [x * 2 for x in generic_data[row_num]]

	# Produces some column headers
	col_headers = ['Column {}'.format(x) for x in range(0, dimensions)]

	# Produce DataFrame
	df_raw = pd.DataFrame(data=generic_data, columns=col_headers)

	return df_raw


def highlight_diff(full_dataset, cells_to_highlight, color='yellow'):
	"""
		Small function to deal with highlighting the cells in the excel output that have changed from the original
		data set
	:param pd.DataFrame full_dataset:  Full dataframe
	:param pd.DataFrame cells_to_highlight:  DataFrame containing entries in only the cells which need highlighting
	:param str color:  Colour to highlight the cells
	"""
	# Set the attribute based on the colour input
	attr = 'background-color: {}'.format(color)
	print(attr)
	# Is difference
	is_diff = full_dataset == cells_to_highlight
	# Return a DataFrame with the style matching
	return pd.DataFrame(
		np.where(is_diff, attr, ''),
		index=full_dataset.index, columns=full_dataset.columns)


def compare_dataframes(df1, df2):
	"""
		Function compares two DataFrames of the same size and returns a dataframe of the same dimensions but with only
		the differences shown.  Keeps values in df2
	:param pd.DataFrame df1:  DataFrame 1
	:param pd.DataFrame df2:  DataFrame 2
	:return (pd.DataFrame, pd.DataFrame.Styled) (df_diff, df2_styled):  Different values between DataFrames and
																		original dataframe with changes highlighted
	"""

	# Confirm that DataFrames are the same dimensions otherwise raise error
	if df1.shape != df2.shape:
		raise ValueError('The two DataFrames provided are not the same dimensions ({} != {})'.format(df1.shape, df2.shape))

	# Keeps the values in df2 which are different to df1, all other values are set to nan
	local_df_diff = df2.where(df2 != df1)

	# Highlight the cells in df2 which have changed
	df2_styled = df2.style.apply(highlight_diff, cells_to_highlight=local_df_diff, axis=None)

	return local_df_diff, df2_styled


def write_dataframe(workbook, df, sheet_name, tab_color=None):
	"""
		Function deals with writing a DataFrame to new worksheet in excel whilst also formatting the worksheet tab
		color
	:param pd.ExcelWriter workbook:  Handle to workbook writer to use
	:param pd.DataFrame df:  Data to be written (if is a df.Styled) then will include highlighting and colour)
	:param str sheet_name:  Name to give worksheet
	:param str tab_color:  Color for tab
	:return: None
	"""
	# Create a new worksheet
	wksh = workbook.book.add_worksheet(name=sheet_name)

	# Have to add worksheet to Pandas list of worksheets
	# (https://stackoverflow.com/questions/32957441/putting-many-python-pandas-dataframes-to-one-excel-worksheet)
	workbook.sheets[sheet_name] = wksh

	# Only set tab_color if not None
	if tab_color:
		wksh.set_tab_color(tab_color)

	# Write DataFrame to excel worksheet
	df.to_excel(workbook, sheet_name=sheet_name)
	return None


def excel_data_comparison_maker(FILE_NAME_INPUT_1,FILE_NAME_INPUT_2,Bad_Data_Input_Name,Good_Data_Input_Name):
	"""
			Function compares two DataFrames of the same size and returns a dataframe of the same dimensions but with only
			the differences shown.  Keeps values in df2
		:param pd.DataFrame df1:  DataFrame 1
		:param pd.DataFrame df2:  DataFrame 2
		:return (pd.DataFrame, pd.DataFrame.Styled) (df_diff, df2_styled):  Different values between DataFrames and
																			original dataframe with changes highlighted
		"""
	# todo: modify the description

	FILE_PTH_INPUT_1 = common.get_local_file_path(file_name=FILE_NAME_INPUT_1)
	FILE_PTH_INPUT_2 = common.get_local_file_path(file_name=FILE_NAME_INPUT_2)
	FILE_PTH_INPUT_Bad_Data = common.get_local_file_path(file_name=Bad_Data_Input_Name)
	FILE_PTH_INPUT_Good_Data = common.get_local_file_path(file_name=Good_Data_Input_Name)

	FILE_NAME_OUTPUT = common.excel_file_names.data_comparison_excel_name
	FILE_PTH_OUTPUT = common.get_local_file_path(file_name=FILE_NAME_OUTPUT)

	# Engine to use when writing excel workbooks (XlsxWriter needed for formatting of tabs)
	excel_engine = 'xlsxwriter'


	df_main=common.import_excel(pth_load_est=FILE_PTH_INPUT_1)
	df_modified = common.import_excel(pth_load_est=FILE_PTH_INPUT_2)
	df_bad_data = common.import_excel(pth_load_est=FILE_PTH_INPUT_Bad_Data)
	df_good_data = common.import_excel(pth_load_est=FILE_PTH_INPUT_Good_Data)

	# Compare DataFrames to get differences and an updated df_modified with cells highlighted
	df_diff, df_modified_styled = compare_dataframes(df1=df_main, df2=df_modified)

	# Write DataFrames to excel workbook
	# Create an instance of excel
	with pd.ExcelWriter(path=FILE_PTH_OUTPUT, engine=excel_engine) as wkbk:
		# Write main data
		write_dataframe(workbook=wkbk, df=df_main, sheet_name='Raw Data')
		# Write modified data
		write_dataframe(workbook=wkbk, df=df_modified_styled, sheet_name='Modified Data', tab_color='green')
		# Write difference data
		write_dataframe(workbook=wkbk, df=df_diff, sheet_name='Difference Data', tab_color='blue')
		write_dataframe(workbook=wkbk, df=df_bad_data, sheet_name='Bad Data', tab_color='red')
		write_dataframe(workbook=wkbk, df=df_good_data, sheet_name='Good Data')
	k=1



if __name__ == '__main__':
	# Produce 2 different dataframes with the differences being on the row number
	#FILE_NAME_OUTPUT = 'Example Comparison2.xlsx'
	# FILE_PTH_OUTPUT = common.get_local_file_path(file_name=FILE_NAME_OUTPUT)
	#FILE_PTH_OUTPUT = common.get_local_file_path(file_name=FILE_NAME_OUTPUT)
	# Engine to use when writing excel workbooks (XlsxWriter needed for formatting of tabs)
	# FILE_NAME_INPUT_1 = common.excel_file_names.df_raw_excel_name
	# FILE_NAME_INPUT_2 = common.excel_file_names.df_modified_excel_name
	# Bad_Data_Input_Name = common.excel_file_names.bad_data_excel_name
	# FILE_NAME_INPUT_2 = 'Processed Load Estimates_p_modified.xlsx'
	FILE_NAME_INPUT_1 = common.excel_file_names.df_raw_excel_name
	FILE_NAME_INPUT_2 = common.excel_file_names.df_modified_excel_name
	Bad_Data_Input_Name = common.excel_file_names.bad_data_excel_name
	Good_Data_Input_Name = common.excel_file_names.good_data_excel_name

	excel_data_comparison_maker(FILE_NAME_INPUT_1=FILE_NAME_INPUT_1,FILE_NAME_INPUT_2=FILE_NAME_INPUT_2,\
								Bad_Data_Input_Name=Bad_Data_Input_Name,Good_Data_Input_Name=Good_Data_Input_Name)

	k=1


