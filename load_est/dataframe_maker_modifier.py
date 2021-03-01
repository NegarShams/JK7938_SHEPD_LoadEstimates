"""
#######################################################################################################################
###											Example DataFrame Processing											###
###																													###
###		Code developed by David Mills (david.mills@PSCconsulting.com, +44 7899 984158) as part of PSC 		 		###
###		project JK7938 - SHEPD - studies and automation																###
###																													###
#######################################################################################################################
"""

# Generic Imports
import pandas as pd
import numpy as np
# Unique imports
import common_functions as common
import data_comparison as comparison
import collections


# Functions
def assign_gsp(df_raw):
    """
		Function to assign a GSP to each Primary substation included within the dataset
	:param pd.DataFrame df_raw:  Raw dataframe
	:return pd.DataFrame df:  DataFrame with GSP headers added
	"""
    # Using Forward Fill to populate the GSP associated with each Primary down.
    # TODO: Potential risk here if the GSP name box is empty then an error will occur
    # Only forward fills where the GSP row has been identified by the sub_gsp column
    df_raw[common.Headers.gsp] = df_raw[common.Headers.gsp].where(df_raw[common.Headers.sub_gsp] == True).ffill()

    return df_raw


def determine_gsp_primary_flag(df_raw):
    """
		Determines whether a row contains GSP or Primary substation data
	:param pd.DataFrame df_raw:  Raw DataFrame
	:return pd.DataFrame df:  DataFrame with status flag identifying GSP or Primary substation data
	"""
    # If no entry in Name but an entry in GSP then assume GSP substation
    # TODO: There may be situations where this rule is not true
    # The following line determines all the rows for which this is True
    idx = (
            df_raw[common.Headers.name].isna() &  # Confirm that no Primary substation name entry exists
            ~df_raw[common.Headers.gsp].isna() &  # Confirm that a GSP name entry exists
            ~df_raw[common.Headers.voltage].isna()  # Confirm that a voltage ratio for the substation exists
    )
    # The following line sets those rows under the column sub_gsp to True
    df_raw.loc[idx, common.Headers.sub_gsp] = True

    # Find the situations where there is an entry in the Name but no entry in the GSP column but that there is an entry
    # in the NRN column as this determines that it is a Primary substation and not other data.
    # TODO: In case an NRN number is missing an alternative check is that there is a PSSE busbar number
    #  TODO: in column PSSE_bus 1 but this will require some further checking for other conditions
    idx = (
            ~df_raw[common.Headers.name].isna() &  # Confirm that an entry is included for Primary substation name
            df_raw[common.Headers.gsp].isna() &  # Confirm that no entry exists for the GSP substation name
            ~df_raw[common.Headers.nrn].isna()  # Confirm that an entry is included for the NRN number
    )
    # The following line sets those rows under the column sub_gsp to True
    df_raw.loc[idx, common.Headers.sub_primary] = True

    return df_raw


def bus_percentage_adder_modified(df_raw, fill):
    """
		adds the buses percentages as a new column
	:param pd.DataFrame df_raw:
	:param bool fill:  Whether to fill values or not
	:return pd.DataFrame df_raw:
	"""
    bus_list = filter(lambda local_x: local_x.startswith('PS'), df_raw.columns)
    percentage_list = [
        '{}_{}'.format(common.Headers.percentage, x) for x in
        bus_list
    ]  # makes a new list of column headers with the existing psse bus list

    bus_percentage_dict = collections.OrderedDict()
    percentage_miss_list = []
    df_raw[common.Headers.sum_percentages] = pd.Series([0 for _ in range(len(df_raw.index))])

    for b in percentage_list:  # this adds nan columns for the headers listed
        df_raw[b] = np.nan

    for b in bus_list:
        idx = (
                (
                        ~df_raw[common.Headers.sub_gsp].isna() |  # filters the rows that are gsp
                        ~df_raw[common.Headers.sub_primary].isna()
                ) &  # filters the rows that are primary
                ~df_raw[b].isna()
        )  # Confirm that the pss bus column has a name assigned to it

        bus_percentage_dict[b] = idx[
            idx == True].index  # makes a dictionary with psse bus column names as keys and with the index of the
    # rows with bus names as the values

    for n in range(0, len(bus_list)):
        for local_id in bus_percentage_dict[bus_list[n]]:
            if pd.isnull(df_raw.loc[local_id + 1, bus_list[n]]):
                df_raw.loc[local_id, percentage_list[n]] = 'missing'
                percentage_miss_list.append(local_id)
            else:
                df_raw.loc[local_id, percentage_list[n]] = df_raw.loc[
                    local_id + 1, bus_list[n]]  # the percentages of each bus is extracted and added as a new
                df_raw.loc[local_id, common.Headers.sum_percentages] = (
                        df_raw.loc[local_id, common.Headers.sum_percentages] + df_raw.loc[local_id, percentage_list[n]]
                )

    if fill:
        for i in percentage_miss_list:
            miss_row = df_raw.loc[i, percentage_list]
            miss_row_subset = miss_row.loc[miss_row[percentage_list] == 'missing'].index
            df_raw.loc[i, miss_row_subset] = (1 - df_raw.loc[i, common.Headers.sum_percentages]) / len(miss_row_subset)

    return df_raw


def extract_aggregate_demand(df_raw):
    """
		Extract the aggregate demand from the diversified demand for each GSP
	:param pd.DataFrame df_raw:
	:return pd.DataFrame df_raw:
	"""
    # Find the years being considered for the forecast
    forecast_years = common.adjust_years(headers_list=list(df_raw.columns))

    # Adjust the list to include a leading string value to identify this as a certain type of forecast (i.e aggregate)
    # TODO: Could make use of MultiIndex Pandas DataFrame Columns instead which would allow for more efficient filtering
    adjusted_list = ['{}_{}'.format(common.Headers.aggregate, x) for x in forecast_years]
    # Add columns to DataFrame with no files by adding in empty
    df_raw = pd.concat([df_raw, pd.DataFrame(columns=adjusted_list)], sort=False)

    # For columns which have been identified as GSP extract the aggregate demand from the row below and add to the GSP
    # row under the new sections for aggregate demand
    idx_gsp = df_raw[df_raw[common.Headers.sub_gsp] == True].index
    # loop through these index values for each GSP and get the aggregate values from the row below
    for idx in idx_gsp:
        df_raw.loc[idx, adjusted_list] = df_raw.loc[idx + 1, forecast_years].values

    return df_raw


def assign_pf(df_raw, fill):
    """
		Function to assign a GSP to each Primary substation included within the dataset
	:param pd.DataFrame df_raw:  Raw dataframe
	:return pd.DataFrame df:  DataFrame with GSP headers added
	"""
    idx_gsp = df_raw[df_raw[common.Headers.sub_gsp] == True].index
    idx_pf = df_raw[df_raw[common.Headers.pf_column_df_1] == 'Power Factor:'].index

    df_raw[common.Headers.PF] = np.nan

    df_pf = df_raw.loc[idx_pf, :]
    idx = (
            ~df_pf[common.Headers.pf_column_df].isna() &  #
            ~df_pf[common.Headers.pf_column_df].le(-1) &  #
            df_pf[common.Headers.pf_column_df].le(1))  #

    common.powerfactor.average_pf = df_pf.loc[idx, common.Headers.pf_column_df].mean()
    if fill:
        df_pf.loc[~idx, common.Headers.pf_column_df] = common.powerfactor.average_pf

    for j in range(0, len(idx_gsp)):
        df_raw.loc[idx_gsp[j], common.Headers.PF] = df_pf.loc[idx_pf[j], common.Headers.pf_column_df]

    df_raw[common.Headers.PF] = df_raw[common.Headers.PF].where(df_raw[common.Headers.sub_gsp] == True).ffill()

    return df_raw


def remove_unnecessary_rows(df_raw):
    """
		Function removes all of the rows which do not correspond to the usable data for GSP or Primary substations
	:param pd.DataFrame df_raw: Input DataFrame to be processed
	:return pd.DataFrame df_out:  Output DataFrame after processing
	"""

    # Get list of all rows which are determined as either a GSP or Primary substation
    idx = (
            (df_raw[common.Headers.sub_gsp] == True) |  # Confirm if the row is a GSP
            (df_raw[common.Headers.sub_primary] == True)  # Confirm if the row is a Primary
    )

    # Return subset of the original DataFrame including only these rows
    df_out = df_raw[idx]

    return df_out


def missing_year_load_estimator(df_raw, fill):
    """
		Function estimates all the missing years load values for both GSPs and primaries by linear interpolation and
		fills in adds the following columns: available years(years that have values, only estimates if at least 2 values
		are provided for years), year forecasted (if any year is forecasted for that substation), columns for estimated
		loads for each year
	:param pd.DataFrame df_raw: Input DataFrame to be processed
	:param bool fill:  Whether to fill missing values or not
	:return pd.DataFrame df_out:  Output DataFrame after processing
	"""

    forecast_years = common.adjust_years(headers_list=list(df_raw.columns))

    df_raw['available_years'] = np.nan
    df_raw['year_forecasted'] = np.nan

    year_estimate_list = ['{}_{}'.format(common.Headers.estimate, x) for x in forecast_years]
    # Add columns to DataFrame with no files by adding in empty
    df_raw = pd.concat([df_raw, pd.DataFrame(columns=year_estimate_list)], sort=False)

    forecast_years = common.adjust_years(headers_list=list(df_raw.columns))
    # todo: maybe add to idx to identify the rows which the values of the loads are negative
    idx = df_raw[forecast_years].isna()

    _ = (~idx.loc[1, :]).sum()

    df_raw['available_years'] = (~idx).sum(1)

    d = (df_raw['available_years'] > 1) & (df_raw['available_years'] < len(forecast_years))

    df_raw['year_forecasted'] = d

    est_row_list = d[d == True].index

    years_estimate_df = df_raw.loc[est_row_list, forecast_years]

    number_of_years = len(years_estimate_df.columns)
    years_estimate_df.columns = range(number_of_years)

    if fill:

        for j in range(len(est_row_list)):

            year_row1 = years_estimate_df.loc[est_row_list[j], :]
            t1 = year_row1.to_frame()
            estimated_array = common.interpolator(t1).T  # common.interpolator gets a dataframe with one column and
            # number of indexes equal to the number of years and the missing values as nan, then interpolate the missing
            # values by using indexes as _ and y being the values in the dataframe, then gives out a df with
            # the interpolated values
            n = 0
            for i in range(len(forecast_years)):
                if pd.isnull(df_raw.loc[est_row_list[j], forecast_years[i]]):
                    df_raw.loc[est_row_list[j], forecast_years[i]] = 1
                    df_raw.loc[est_row_list[j], forecast_years[i]] = estimated_array.iloc[0, n]
                    df_raw.loc[est_row_list[j], year_estimate_list[i]] = estimated_array.iloc[0, n]
                    n += 1

    return df_raw


def primary_diverse_load_adder(df_raw):
    """
		Function calculates the divers factor for all gsp subs and then fill it for primary sub the same value as their
		gsp subs and add it as a new column then it replaces the aggregate value of primary loads with the values
		written under the years for each primary then use the gsp divers factors to calculate the diversified loads of
		the primaries and write it in loads for primaries for each year
	:param pd.DataFrame df_raw: Input DataFrame to be processed
	:return pd.DataFrame df_out:  Output DataFrame after processing
	"""

    forecast_years = common.adjust_years(headers_list=list(df_raw.columns))
    adjusted_list = ['{}_{}'.format(common.Headers.aggregate, x) for x in forecast_years]
    df_raw[common.Headers.diverse_factor] = np.nan

    idx_change = (
            ~df_raw[common.Headers.sub_gsp].isna() &  # Confirm that its a gsp
            (~df_raw[forecast_years[0]].le(0)) & ~df_raw[forecast_years[0]].isna()
    )  # # Confirm that the the aggregated load value of the gsp is not zero or NA

    df_raw.loc[idx_change, common.Headers.diverse_factor] = df_raw.loc[idx_change, forecast_years[0]] / df_raw.loc[
        idx_change, adjusted_list[0]]

    idx_change = (
            ~df_raw[common.Headers.sub_gsp].isna() &  # Confirm that its a gsp
            (df_raw[forecast_years[0]].le(0)) | df_raw[
                forecast_years[
                    0]].isna())  # # Confirm that the the aggregated load value of the gsp is either zero or NA

    df_raw.loc[
        idx_change, common.Headers.diverse_factor] = 1  # if the aggregated value of the gsp (peak value) is zero or
    # NA then assume 1 as the diversity factor (this is to avoid division by zero)

    df_raw[common.Headers.diverse_factor] = df_raw[common.Headers.diverse_factor].where(
        df_raw[common.Headers.sub_gsp] == True).ffill()

    year_numbers = len(forecast_years)
    clipped_df = df_raw.clip(upper=pd.Series({common.Headers.diverse_factor: 1}), axis=1)
    df_raw[common.Headers.diverse_factor] = clipped_df[common.Headers.diverse_factor]

    for i in range(0, year_numbers):
        df_raw.loc[df_raw[common.Headers.sub_gsp].isna(), adjusted_list[i]] = df_raw.loc[
            df_raw[common.Headers.sub_gsp].isna(), forecast_years[i]]
        df_raw.loc[
            df_raw[common.Headers.sub_gsp].isna(), forecast_years[i]
        ] = df_raw.loc[
                df_raw[common.Headers.sub_gsp].isna(), forecast_years[i]
            ] * df_raw.loc[df_raw[common.Headers.sub_gsp].isna(), common.Headers.diverse_factor]

    return df_raw


def season_load_filler(df_raw, fill):
    """
		Function calculates the quantile values for season loads for both GSP and primary substations using available
		values (non zero and non NA) then saves them as attributes of class seasons (in common), where the percentile
		values are also setup. Then fill in the missing values for season loads using the calculated quantile values.
	:param pd.DataFrame df_raw: Input DataFrame to be processed
	:param bool fill:  Whether to fill data or not
	:return pd.DataFrame df_out:  Output DataFrame after processing
	"""
    if fill:
        idx = (
                ~df_raw[common.Headers.sub_gsp].isna() &  # Confirm that it's a gsp
                ~df_raw[common.Headers.spring_autumn].isna() &  # Confirm that a season value is not NA
                ~df_raw[common.Headers.spring_autumn].le(0))  # Confirm that season value is not negative

        common.Seasons.gsp_spring_autumn_val = np.percentile(
            df_raw.loc[idx, common.Headers.spring_autumn], common.Seasons.spring_autumn_q
        )

        idx_change = (
                ~df_raw[common.Headers.sub_gsp].isna() &
                (df_raw[common.Headers.spring_autumn].isna() | df_raw[common.Headers.spring_autumn].le(0))
        )

        df_raw.loc[idx_change, common.Headers.spring_autumn] = common.Seasons.gsp_spring_autumn_val

        idx = (
                ~df_raw[common.Headers.sub_gsp].isna() &  #
                ~df_raw[common.Headers.summer].isna() &  #
                ~df_raw[common.Headers.summer].le(0))  #

        common.Seasons.gsp_summer_val = np.percentile(df_raw.loc[idx, common.Headers.summer], common.Seasons.summer_q)

        idx_change = (
                ~df_raw[common.Headers.sub_gsp].isna() &
                (df_raw[common.Headers.summer].isna() | df_raw[common.Headers.summer].le(0))
        )

        df_raw.loc[idx_change, common.Headers.summer] = common.Seasons.gsp_summer_val

        idx = (
                ~df_raw[common.Headers.sub_gsp].isna() &  #
                ~df_raw[common.Headers.min_demand].isna() &  #
                ~df_raw[common.Headers.min_demand].le(0))  #

        common.Seasons.gsp_min_demand_val = np.percentile(
            df_raw.loc[idx, common.Headers.summer], common.Seasons.min_demand_q
        )

        idx_change = (
                ~df_raw[common.Headers.sub_gsp].isna() &
                (df_raw[common.Headers.min_demand].isna() | df_raw[common.Headers.min_demand].le(0))
        )

        df_raw.loc[idx_change, common.Headers.min_demand] = common.Seasons.gsp_min_demand_val

        # assessing primary season load values

        idx = (
                ~df_raw[common.Headers.sub_primary].isna() &  #
                ~df_raw[common.Headers.spring_autumn].isna() &  #
                ~df_raw[common.Headers.spring_autumn].le(0))  #

        common.Seasons.primary_spring_autumn_val = np.percentile(
            df_raw.loc[idx, common.Headers.spring_autumn], common.Seasons.spring_autumn_q
        )

        idx_change = (
                ~df_raw[common.Headers.sub_primary].isna() &
                (df_raw[common.Headers.spring_autumn].isna() | df_raw[common.Headers.spring_autumn].le(0))
        )

        df_raw.loc[idx_change, common.Headers.spring_autumn] = common.Seasons.primary_spring_autumn_val

        idx = (
                ~df_raw[common.Headers.sub_primary].isna() &  #
                ~df_raw[common.Headers.summer].isna() &  #
                ~df_raw[common.Headers.summer].le(0))  #

        common.Seasons.primary_summer_val = np.percentile(
            df_raw.loc[idx, common.Headers.summer], common.Seasons.summer_q
        )

        idx_change = (
                ~df_raw[common.Headers.sub_primary].isna() &
                (df_raw[common.Headers.summer].isna() | df_raw[common.Headers.summer].le(0))
        )

        df_raw.loc[idx_change, common.Headers.summer] = common.Seasons.primary_summer_val

        idx = (
                ~df_raw[common.Headers.sub_primary].isna() &  #
                ~df_raw[common.Headers.min_demand].isna() &  #
                ~df_raw[common.Headers.min_demand].le(0))  #

        common.Seasons.primary_min_demand_val = np.percentile(
            df_raw.loc[idx, common.Headers.min_demand], common.Seasons.min_demand_q
        )

        idx_change = (
                ~df_raw[common.Headers.sub_primary].isna() &
                (df_raw[common.Headers.min_demand].isna() | df_raw[common.Headers.min_demand].le(0))
        )

        df_raw.loc[idx_change, common.Headers.min_demand] = common.Seasons.primary_min_demand_val

    # TODO: @NS (addressed) - Add in a new column that introduces Winter Peak = 1.0 so that if the user selects Winter Peak in
    # TODO: @NS (addressed)... in the GUI it easily allows the user to select the appropriate values
    # TODO: @NS (addressed)... I've had a guess at what will be needed but please check this doesn't cause issues anywhere else -
    #
    # Winter peak demand values will all be 100% of the value for that year.

    df_raw.loc[:, common.Headers.winter_peak] = 1.0

    return df_raw


def bad_data_identifier(df_raw):
    """
		Function removes all of the rows which do not correspond to the usable data for GSP or Primary substations
	:param pd.DataFrame df_raw: Input DataFrame to be processed
	:return pd.DataFrame df_out:  Output DataFrame after processing
	"""

    forecast_years = common.adjust_years(headers_list=list(df_raw.columns))
    bus_list = filter(lambda x: x.startswith('PS'), df_raw.columns)

    idx = (
            df_raw[forecast_years].isna().all(axis='columns') | df_raw[forecast_years].le(0).all(axis='columns') |
            df_raw[bus_list].isna().all(axis='columns') | df_raw[bus_list].le(0).all(axis='columns')
    )

    bad_data = df_raw.loc[idx, :]
    good_data = df_raw.loc[~idx, :]
    bad_data.to_excel(common.ExcelFileNames.bad_data_excel_name)
    good_data.to_excel(common.ExcelFileNames.good_data_excel_name)
    return bad_data, good_data


def main(xl_path, dill):
    """
		Function
	:param str xl_path:  Path to raw results file to be imported
	:param dill:  if true it would dill the dataframes generated in this code
	:return pd.DataFrame df_return:  DataFrame to return
	"""
    local_fill_estimate_list = [False, True]
    excel_output_name_list = [common.ExcelFileNames.df_raw_excel_name, common.ExcelFileNames.df_modified_excel_name]

    # Create an empty list to store the results in, the final set is then returned
    dfs = list()

    # Function loops through twice to produce 2 DataFrames
    for i in range(len(local_fill_estimate_list)):
        df = common.import_raw_load_estimates(pth_load_est=xl_path)
        # Identify whether a GSP or Primary substation for each row
        # raw_dataframe = common.sse_load_xl_to_df(xl_filename=FILE_PTH_INPUT,
        # xl_ws_name='MASTER Based on SubstationLoad', headers=True)
        df = determine_gsp_primary_flag(df_raw=df)
        # Extract aggregate demand for each GSP
        df = extract_aggregate_demand(df_raw=df)
        # Assign GSPs
        df = assign_gsp(df_raw=df)
        # Extract bus percentages as new columns
        df = bus_percentage_adder_modified(df_raw=df, fill=local_fill_estimate_list[i])
        df = assign_pf(df_raw=df, fill=local_fill_estimate_list[i])  # Neg added

        df = remove_unnecessary_rows(df_raw=df)

        #  Estimates the missing load values for each year by inter/extrapolation.
        df = missing_year_load_estimator(df_raw=df, fill=local_fill_estimate_list[i])  # Neg added

        # Calculate the diversity factors as new column then fill in the aggregate and actual(divers) loads and assumes
        # divers factor of 1 for GSPs with 0 or NA peak loads
        df = primary_diverse_load_adder(df_raw=df)  # Neg added

        # Fill in the missing season load values by the quantiles
        df = season_load_filler(df_raw=df, fill=local_fill_estimate_list[i])  # Neg added

        # Export processed DataFrame
        file_pth_output = common.get_local_file_path(file_name=excel_output_name_list[i])
        df.to_excel(file_pth_output)

        dfs.append(df)

    # make an excel file of bad data
    bad_data, good_data = bad_data_identifier(df)

    comparison.excel_data_comparison_maker(
        FILE_NAME_INPUT_1=common.ExcelFileNames.df_raw_excel_name,
        FILE_NAME_INPUT_2=common.ExcelFileNames.df_modified_excel_name,
        Bad_Data_Input_Name=common.ExcelFileNames.bad_data_excel_name,
        Good_Data_Input_Name=common.ExcelFileNames.good_data_excel_name
    )

    if dill:  # if true it would dill the dataframes generated in this code

        sheet_list = ['Sheet1', 'Sheet1', 'Sheet1', 'Sheet1']
        input_list = [common.ExcelFileNames.df_raw_excel_name, common.ExcelFileNames.df_modified_excel_name,
                      common.ExcelFileNames.good_data_excel_name, common.ExcelFileNames.bad_data_excel_name,
                      ]
        variable_list = [common.folder_file_names.dill_raw_data, common.folder_file_names.dill_modified_data,
                         common.folder_file_names.dill_good_data_name, common.folder_file_names.dill_bad_data_name]
        variables_to_be_dilled = [True, True, True, True]
        dill_folder_name = common.folder_file_names.dill_folder
        load_dill = True  # when True it would try to load the dills for the variable to be dilled that are False
        variable_dict = common.batch_dill_maker_loader(input_list=input_list, sheet_list=sheet_list,
                                                       variable_list=variable_list,
                                                       variables_to_be_dilled=variables_to_be_dilled,
                                                       load_dill=load_dill,
                                                       dill_folder_name=dill_folder_name)

    # DataFrame to return is the last one that was processed
    df_returned = dfs[-1]
    return df_returned


if __name__ == '__main__':
    FILE_NAME_INPUT = common.ExcelFileNames.FILE_NAME_INPUT
    # FILE_NAME_OUTPUT = 'Processed Load Estimates_p_non_modified.xlsx'
    FILE_PTH_INPUT = common.get_local_file_path(file_name='2019-20 SHEPD Load Estimates - v6-check.xlsx')
    dill = True  # if true it would dill the dataframes generated in this code

    # Import raw DF
    fill_estimate_list = [False, True]
    df_modified = main(xl_path=FILE_PTH_INPUT, dill=dill)

    k = 1
