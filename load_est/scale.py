"""
#######################################################################################################################
###											scale											###
###																													###
###		Code developed by Negar Shams (negar.shams@PSCconsulting.com, +44 7436 544893) as part of PSC 		 		###
###		project JK7938 - SHEPD - studies and automation																###
###																													###
#######################################################################################################################
"""

# Generic Imports
import pandas as pd
import numpy as np
# Unique imports
import common_functions as common
import constants as constants
import collections
import load_est
import load_est.psse as psse
import math
import time
import os
import dill


# Functions
#
def scale_loads(df_load_values, year=str(), season=str(), diverse=False, zone=tuple(), gsp=tuple()):
    # type: (df, str, str, bool, tuple, tuple) -> object
    """
        Function to update station loads
    :param pd.DataFrame df_load_values:  DataFrame with all of the load values needed for scaling
    :param str year:  Year loads should be scaled for
# 	:param str() season:  Season years should be scaled for
# 	:param bool diverse:  Whether to scale for the diversified or aggregate load values
# 	:param tuple zone:  Populated with a list of zones if the user selects specific zones to consider
# 	:param tuple gsp:  Populated with a list of gsps if the user selects specific zones to consider
# 	:return:
# 	"""
    #
    df = df_load_values
    df = df.loc[df['Sub_Primary'] == 1, :]  # filters the primary buses
    zone_nu = len(list(zone))
    gsp_nu = len(list(gsp))
    k = 1


    diverse_load_year_list = common.adjust_years(headers_list=list(df.columns))
    aggregate_load_year_list = filter(lambda local_x: local_x.startswith('agg'), df.columns)
    year_aggregate_name = filter(lambda local_x: year[0] in local_x, aggregate_load_year_list)
    df_season = df.loc[:, season]

    index = range(0, 500)
    columns = ['Bus Number', 'MVA', 'p.f', 'GSP', 'Primary', 'P', 'Q']
    df_loads = pd.DataFrame(index=index, columns=columns)  # an empty df is made so its values can be changed later on

    if diverse == False:
        df_year = df.loc[:, year_aggregate_name]
        year_column = year_aggregate_name
    else:
        df_year = df.loc[:, year]
        year_column = year

    df_year.columns = ['temp']  # this is done to ease multiplying two dfs
    df_season.columns = ['temp']
    df_year_season = df_year['temp'] * df_season['temp']
    df.loc[:, 'year_season'] = df_year_season
    bus_list = filter(lambda local_x: local_x.startswith('PS'), df.columns)
    percent_list = filter(lambda local_x: local_x.startswith('per'), df.columns)

    df = df.reset_index(drop=True)  # this is done to ease looping with index starting from 0 and increasing normally
    idx = (~df[bus_list].isna())  # makes a boolean dataframe of the size df rows and number of columns equal to bus
    # numbers and gives True if value is not NA and False if it is, then this df is used in
    # the following for loop

    m = 0
    # the following makes a dataframe of the load buses with their associated values
    for i in range(0, len(df)):
        for j in range(0, len(bus_list)):
            if idx.loc[i, bus_list[j]]:
                df_loads.loc[m, 'Bus Number'] = df.loc[i, bus_list[j]]
                df_loads.loc[m, 'MVA'] = df.loc[i, 'year_season'] * df.loc[i, percent_list[j]]
                df_loads.loc[m, 'p.f'] = df.loc[i, common.Headers.PF]
                df_loads.loc[m, 'GSP'] = df.loc[i, common.Headers.gsp]
                df_loads.loc[m, 'Primary'] = df.loc[i, common.Headers.name]
                # df_loads.loc[m, 'zone'] = df.loc[i, 'zone']
                df_loads.loc[m, 'P'] = df_loads.loc[m, 'MVA'] * df_loads.loc[m, 'p.f']
                # df_loads.loc[m, 'Q']=(1-(df_loads.loc[m,'p.f'])**2)**(1/2)
                df_loads.loc[m, 'Q'] = math.sqrt((1 - ((df_loads.loc[m, 'p.f']) ** 2))) * df_loads.loc[m, 'MVA']
                m += 1

    idx_na = (~df_loads['Bus Number'].isna())
    df_loads = df_loads.loc[idx_na, :]  # this line drops the NA bus number values (as earlier a big NA pd was formed)
    # df_loads['Zone'] = np.nan

    loads = psse.LoadData()  # gets the loads df from the psse

    idx = df_loads['Bus Number'].isin(loads.df[constants.Loads.bus])
    # loads_in_psse = df_loads.loc[idx == True, :]
    loads_not_in_psse = df_loads.loc[idx == False, :]  # load buses that are not in psse (but are

    if gsp_nu > 0:
        idx = (df_loads['GSP'].isin(list(gsp)))  # filters the buses that have GSPs inside the given GSP list
        df_loads = df_loads.loc[idx == True, :]

    if zone_nu > 0:
        idx = loads.df[constants.Loads.zone].isin(list(zone))
        psse_loads = loads.df.loc[idx == True, :]  # filters the buses that have ZONEs inside the given ZONE list
    else:
        psse_loads = loads.df

    idx = df_loads['Bus Number'].isin(psse_loads[constants.Loads.bus])
    loads_in_psse = df_loads.loc[idx == True, :]  # loads in psse that are also filtered by gsp and zones
    k = 1

    if len(loads_not_in_psse) > 0:
        msg0 = (
            'The following load buses were not found in PSSE:'
        )
        msg1 = '\n'.join((map(str, loads_not_in_psse['Bus Number'])))
        logger.warning('{}\n{}'.format(msg0, msg1))

        # logger.debug('Warning: the following load buses were not found in psse:')
        # logger.debug('\n'.join(map(str, loads_not_in_psse['Bus Number'])))

    loads_in_psse = loads_in_psse.reset_index(drop=True)

    idx = psse_loads[constants.Loads.bus].isin(loads_in_psse['Bus Number'])
    loads_in_both = psse_loads.loc[idx == True, :]
    loads_in_both = loads_in_both.reset_index(drop=True)
    loads_only_in_psse = psse_loads.loc[idx == False, :]

    # changes the id type to int so that we can filter it
    loads_in_both[constants.Loads.identifier] = loads_in_both[constants.Loads.identifier].astype(int)
    load_in_both_id_not_1 = loads_in_both[loads_in_both[constants.Loads.identifier] <> 1]
    # changes the id type back to string as the load change gets string
    load_in_both_id_not_1[constants.Loads.identifier] = load_in_both_id_not_1[constants.Loads.identifier].astype(str)

    # changes the Bus Number type to int as load change function of psse does not accept float or string as bus number
    loads_in_psse['Bus Number'] = loads_in_psse['Bus Number'].astype(int)

    load_in_both_id_not_1 = load_in_both_id_not_1.reset_index(drop=True)

    # loads.loads_to_change = loads_in_psse  #
    loads.disable_rest_loads(loads_id_not_1=load_in_both_id_not_1)

    loads.change_load(loads_to_change=loads_in_psse)

    k = 1

    return None


def scale_gens(pc=float(),zone=tuple()):
    """
    Function to Scale all generation in a given zone tuple by scaling percentage
    :param pc: float percentage value for scaling the generators
    :param tuple zone:  Populated with a list of zones if the user selects specific zones to consider
    :return: None
    """

    machine_data = psse.MachineData()
    bus_data = psse.BusData()
    machine_data.df['ZONE'] = np.nan

    df_map = bus_data.df
    df = machine_data.df
    columns_df_map = ['NUMBER', 'ZONE']
    columns_df = ['NUMBER', 'ZONE']
    df_mapped = mapper(df_map=df_map, df=df, columns_df_map=columns_df_map, columns_df=columns_df)

    zone_nu = len(list(zone))

    if zone_nu > 0:
        idx = df_mapped[constants.Loads.zone].isin(list(zone))
        psse_gens = df_mapped.loc[idx == True, :]  # filters the gens that have ZONEs inside the given ZONE list
    else:
        psse_gens = df_mapped

    for gen in psse_gens.itertuples():

        if gen.RPOS != 999:
            machine_data.machine_change(gen_pd=gen, pc=pc)

    return None


def mapper(df_map, df, columns_df_map, columns_df):
    """
        Function returns dataframe with mapped values.
    :param dataframe df_map:  dataframe that is used as the map
    :param dataframe df:  dataframe that needs to be mapped
    :param list columns_df_map: a list of two column headers from df_map. First header being the attribute which is used
    to map and the second the attribute that is the result of the mapping
    :param list columns_df: a list of two column headers from df. First header being the attribute which is used
    to map and the second the attribute that is the result of the mapping
    :return  dataframe df: The mapped version of the df
    """

    df_map = df_map.set_index(columns_df_map[0])
    key_list= df_map.index
    is_duplicate_1 = key_list.duplicated(keep="first")
    not_duplicate = ~is_duplicate_1
    df_map_not_dublicated = df_map[not_duplicate]
    df_to_be_mapped=df[columns_df[0]]
    df_mapper= df_map_not_dublicated[columns_df_map[1]]
    df_mapper_dict = df_mapper.to_dict()
    df_mapped = df_to_be_mapped.map(df_mapper_dict)
    df_mapped = df_mapped.to_frame()
    df_mapped.columns = [columns_df[1]]
    df[columns_df[1]]=df_mapped
    k=1

    return df


if __name__ == '__main__':
    uid = 'BKDY_{}'.format(time.strftime('%Y%m%d_%H%M%S'))
    script_path = os.path.realpath(__file__)
    script_folder = os.path.dirname(script_path)
    temp_folder = os.path.join(script_folder, 'temp')
    if not os.path.exists(temp_folder):
        os.mkdir(temp_folder)

    psse_case = r'C:\Users\NegarShams\PycharmProjects\JK7938_SHEPD_LoadEstimates - Copy\load_est\test_files\SHEPD 2018 LTDS Winter Peak v33_test1.sav'
    psse_con = psse.PsseControl()
    psse_con.load_data_case(pth_sav=psse_case)
    logger = load_est.Logger(pth_logs=temp_folder, uid=uid, debug=True)

    FILE_NAME_INPUT = 'processed_load_estimate_modified.xlsx'
    FILE_NAME_INPUT = 'good_data.xlsx'
    FILE_PTH_INPUT = common.get_local_file_path(file_name=FILE_NAME_INPUT)
    df = common.import_excel(FILE_PTH_INPUT)
    i = [1]
    variable_name = ['good_data']
    dill_or_not = [False]
    load_dill = True
    machine_data = psse.MachineData()
    bus_data = psse.BusData()
    machine_data.df['ZONE'] = np.nan

    df_map = bus_data.df
    df = machine_data.df
    columns_df_map = ['NUMBER', 'ZONE']
    columns_df = ['NUMBER', 'ZONE']
    df_mapped = mapper(df_map=df_map, df=df, columns_df_map=columns_df_map, columns_df=columns_df)
    zone = (5, 102)
    pc = 50
    scale_gens(pc=pc,zone=zone)

    k = 1

    dill_folder_name = common.folder_file_names.dill_folder
    variable_dict = common.batch_dill_maker_loader_no_xl(variable_value_list=i, variable_list=variable_name,
                                                         variables_to_be_dilled=dill_or_not, load_dill=load_dill,
                                                         dill_folder_name=dill_folder_name)
    df_modified_dill = common.load_dill(
        variable_name=[common.folder_file_names.dill_modified_data],
        dill_folder_name=common.folder_file_names.dill_folder)

    df_from_dill = common.load_dill(variable_name=variable_name, dill_folder_name=dill_folder_name)

    dict_test = df_from_dill[['GSP']].to_dict()['GSP']

    year = ['2020 / 2021']
    diverse = False
    season = ['Summer']
    # zone=(5,102)# this should be a tuple of integer values
    zone = ()
    gsp = ('ABERNETHY', 'ALNESS', 'PETERHEAD SHELL', 'LUNANHEAD', 'THURSO')
    # gsp = ()
    pc = 50

    # scale_loads(df_load_values=df, year=year, season=season, diverse=diverse, zone=zone, gsp=gsp)

    scale_gens(pc=pc)

    k = 1
