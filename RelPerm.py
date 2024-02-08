import csv
import re
import warnings
import os
import math

import os.path as osp
import numpy as np
import pandas as pd

#INFILE = r""
FOLDER = r""
TEST = "SCEN"
TEMPLATE = 'OY'


class RelPerm:

    def __init__(self):

        self.mnemonics: dict = self.load_mnemonics()

        with open('lab_names.csv', 'r') as f:
            next(f)
            self.lab_names: dict = {key: val.upper() for key, val in csv.reader(f)}

        self.accepted_units = {'TEXT': [],
                               'UNITLESS': [],
                               'V/V': ['FRACTION', 'DSPC/DPC'],
                               'G': ['GRAM'],
                               'G/CC': ['G/CM3', 'GM/CC', 'GRAM/CM3'],
                               'DYNE/CM': ['MN/M'],
                               'CC': ['CM3'],
                               'MD': ['MILLIDARCIES'],
                               'MD/MD': ['=DQ/DT'],
                               'CP': ['MPA.S'],
                               'FT': ['FEET']}

        self.log_templates = self.load_templates()
        self.samples = []
        self.log = None

    @staticmethod
    def load_mnemonics():
        mnemonics: dict = {'OY': {}, 'WR': {}}

        with open(r'Mnemonics\mnem_OY_MCEN.csv', 'r') as f:
            next(f)
            mnemonics['OY']['MCEN'] = {key.upper(): val.upper() for key, val in csv.reader(f)
                                                  if not val == ''}

        with open(r'Mnemonics\mnem_OY_SCEN.csv', 'r') as f:
            next(f)
            mnemonics['OY']['SCEN'] = {key.upper(): val.upper() for key, val in csv.reader(f)
                                                   if not val == ''}

        with open(r'Mnemonics\mnem_OY_SS.csv', 'r') as f:
            next(f)
            mnemonics['OY']['SS'] = {key.upper(): val.upper() for key, val in csv.reader(f)
                                        if not val == ''}

        with open(r'Mnemonics\mnem_WR_MCEN.csv', 'r') as f:
            next(f)
            mnemonics['WR']['MCEN'] = {key.upper(): val.upper() for key, val in csv.reader(f)
                                                  if not val == ''}

        with open(r'Mnemonics\mnem_WR_SCEN.csv', 'r') as f:
            next(f)
            mnemonics['WR']['SCEN'] = {key.upper(): val.upper() for key, val in csv.reader(f)
                                                  if not val == ''}

        with open(r'Mnemonics\mnem_WR_SS.csv', 'r') as f:
            next(f)
            mnemonics['WR']['SS'] = {key.upper(): val.upper() for key, val in csv.reader(f)
                                                  if not val == ''}

        return mnemonics

    @staticmethod
    def load_templates():
        templates: dict = {'MCEN': pd.read_csv(r'Templates\tpl_MCEN.csv'),
                           'SCEN': pd.read_csv(r'Templates\tpl_SCEN.csv'),
                           'SS': pd.read_csv(r'Templates\tpl_SS.csv')}

        return templates

    def prepare_sample(self, infile, test, template):

        xl = pd.ExcelFile(infile)

        if template == 'OY':

            if test == 'MCEN':
                # Parse all tables from input excel
                user_data = xl.parse('PC_Multi_Speed_Centrifuge', header=None, index_col=0, usecols='A:C', skiprows=1) \
                    .dropna(how='all') \
                    .transpose() \
                    .sort_index(ascending=False) \
                    .reset_index(drop=True)

                additional_data = xl.parse('AdditionalData', header=None, index_col=0, usecols='A:C') \
                    .dropna(how='all') \
                    .transpose() \
                    .reset_index(drop=True)

                analytical_analysis1 = xl.parse('PC_Multi_Speed_Centrifuge', header=2, usecols='Q,R,S') \
                    .dropna(how='all') \
                    .reset_index(drop=True)

                analytical_analysis2 = xl.parse('PC_Multi_Speed_Centrifuge', header=3, usecols='K,L') \
                    .dropna(how='all') \
                    .reset_index(drop=True)

                numerical_interp = pd.DataFrame({'Sw': 'V/V', 'Pc (bar)': 'BAR'}, index=[0])
                numerical_interp = numerical_interp.append(
                    xl.parse('PC_Multi_Speed_Centrifuge', header=2, usecols='AD,AE'),
                    ignore_index=True) \
                    .dropna(how='all') \
                    .reset_index(drop=True)

                raw_test = xl.parse('PC_Multi_Speed_Centrifuge', header=1, usecols='U:X') \
                    .dropna(how='all') \
                    .reset_index(drop=True)

                # Merge tables
                df = pd.concat(
                    [user_data, additional_data, analytical_analysis1, analytical_analysis2, numerical_interp, raw_test],
                    axis=1)

            elif test == 'SCEN':
                # Parse all tables from input excel
                user_data = xl.parse('SingleSpeedCentrifuge',
                                      header=None, index_col=0, usecols='A:C', skiprows=1) \
                    .dropna(how='all') \
                    .transpose() \
                    .drop(columns='Rotational speed') \
                    .sort_index(ascending=False) \
                    .reset_index(drop=True)

                additional_data = xl.parse('AdditionalData', header=None, index_col=0, usecols='A:C') \
                    .dropna(how='all') \
                    .transpose() \
                    .reset_index(drop=True)

                analytical_analysis = xl.parse('SingleSpeedCentrifuge', header=1, usecols='I, L') \
                    .dropna(how='all') \
                    .reset_index(drop=True)

                numerical_interp1 = xl.parse('SingleSpeedCentrifuge', header=1, usecols='AA,AC') \
                    .dropna(how='all') \
                    .reset_index(drop=True)

                numerical_interp2 = pd.concat(
                    [pd.DataFrame({
                        'Swi': 'v/v',
                        'Sor': 'v/v',
                        'krw(Sor)': 'mD/mD',
                        'krow(Swi)': 'mD/mD',
                        'nw': 'unitless',
                        'now': 'unitless',
                        'Sorg': 'v/v',
                        'Sgr': 'v/v',
                        'krog(Sgr)': 'mD/mD',
                        'krg(Sorg)': 'mD/mD',
                        'nog': 'unitless',
                        'ng': 'unitless',
                        'cw': 'unitless',
                        'co': 'unitless'},
                        index=[0]),
                    xl.parse('SingleSpeedCentrifuge', header=None, index_col=0, usecols='AD:AE')
                        .dropna(how='all')
                        .transpose()
                        .reset_index(drop=True)],
                    join='inner',
                    ignore_index=True)

                raw_test = xl.parse('SingleSpeedCentrifuge', header=1, usecols='N:Q')\
                            .dropna(how='all') \
                            .reset_index(drop=True)

                # Merge tables
                df = pd.concat(
                    [user_data, additional_data, analytical_analysis,
                     numerical_interp1, numerical_interp2, raw_test],
                    axis=1)

            elif test == 'SS':
                user_data = xl.parse(1, header=None, index_col=0, usecols='A:C', skiprows=1) \
                    .dropna(how='all') \
                    .transpose() \
                    .sort_index(ascending=False) \
                    .reset_index(drop=True)

                additional_data = xl.parse('AdditionalData', header=None, index_col=0, usecols='A:C') \
                    .dropna(how='all') \
                    .transpose() \
                    .reset_index(drop=True)

                analytical_analysis = xl.parse(1, header=1, usecols='F:G, I:M') \
                    .dropna(how='all') \
                    .reset_index(drop=True) \
                    .rename(columns={'Water Rate': 'Water Rate_ANA',
                                     'Oil Rate': 'Oil Rate_ANA',
                                     'krw': 'krw_ANA',
                                     'kro': 'kro_ANA'})

                numerical_interp1 = pd.concat(
                    [pd.DataFrame({
                        'Sw': 'v/v',
                        'krw': 'mD/mD',
                        'kro': 'mD/mD'},
                        index=[0]),
                        xl.parse(1, header=2, usecols='X,Y,Z')
                        .dropna(how='all')
                        .reset_index(drop=True)],
                    join='inner',
                    ignore_index=True) \
                    .rename(columns={'krw': 'krw_NUM',
                                     'kro': 'kro_NUM'})

                numerical_interp2 = xl.parse(1, header=1, usecols='AC,AE') \
                    .dropna(how='all') \
                    .reset_index(drop=True) \
                    .rename(columns={'Sw': 'Sw_PC'})

                raw_test = xl.parse(1, header=1, usecols='O:T') \
                    .dropna(how='all') \
                    .rename(columns=lambda x: x.replace('.1', '')) \
                    .reset_index(drop=True)

                sat = pd.DataFrame({
                    'DIST_IN_CORE': 'CM',
                    'XN_BASE_WAT': 'UNITLESS',
                    'XN_BASE_OIL': 'UNITLESS',
                    'XN_FW_0': 'UNITLESS',
                    'XN_FW_1': 'UNITLESS',
                    'XN_FW_2': 'UNITLESS',
                    'XN_FW_3': 'UNITLESS',
                    'XN_FW_4': 'UNITLESS',
                    'XN_FW_5': 'UNITLESS',
                    'XN_FW_6': 'UNITLESS',
                    'XN_FW_7': 'UNITLESS',
                    'XN_FW_8': 'UNITLESS',
                    'XN_FW_9': 'UNITLESS',
                    'XN_FW_10': 'UNITLESS',
                    'XN_FW_11': 'UNITLESS',
                    'XN_FW_12': 'UNITLESS',
                    'XN_FW_BUMP_1': 'UNITLESS',
                    'XN_FW_BUMP_2': 'UNITLESS',
                    'SATX_SS_FW_0': 'V/V',
                    'SATX_SS_FW_1': 'V/V',
                    'SATX_SS_FW_2': 'V/V',
                    'SATX_SS_FW_3': 'V/V',
                    'SATX_SS_FW_4': 'V/V',
                    'SATX_SS_FW_5': 'V/V',
                    'SATX_SS_FW_6': 'V/V',
                    'SATX_SS_FW_7': 'V/V',
                    'SATX_SS_FW_8': 'V/V',
                    'SATX_SS_FW_9': 'V/V',
                    'SATX_SS_FW_10': 'V/V',
                    'SATX_SS_FW_11': 'V/V',
                    'SATX_SS_FW_12': 'V/V',
                    'SATX_SS_FW_BUMP_1': 'V/V',
                    'SATX_SS_FW_BUMP_2': 'V/V'},
                    index=[0])

                sat_prof = pd.concat(
                    [sat,
                    xl.parse('SaturationProfiles', header=None, names=sat.columns ,index_col=None, usecols='A:C, F:T, W:AK',
                             skiprows=5)
                        .dropna(how='all')
                        .reset_index(drop=True)],
                    join='outer',
                    ignore_index=True)

                # Merge tables
                df = pd.concat(
                    [user_data, additional_data, analytical_analysis,
                     numerical_interp1, numerical_interp2, raw_test, sat_prof],
                    axis=1)

        elif template == 'WR':

            if test == 'MCEN_DEACTIVATED': #_ORIGINAL WR TEMPLATE
                user_data = pd.concat([
                    xl.parse(0, header=None, index_col=0, usecols='A:C', skiprows=2, nrows=2),
                    xl.parse(0, header=None, index_col=0, usecols='A:C', skiprows=35, nrows=14),
                    xl.parse(0, header=None, names=[0, 1, 2], index_col=0, usecols='D:F', skiprows=35, nrows=8),
                    xl.parse(0, header=None, names=[0, 1, 2], index_col=0, usecols='H:J', skiprows=35, nrows=1)
                ]) \
                    .dropna(how='all') \
                    .transpose() \
                    .sort_index(ascending=False) \
                    .reset_index(drop=True)

                additional_data = pd.concat([
                    xl.parse(0, header=None, index_col=0, usecols='A:B', skiprows=50, nrows=20),
                    xl.parse(0, header=None, names=[0, 1], index_col=0, usecols='A,C', skiprows=72, nrows=6)])
                additional_data.dropna(how='all', inplace=True)
                additional_data[2] = "UNITLESS"
                additional_data = additional_data.transpose().sort_index(ascending=False).reset_index(drop=True)
                additional_data.loc[0, ['Depth (ft)', 'Texp (F)', 'Stress (psi)']] = ['FT', 'F', 'PSI']
                additional_data.rename(columns={'Texp (F)': 'Experimental Temperature',
                                                'Stress (psi)': 'Applied stress'},
                                       inplace=True)

                analytical_analysis = xl.parse(0, header=5, usecols='A,B,H,I', nrows=28) \
                    .dropna(how='any') \
                    .reset_index(drop=True)
                """
                if sheet == 'ImPc_Analytical':
                    numerical_interp = pd.concat([
                            pd.DataFrame({"Sw": "V/V", "Pc": "BAR"}, index=[0]),
                            xl.parse(
                                sheet, header=None, names=["Sw", "Pc"], usecols='W,X', skiprows=8)\
                                .dropna(how='all')],
                        join='inner',
                        ignore_index=True)
                elif sheet == '2ndDrPc_Analytical':
                    numerical_interp = xl.parse(
                            sheet, header=7, usecols='W,X')\
                            .dropna(how='all')\
                            .reset_index(drop=True)"""
                numerical_interp = pd.DataFrame()

                raw_test = xl.parse(0, header=6, usecols='AA:AC') \
                    .dropna(how='all') \
                    .reset_index(drop=True)

                # Merge tables
                df = pd.concat(
                    [user_data, additional_data, analytical_analysis, numerical_interp, raw_test],
                    axis=1)

            elif test == 'MCEN_DEACTIVATED':
                # For combination of doc style and WR templates
                # 1.-3. Tables at row 5; 4.-6. Tables at row 35;
                # Additional Data Row 51; Additional Tables Column W and Z

                user_data = pd.concat([
                    xl.parse('Base Properties', header=None, names=[0, 1, 2], index_col=0, usecols='B:D', skiprows=2, nrows=3),
                    xl.parse('Base Properties', header=None, index_col=0, usecols='A:C', skiprows=8, nrows=3),
                    xl.parse('Base Properties', header=None, index_col=0, usecols='A:C', skiprows=16, nrows=1),
                    xl.parse('Base Properties', header=None, index_col=0, usecols='A:B', skiprows=44, nrows=1),
                    xl.parse('ImPc_Cfg', header=None, index_col=0, usecols='A:C', skiprows=35, nrows=8),
                    xl.parse('ImPc_Cfg', header=None, names=[0, 1, 2], index_col=0, usecols='D:F', skiprows=35, nrows=8),
                    xl.parse('ImPc_Cfg', header=None, names=[0, 1, 2], index_col=0, usecols='H:J', skiprows=35, nrows=1)])\
                    .dropna(how='all') \
                    .transpose() \
                    .sort_index(ascending=False) \
                    .reset_index(drop=True)

                user_data = pd.concat([user_data,
                                       xl.parse('Base Properties', names=['Experimental Temperature'],
                                                usecols='C', skiprows=31, nrows=2)], axis=1)\
                    .rename(columns={'Net Confining Stress (Centrifuge)': 'Applied stress'})
                user_data.loc[0, 'Experimental Temperature'] = 'F'

                additional_data = pd.concat([
                    xl.parse('ImPc_Cfg', header=None, index_col=0, usecols='A:C', skiprows=50, nrows=5),
                    xl.parse('ImPc_Cfg', header=None, names=[0, 1 ,2], index_col=0, usecols='A,C,D', skiprows=55, nrows=7)])\
                    .dropna(how='all')\
                    .transpose()\
                    .sort_index(ascending=False)\
                    .reset_index(drop=True)

                analytical_analysis = xl.parse('ImPc_Cfg', header=5, usecols='A,B', nrows=27) \
                    .dropna(how='any') \
                    .reset_index(drop=True)

                analytical_analysis2 = pd.concat([
                            pd.DataFrame({"S inflow": "V/V", "<Pc_i>": "BAR"}, index=[0]),
                            xl.parse('ImPc_Cfg', header=None, names=["S inflow", "<Pc_i>"],
                                     usecols='W,Z', skiprows=8, nrows=10)])\
                    .dropna(how='all')\
                    .reset_index(drop=True)

                numerical_interp = pd.concat([
                            pd.DataFrame({"Sw": "V/V", "Pc": "BAR"}, index=[0]),
                            xl.parse('ImPc_Cfg', header=None, names=["Sw", "Pc"],
                                     usecols='W,Z', skiprows=20, nrows=10)])\
                    .dropna(how='all')\
                    .reset_index(drop=True)

                raw_test = xl.parse('Centrifuge Imbibition_PC', header=3, usecols='L:N') \
                    .dropna(how='all') \
                    .reset_index(drop=True)

                # Merge tables
                df = pd.concat(
                    [user_data, additional_data, analytical_analysis, analytical_analysis2, numerical_interp, raw_test],
                    axis=1)

            elif test == 'MCEN':
                """Tables 1-3 at row 5, Tables 4-6 at row 22, additional data at row 35
                    Separate tab for Raw test data, multiple sections per Speed
                    Example:"""

                tabs = [tab for tab in xl.sheet_names if tab.upper() in ['PC_PRIMDRA', 'PC_IMB']]

                tab_dfs = []

                for tabname in tabs:

                    user_data = pd.concat([
                        xl.parse(tabname, header=None, index_col=0, usecols='A:C', skiprows=2, nrows=1),
                        xl.parse(tabname, header=None, index_col=0, usecols='A:C', skiprows=22, nrows=10),
                        xl.parse(tabname, header=None, names=[0, 1, 2], index_col=0, usecols='D:F', skiprows=22, nrows=8),
                        xl.parse(tabname, header=None, names=[0, 1, 2], index_col=0, usecols='H:J', skiprows=22, nrows=2)
                    ]) \
                        .dropna(how='all') \
                        .transpose() \
                        .sort_index(ascending=False) \
                        .reset_index(drop=True)

                    if tabname == 'Pc_PrimDra':
                        additional_data = pd.concat([
                            xl.parse(tabname, header=None, index_col=0, usecols='A:C', skiprows=34, nrows=15),
                            xl.parse(tabname, header=None, names=[0, 1, 2], index_col=0, usecols='A,C,D', skiprows=50, nrows=6)])
                    elif tabname == 'Pc_Imb':
                        additional_data = pd.concat([
                            xl.parse(tabname, header=None, index_col=0, usecols='A:C', skiprows=32, nrows=4),
                            xl.parse(tabname, header=None, names=[0,1,2], index_col=0, usecols='D,F,G', skiprows=30, nrows=6),
                            xl.parse('Pc_PrimDra', header=None, index_col=0, usecols='A:C',skiprows=34, nrows=6),
                            xl.parse('Pc_PrimDra', header=None, index_col=0, usecols='A:C', skiprows=45, nrows=4)
                            ])
                    additional_data.loc[['Depth (ft)', 'Texp (F)', 'Stress (psi)'], 2] = ['FT', 'F', 'PSI']
                    additional_data = additional_data \
                        .dropna(how='all') \
                        .transpose() \
                        .sort_index(ascending=False) \
                        .reset_index(drop=True) \
                        .rename(columns={'Texp (F)': 'Experimental Temperature',
                                         'Stress (psi)': 'Applied stress'})

                    analytical_analysis = xl.parse(tabname, header=5, usecols='A,B,H,I', nrows=15) \
                        .dropna(how='all') \
                        .reset_index(drop=True)

                    numerical_interp = pd.DataFrame()

                    raw_test = pd.DataFrame({'Actl. Cfg Speed': ['rpm'], 'Cum Time': ['s'], 'IP Production': ['cc']})
                    if tabname.upper()=='PC_PRIMDRA':
                        rawtest_tabname = 'Prod (PrimDra)'
                    elif tabname.upper()=='PC_IMB':
                        rawtest_tabname = 'Prod (Imb)'
                    for i in range(20):
                        raw_test = raw_test.append(xl.parse(rawtest_tabname, names=raw_test.columns, skiprows=21, usecols=[x + (5 * i) for x in [2,3,4]]))
                    raw_test = raw_test.dropna(how='any')\
                                       .reset_index(drop=True)
                    if raw_test.loc[0, 'Cum Time']=='s':
                        raw_test.loc[1:, 'Cum Time'] = raw_test.loc[1:, 'Cum Time'] / 3600
                        raw_test.loc[0, 'Cum Time'] = 'hour'

                    # Merge tables
                    df = pd.concat(
                        [user_data, additional_data, analytical_analysis, numerical_interp, raw_test],
                        axis=1)

                    tab_dfs.append(df)

            elif test == 'SCEN_DEACTIVATED': #DOC + SCEN
                """STYLE: DOC + WR"""

                tabname = 'KR_Centr'

                user_data = pd.concat([
                    xl.parse('Base Properties', header=None, names=[0, 1, 2], index_col=0, usecols='B:D', skiprows=2, nrows=3),
                    xl.parse('Base Properties', header=None, index_col=0, usecols='A:C', skiprows=8, nrows=3),
                    xl.parse('Base Properties', header=None, index_col=0, usecols='A:C', skiprows=16, nrows=1),
                    xl.parse('Base Properties', header=None, index_col=0, usecols='A:B', skiprows=44, nrows=1),
                    xl.parse(tabname, header=None, index_col=0, usecols='A:C', skiprows=6, nrows=8),
                    xl.parse(tabname, header=None, names=[0, 1, 2], index_col=0, usecols='D:F', skiprows=6, nrows=8),
                    xl.parse(tabname, header=None, names=[0, 1, 2], index_col=0, usecols='H:J', skiprows=6, nrows=1)])\
                    .dropna(how='all') \
                    .transpose() \
                    .sort_index(ascending=False) \
                    .reset_index(drop=True)

                user_data = pd.concat([user_data,
                                       xl.parse('Base Properties', names=['Experimental Temperature'],
                                                usecols='C', skiprows=31, nrows=2)], axis=1)\
                    .rename(columns={'Net Confining Stress (Centrifuge)': 'Applied stress'})
                user_data.loc[0, 'Experimental Temperature'] = 'F'

                additional_data = pd.concat([
                    xl.parse(tabname, header=None, index_col=0, usecols='A:C', skiprows=74, nrows=4),
                    xl.parse(tabname, header=None, names=[0, 1, 2], index_col=0, usecols='A,C,D', skiprows=79, nrows=6)])\
                    .dropna(how='all')\
                    .transpose()\
                    .sort_index(ascending=False)\
                    .reset_index(drop=True)

                analytical_analysis = xl.parse(tabname, header=22, usecols='G,J', ) \
                    .dropna(how='any') \
                    .reset_index(drop=True)

                numerical_interp = pd.concat([
                    xl.parse(tabname, header=None, names=[0, 1], index_col=0, usecols='H:I', skiprows=14, nrows=6),
                    xl.parse(tabname, header=None, names=[0, 1], index_col=0, usecols='J:K', skiprows=14, nrows=6)])\
                    .dropna(how='all')
                numerical_interp[2] = "UNITLESS"
                numerical_interp = numerical_interp.transpose().sort_index(ascending=False).reset_index(drop=True)

                numerical_interp2 = pd.concat([
                        xl.parse(tabname, header=39, usecols='Q,R')])\
                    .rename(columns={'S.1': 'S_NUM', 'kr.1': 'Kr_NUM'})\
                    .dropna(how='all')\
                    .reset_index(drop=True)
                numerical_interp2.loc[0, 'Kr_NUM'] = 'mD/mD'

                numerical_interp3 = pd.concat([
                            pd.DataFrame({"Sw": "V/V", "Pc": "BAR"}, index=[0]),
                            xl.parse('ImPc_Cfg', header=None, names=["Sw", "Pc"], usecols='W,Z', skiprows=34)])\
                    .dropna(how='all')\
                    .reset_index(drop=True)

                raw_test = xl.parse(tabname, header=22, usecols='A:B', nrows=50) \
                    .dropna(how='all') \
                    .reset_index(drop=True)

                # Merge tables
                df = pd.concat(
                    [user_data, additional_data, analytical_analysis, numerical_interp,
                     numerical_interp2, numerical_interp3, raw_test],
                    axis=1)

            elif test == 'SCEN_DEACTIVATED': #DOC + SCEN v2
                """STYLE: DOC + WR; EXAMPLES:"""

                tabname='KR_Centr'

                user_data = pd.concat([
                    xl.parse('Base Properties', header=None, names=[0, 1, 2], index_col=0, usecols='B:D', skiprows=4, nrows=1),
                    xl.parse('Base Properties', header=None, index_col=0, usecols='A:C', skiprows=8, nrows=3),
                    xl.parse('Base Properties', header=None, index_col=0, usecols='A:C', skiprows=16, nrows=1),
                    xl.parse('Base Properties', header=None, index_col=0, usecols='A:B', skiprows=44, nrows=1),
                    xl.parse(tabname, header=None, index_col=0, usecols='A:C', skiprows=2, nrows=14),
                    xl.parse(tabname, header=None, names=[0, 1, 2], index_col=0, usecols='D:F', skiprows=6, nrows=8),
                    xl.parse(tabname, header=None, names=[0, 1, 2], index_col=0, usecols='H:J', skiprows=6, nrows=5),
                ]) \
                    .dropna(how='all') \
                    .transpose() \
                    .sort_index(ascending=False) \
                    .reset_index(drop=True)

                user_data = pd.concat([user_data,
                                       xl.parse('Base Properties', names=['Experimental Temperature'],
                                                usecols='C', skiprows=31, nrows=2)], axis=1) \
                    .rename(columns={'Net Confining Stress (Centrifuge)': 'Applied stress'})
                user_data.loc[0, 'Experimental Temperature'] = 'F'

                additional_data = pd.concat([
                    xl.parse(tabname, header=None, index_col=0, usecols='A:C', skiprows=16, nrows=4),
                    xl.parse(tabname, header=None, names=[0, 1, 2], index_col=0, usecols='D,F,G', skiprows=14, nrows=6),
                    ]) \
                    .dropna(how='all') \
                    .transpose() \
                    .sort_index(ascending=False) \
                    .reset_index(drop=True)

                analytical_analysis = xl.parse(tabname, header=22, usecols='G,J', na_values=0) \
                    .dropna(how='any') \
                    .reset_index(drop=True)

                numerical_interp = pd.concat([
                    xl.parse(tabname, header=None, names=[0, 1], index_col=0, usecols='H:I', skiprows=14, nrows=5),
                    xl.parse(tabname, header=None, names=[0, 1], index_col=0, usecols='J:K', skiprows=14, nrows=5)]) \
                    .dropna(how='all')
                numerical_interp[2] = "UNITLESS"
                numerical_interp = numerical_interp \
                    .transpose() \
                    .sort_index(ascending=False) \
                    .reset_index(drop=True)

                numerical_interp2 = xl.parse(tabname, header=39, usecols='N,O') \
                    .rename(columns={'S': 'S_NUM', 'kr': 'Kr_NUM'}) \
                    .dropna(how='all') \
                    .reset_index(drop=True)
                numerical_interp2.loc[0, 'Kr_NUM'] = 'mD/mD'

                numerical_interp3 = xl.parse(tabname, header=39, usecols='Q,R') \
                    .rename(columns={'S.1': 'S_PC', 'Pc.1': 'Pc'}) \
                    .dropna(how='any') \
                    .reset_index(drop=True)

                raw_test = xl.parse(tabname, header=22, usecols='A:B') \
                    .dropna(how='all') \
                    .reset_index(drop=True)

                # Merge tables
                df = pd.concat(
                    [user_data, additional_data, analytical_analysis, numerical_interp,
                     numerical_interp2, numerical_interp3, raw_test],
                    axis=1)

            elif test == 'SCEN':
                """STYLE: Clean WR template for SCEN; EXAMPLES:"""

                tabname = 'Cfg_Kro'

                user_data = pd.concat([
                    xl.parse(tabname, header=None, index_col=0, usecols='A:C', skiprows=1, nrows=20),
                    xl.parse(tabname, header=None, names=[0,1,2], index_col=0, usecols='C:E', skiprows=1, nrows=3),
                    xl.parse(tabname, header=None, names=[0,1,2], index_col=0, usecols='D:F', skiprows=6, nrows=8),
                    xl.parse(tabname, header=None, names=[0,1,2], index_col=0, usecols='H:J', skiprows=1, nrows=12),
                    ])\
                    .dropna(how='all') \
                    .transpose() \
                    .sort_index(ascending=False) \
                    .reset_index(drop=True)

                additional_data = pd.concat([
                    xl.parse(tabname, header=None, index_col=0, usecols='A:C', skiprows=80, nrows=10),
                    xl.parse(tabname, header=None, names=[0,1,2], index_col=0, usecols='D,F,G', skiprows=14, nrows=6)
                    ])\
                    .dropna(how='all')\
                    .transpose()\
                    .sort_index(ascending=False)\
                    .reset_index(drop=True)
                additional_data.loc[0, ['Depth (ft)', 'Texp (F)', 'Stress (psi)']] = ['FT', 'F', 'PSI']
                additional_data.rename(columns={'Texp (F)': 'Experimental Temperature',
                                                'Stress (psi)': 'Applied stress'},
                                       inplace=True)

                analytical_analysis = xl.parse(tabname, header=22, usecols='G,J', na_values=0) \
                    .dropna(how='any') \
                    .reset_index(drop=True)

                numerical_interp1 = pd.concat([
                    xl.parse(tabname, header=None, names=[0, 1], index_col=0, usecols='H:I', skiprows=14, nrows=5),
                    xl.parse(tabname, header=None, names=[0, 1], index_col=0, usecols='J:K', skiprows=14, nrows=5)])\
                    .dropna(how='all')
                numerical_interp1[2] = "UNITLESS"
                numerical_interp1 = numerical_interp1.transpose().sort_index(ascending=False).reset_index(drop=True)

                numerical_interp2 = xl.parse(tabname, header=39, usecols='Q:R') \
                    .dropna(how='all') \
                    .reset_index(drop=True)
                numerical_interp2.columns = ['S_Pc','Pc']

                raw_test = xl.parse(tabname, header=22, usecols='A:B', nrows=55) \
                    .dropna(how='all') \
                    .reset_index(drop=True)

                # Merge tables
                df = pd.concat(
                    [user_data, additional_data, analytical_analysis, numerical_interp1, numerical_interp2, raw_test],
                    axis=1)

            elif test == 'SCEN_DEACTIVATED':
                """STYLE: Somewhat clean WR template for SCEN;"""

                tabs = [tab for tab in xl.sheet_names if tab.upper() in ['KRO_2NDIMB', 'KRO_IMB','KRW_DRA','KRW_2NDDRA']]
                is_B = 'B' in infile.name.upper()

                tab_dfs = []

                for tabname in tabs:

                    user_data = pd.concat([
                        xl.parse(tabname, header=None, index_col=0, usecols='A:C', skiprows=2, nrows=14),
                        xl.parse(tabname, header=None, names=[0,1,2], index_col=0, usecols='D:F', skiprows=6, nrows=8),
                        xl.parse(tabname, header=None, names=[0,1,2], index_col=0, usecols='H:J', skiprows=6, nrows=6),
                        ])\
                        .dropna(how='all') \
                        .transpose() \
                        .sort_index(ascending=False) \
                        .reset_index(drop=True)

                    if is_B:
                        additional_data = pd.concat([
                            xl.parse(tabname, header=None, index_col=0, usecols='A:C', skiprows=16, nrows=4),
                            xl.parse(tabname, header=None, names=[0, 1, 2], index_col=0, usecols='D,F,G', skiprows=14,
                                     nrows=6),
                            xl.parse('Test Info', header=None, names=[0, 1, 2], index_col=0, usecols='A:C',
                                     skiprows=1, nrows=15)
                        ])
                    else:
                        additional_data = pd.concat([
                            xl.parse(tabname, header=None, index_col=0, usecols='A:C', skiprows=16, nrows=4),
                            xl.parse(tabname, header=None, names=[0,1,2], index_col=0, usecols='D,F,G', skiprows=14, nrows=6),
                            xl.parse('Pc_PrimDra', header=None, names=[0,1,2], index_col=0, usecols='A:C',skiprows=34, nrows=6),
                            xl.parse('Pc_PrimDra', header=None, names=[0,1,2], index_col=0, usecols='A:C', skiprows=45, nrows=4)
                            ])
                    additional_data.loc[['Depth (ft)', 'Texp (F)', 'Stress (psi)'],2] = ['FT', 'F', 'PSI']
                    additional_data = additional_data\
                        .dropna(how='all') \
                        .transpose()\
                        .sort_index(ascending=False)\
                        .reset_index(drop=True)\
                        .rename(columns={'Texp (F)': 'Experimental Temperature',
                                                   'Stress (psi)': 'Applied stress'})

                    analytical_analysis = xl.parse(tabname, header=22, usecols='G,J', na_values=0) \
                        .dropna(how='any') \
                        .reset_index(drop=True)

                    numerical_interp = pd.concat([
                        xl.parse(tabname, header=None, names=[0, 1], index_col=0, usecols='H:I', skiprows=14, nrows=5),
                        xl.parse(tabname, header=None, names=[0, 1], index_col=0, usecols='J:K', skiprows=14, nrows=5)])\
                        .dropna(how='all')
                    numerical_interp[2] = "UNITLESS"
                    numerical_interp = numerical_interp\
                        .transpose()\
                        .sort_index(ascending=False)\
                        .reset_index(drop=True)

                    numerical_interp2 = xl.parse(tabname, header=39, usecols='N,O') \
                        .rename(columns={'S': 'S_NUM', 'kr': 'Kr_NUM'})                    \
                        .dropna(how='all') \
                        .reset_index(drop=True)
                    numerical_interp2.loc[0, 'Kr_NUM'] = 'mD/mD'

                    raw_test = xl.parse(tabname, header=22, usecols='A:B') \
                        .dropna(how='all') \
                        .reset_index(drop=True)

                    # Merge tables
                    df = pd.concat(
                        [user_data, additional_data, analytical_analysis, numerical_interp,numerical_interp2, raw_test],
                        axis=1)

                    tab_dfs.append(df)

            elif test == 'SS':  # Clean white Intertek. Example log file is
                # Parse all tables from input excel

                # Horizontal user_data
                user_data = pd.concat([
                    xl.parse('Base Properties', header=None, index_col=0, usecols='A:C', skiprows=6, nrows=20),
                    xl.parse('Base Properties', header=None, index_col=0, usecols='A:C', skiprows=42, nrows=1),
                    xl.parse('Base Properties', header=None, index_col=0, usecols='A:C', skiprows=46, nrows=4),
                    xl.parse('Rel Perms', header=None, names=[0, 1], index_col=0, usecols='D:E', skiprows=0, nrows=7),
                    xl.parse('Rel Perms', header=None, names=[0, 1], index_col=0, usecols='F:G', skiprows=0, nrows=7)
                ]) \
                    .dropna(how='all') \
                    .transpose() \
                    .sort_index(ascending=False) \
                    .reset_index(drop=True)

                # Carefully formating values from Fluid Properties table in Base Properties tab
                user_data = pd.concat([user_data,
                                       xl.parse('Base Properties', names=['Experimental Temperature'],
                                                usecols='C', skiprows=31, nrows=2),
                                       xl.parse('Base Properties', names=['Oil Viscosity', 'Oil Density'],
                                                usecols='D:E', skiprows=31, nrows=2),
                                       xl.parse('Base Properties', names=['Brine Viscosity', 'Brine Density'],
                                                usecols='D:E', skiprows=33, nrows=2)], axis=1)

                user_data.rename(columns={'Net Confining Stress (Steady State)': 'Applied stress'}, inplace=True)
                user_data.loc[0, 'Experimental Temperature'] = 'F'
                user_data.loc[0, ['Oil Viscosity', 'Water Viscosity']] = 'cP'
                user_data.loc[0, ['Oil Density', 'Water Density']] = 'gm/cc'

                additional_data = xl.parse('Rel Perms', header=None, index_col=0, usecols='A:C', skiprows=23, nrows=10) \
                    .dropna(how='all') \
                    .transpose() \
                    .reset_index(drop=True)

                analytical_analysis = xl.parse('Rel Perms', header=10, usecols='I:K') \
                    .dropna(how='all') \
                    .reset_index(drop=True) \
                    # .add_sufix('_ANA')
                analytical_analysis.columns = ['Sw_ANA', 'krw_ANA', 'kro_ANA']

                numerical_interp1 = pd.concat([
                    xl.parse('Rel Perms', header=None, names=[0, 1], index_col=0, usecols='M:N', skiprows=5, nrows=4),
                    xl.parse('Rel Perms', header=None, names=[0, 1], index_col=0, usecols='O:P', skiprows=5, nrows=4)]) \
                    .dropna(how='all')
                numerical_interp1[2] = "UNITLESS"
                numerical_interp1 = numerical_interp1 \
                    .transpose() \
                    .sort_index(ascending=False) \
                    .reset_index(drop=True)

                numerical_interp2 = xl.parse('Rel Perms', header=10, usecols='M:O, Q:R', nrows=15) \
                    .dropna(how='all') \
                    .reset_index(drop=True) \
                    # .add_sufix('_NUM')
                numerical_interp2.columns = ['Sw_NUM', 'krw_NUM', 'kro_NUM', 'S_PC', 'Pc']

                raw_test = xl.parse('Delta P vs Time', header=7, usecols='A:D') \
                    .dropna(how='all') \
                    .reset_index(drop=True)
                #raw_test.columns = [raw_test.iloc[0], raw_test.iloc[1]]
                raw_test.columns = raw_test.iloc[0] + ' ' + raw_test.iloc[1]
                raw_test = raw_test.iloc[2:].reset_index(drop=True)

                sat_prof = xl.parse('Saturation Scans', header=None, index_col=0, skiprows=1, nrows=12) \
                    .dropna(how='all') \
                    .transpose() \
                    .sort_index(ascending=False) \
                    .reset_index(drop=True)

                # Merge tables
                df = pd.concat(
                    [user_data, additional_data, analytical_analysis, numerical_interp1, numerical_interp2,
                     raw_test, sat_prof],
                    axis=1)

        # df = pd.concat(tab_dfs) # comment out if not dealing with multiple tabs loop

        # Normalize columns to uppercase
        df.rename(columns=lambda x: str(x).strip().upper(), inplace=True)

        # Clear units row
        df.iloc[0] = df.iloc[0].str.strip('()[]').str.upper()

        # self.samples.append(df)
        return df

    def prepare_batch(self, folder, test, template):
        samples = []

        for infile in os.scandir(folder):
            if not infile.name[-5:] in ('.xlsx', '.xlsm'):
                continue
            print(f'\nParsing {infile.name}')
            df = self.prepare_sample(infile, test, template)
            samples.append(df)

        samples_df = pd.concat(samples)

        nunique_units = samples_df.loc[0].nunique()
        if (nunique_units > 1).all():
            print("\nEncountered different units between samples!")
            print(samples_df.loc[0, nunique_units[nunique_units > 1].index])
        else:
            samples_df = samples_df.iloc[:1].append(samples_df.drop(0)).reset_index(drop=True)

        return samples_df

    def get_conditions(self, df):
        # Save temperature and pressure values
        temperature = df.loc[:, 'EXPERIMENTAL TEMPERATURE']
        if temperature.isna().loc[1]:
            temperature = 'UNK'
        elif temperature.loc[0] in ('F', 'FT'):
            temperature = round(5/9 * (float(temperature.loc[1])-32.0))
        else:
            temperature = round(temperature.loc[1])
        self.temperature = temperature

        pressure = df.loc[:, 'APPLIED STRESS']
        if pressure.isna().loc[1]:
            pressure = 'UNK'
        elif str(pressure.loc[1]).upper() in ['AMBIENT', 'AMB']:
            pressure = 'AMB'
        elif pressure.loc[0] == 'BAR':
            pressure = round(float(pressure.loc[1]) * 14.50377)
        elif pressure.loc[0] == 'PSI':
            pressure = round(pressure.loc[1])
        else:
            print(f"Unexpected stress unit: {pressure.loc[1]}. Check column headers.")
            pressure = pressure.loc[1]
        self.pressure = pressure

    @staticmethod
    def merge_duplicated_columns(df):
        # Selecting non unique columns names
        dup_colnames = set()
        colnames = sorted(df.columns)
        for ind in range(1, len(colnames)):
            if colnames[ind] == colnames[ind - 1]:
                dup_colnames.add(colnames[ind])

        df.rename(columns=Renamer(), inplace=True)

        for colname in dup_colnames:
            for dcol in [col for col in df if col.startswith(f'{colname}_')]:
                df[colname].fillna(df[dcol], inplace=True)
                df.drop(columns=dcol, inplace=True)

        return df

    def create_log(self, df, test, template):
        # Create logs by concatenating curves to log template
        mnemonics = self.mnemonics[template][test]
        template = self.log_templates[test]
        log_columns = list(template.columns)

        self.get_conditions(df)
        tmp = self.temperature
        pr = self.pressure

        # Add renamed and deduplicated columns to the template template
        df = df.rename(columns=mnemonics)
        df = self.merge_duplicated_columns(df)
        log = pd.concat([template, df]).reset_index(drop=True)


        # Remove undesired columns
        log = log.loc[:, log_columns]

        # Fill TTTT and XXXX temperature and pressure
        log.rename(columns=lambda x: x.replace('TTTT', str(tmp)), inplace=True)
        log.rename(columns=lambda x: x.replace('XXXX', str(pr)), inplace=True)
        # log.rename(columns=lambda x: re.sub('TTTT$', temperature[1], x), inplace=True)
        # log.rename(columns=lambda x: re.sub('XXXX$', pressure[1], x), inplace=True)

        # Work in progress
        # Catch unexpected units
        with warnings.catch_warnings(record=True) as w:
            for name, items in log.iloc[[0, 1], :].items():
                if items[0] != items[1] \
                        and items[1] == items[1] \
                        and items[1] != 'UNITLESS' \
                        and items[1] not in self.accepted_units.get(items[0], []):
                    # Checks if units are different from template, are not nan, are not UNITLESS and are not in the list of accepted units
                    warnings.warn(f'{name}: Parsed unit {items[1]} is not matching expected unit {items[0]}',
                                  UserWarning)
                    print(w[-1])
            if len(w) > 0:
                print(f'For {len(w)} curves the unit is not as expected. Inspect units manually in the produced log. '
                      'Original units are stored in first row and expected units in second row. Recalculate column if needed'
                      'and delete obsolete unit row.')
            #else:
                #log = log.drop(1).reset_index(drop=True)

        log = log.drop(1).reset_index(drop=True)  # temp

        # Amend metadata
        #log['CREP_LAB_NAME'] = pd.Series([log.loc[0, 'CREP_LAB_NAME'], self.lab_names[log.loc[1, 'CREP_LAB_NAME']]])
        log.loc[1, 'CREP_LAB_NAME'] = self.lab_names[log.loc[1, 'CREP_LAB_NAME']] if log.loc[1, 'CREP_LAB_NAME'] not in ('', np.nan) else ''
        log.loc[2:, 'CRPE_LAB_NAME'] = np.nan
        log.loc[1, 'CREP_TESTTYPE'] = 'RELPERM'
        log.loc[1, 'CREP_SAMPLETYPE'] = 'CRPLUG'

        # self.log = log
        return log

    @staticmethod
    def fix_duplicated_depth(log, test):
        """Increments 0.00001 in case of duplciated depths between samples"""
        duplicated = log['DEPTH'].notna() & log['DEPTH'].duplicated()
        if duplicated.any():
            print(f"Incrementing {duplicated.sum()} duplicated DEPTHs in {test} log")
        while duplicated.any():
            log['DEPTH'].loc[duplicated] = log['DEPTH'].loc[duplicated].astype(float) + 0.00001
            duplicated = log['DEPTH'].notna() & log['DEPTH'].duplicated()
        return log

    def split_to_csv(self, log, folder, test, drop_empty=False):

        SAVENAME = f"{folder}\CORE_RELPERM_{test}_{log.loc[1, 'CREP_SAMPLETYPE']}_{log.loc[1, 'CREP_LAB_NAME']}_LOG"

        N_METADATA_COLS = 7
        METADATA_COLS = ['CREP_LAB_NAME', 'CREP_TESTTYPE', 'CREP_SAMPLETYPE',
                         'CREP_NO', 'CREP_TEST_DATE', 'SAMPNM', 'DEPTH']


        if test == 'MCEN':

            ANA_COLS = ['TIME_CEN_EQ', 'SPEED_CEN_EQ', 'PROD_CEN_EQ', 'SAT_CEN_INFLOW_HB', 'CAP_CEN_INFLOW_HB']
            NUM_COLS = ['SAT_PC_CEN_NUM', 'CAP_CEN_NUM']
            RAW_COLS = ['TIME_CEN', 'SPEED_CEN', 'PROD_CEN', 'TEMP_CEN']

        elif test == 'SCEN':
            ANA_COLS = ['SAT_CEN_END_HAGOORT', 'RLP_CEN_END_HAGOORT']
            NUM_COLS = ['SAT_CEN_NUM', 'RLP_CEN_NUM', 'SAT_PC_CEN_NUM', 'CAP_CEN_NUM']
            RAW_COLS = ['TIME_CEN', 'SPEED_CEN', 'PROD_CEN', 'TEMP_CEN']

        elif test == 'SS':
            ANA_COLS = ['TIME_SS_EQ', 'DP_SS_EQ', 'SSRATE_WAT_EQ', 'SSRATE_OIL_EQ',
                        'SAT_SS_AVG_EQ', 'RLP_WAT_ANA', 'RLP_OIL_ANA']
            NUM_COLS = ['SAT_SS_NUM', 'RLP_WAT_NUM', 'RLP_OIL_NUM', 'SAT_SS_PC_NUM', 'CAP_SS_PC_NUM']
            RAW_COLS = ['TIME_SS', 'DP_SS', 'SSRATE_WAT', 'SSRATE_OIL', 'SAT_SS_AVG', 'TEMP_SS']
            SPR_COLS = ['DIST_IN_CORE', 'XN_BASE_WAT', 'XN_BASE_OIL', 'XN_FW_0', 'XN_FW_1', 'XN_FW_2', 'XN_FW_3',
                        'XN_FW_4', 'XN_FW_5', 'XN_FW_6', 'XN_FW_7', 'XN_FW_8', 'XN_FW_9', 'XN_FW_10', 'XN_FW_11',
                        'XN_FW_12', 'XN_FW_BUMP_1', 'XN_FW_BUMP_2', 'SATX_SS_FW_0', 'SATX_SS_FW_1', 'SATX_SS_FW_2',
                        'SATX_SS_FW_3', 'SATX_SS_FW_4', 'SATX_SS_FW_5', 'SATX_SS_FW_6', 'SATX_SS_FW_7', 'SATX_SS_FW_8',
                        'SATX_SS_FW_9', 'SATX_SS_FW_10', 'SATX_SS_FW_11', 'SATX_SS_FW_12', 'SATX_SS_FW_BUMP_1',
                        'SATX_SS_FW_BUMP_2']

        # Incerement duplicated depth
        log = self.fix_duplicated_depth(log, test)

        if test == 'SS':
            log_scalar = log.drop(ANA_COLS + NUM_COLS + RAW_COLS + SPR_COLS, axis=1).dropna(how='all')
            log_spr = log[METADATA_COLS + SPR_COLS].dropna(how='all')
        else:
            log_scalar = log.drop(ANA_COLS + NUM_COLS + RAW_COLS, axis=1).dropna(how='all')

        log_ana = log[METADATA_COLS + ANA_COLS].dropna(how='all')
        log_num = log[METADATA_COLS + NUM_COLS].dropna(how='all')
        log_raw = log[METADATA_COLS + RAW_COLS].dropna(how='all')

        if drop_empty:
            log = self.drop_empty_columns(log)
            log_scalar = self.drop_empty_columns(log_scalar)
            log_ana = self.drop_empty_columns(log_ana)
            log_num = self.drop_empty_columns(log_num)
            log_raw = self.drop_empty_columns(log_raw)
            if test == 'SS':
                log_spr = self.drop_empty_columns(log_spr)

        log.to_csv(f"{SAVENAME}.csv", index=False)
        log_scalar.to_csv(f"{SAVENAME}_SCA.csv", index=False)
        self.split_to_parts(log_scalar, SAVENAME, suffix='SCA', metadata_cols=N_METADATA_COLS, part_cols=8)
        log_ana.to_csv(f"{SAVENAME}_ANA.csv",index=False)
        log_num.to_csv(f"{SAVENAME}_NUM.csv",index=False)
        log_raw.to_csv(f"{SAVENAME}_RAW.csv",index=False)
        if test=='SS':
            log_spr.to_csv(f"{SAVENAME}_SPR.csv", index=False)
            self.split_to_parts(log_spr, SAVENAME, suffix='SPR', metadata_cols=N_METADATA_COLS, part_cols=6)

    @staticmethod
    def split_to_parts(log, name, suffix='', metadata_cols=7, part_cols=8):
        i = 1
        first_col = metadata_cols
        total_cols = log.shape[1]

        while first_col < total_cols:

            last_col = first_col + part_cols
            savename = f"{name}_{suffix} ({i}).csv"

            if last_col < total_cols:
                log.iloc[:, np.r_[:metadata_cols, first_col:last_col]].to_csv(savename, index=False)
            else:
                log.iloc[:, np.r_[:metadata_cols, first_col:total_cols]].to_csv(savename, index=False)

            first_col += part_cols
            i += 1

    @staticmethod
    def drop_empty_columns(log):
        # Deletes columns from log if they are empty below units row
        excluded = ['CREP_LAB_NAME', 'CREP_TESTTYPE', 'CREP_SAMPLETYPE',
                    'CREP_NO', 'CREP_TEST_DATE', 'SAMPNM', 'DEPTH']
        to_drop = []
        for name, data in log.iteritems():
            if name not in excluded and data.iloc[1:].isnull().all():
                to_drop.append(name)
        return log.drop(columns=to_drop).dropna(how='all').reset_index(drop=True)

class Renamer():
    """
    https://stackoverflow.com/questions/40774787/renaming-columns-in-a-pandas-dataframe-with-duplicate-column-names
    """
    def __init__(self):
        self.d = dict()

    def __call__(self, x):
        if x not in self.d:
            self.d[x] = 0
            return x
        else:
            self.d[x] += 1
            return "%s_%d" % (x, self.d[x])

# log.to_csv(r'Out/centrifuge_out_log.csv', index=False)
# print('Finished!')

# add fail safe only 40k rows allowed

rp = RelPerm()
df = rp.prepare_batch(FOLDER, TEST, TEMPLATE)
#df = rp.prepare_sample(INFILE, TEST, TEMPLATE)
log = rp.create_log(df, TEST, TEMPLATE)
rp.split_to_csv(log, FOLDER, TEST, drop_empty=True)
