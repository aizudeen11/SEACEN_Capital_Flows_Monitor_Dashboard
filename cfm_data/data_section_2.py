import pandas as pd
from datetime import datetime
import os
import re
import warnings
from pathlib import Path
from dateutil.relativedelta import relativedelta
from dotenv import load_dotenv

HERE = Path(__file__).resolve().parent
# The paths come from the .env file in this folder, wherever the script is run from.
load_dotenv(HERE / ".env")

data_path = os.getenv("DATA")
automation_path = os.getenv("AUTOMATION")

# The compiled BoP / IIP workbooks ("IMF BOP Quarterly.xlsx", "CB IIP Annual.xlsx", ...):
# one workbook per dataset, one sheet per economy, named like "Korea (KR)".
# They sit in this folder unless IIPBOP in .env points elsewhere.
BOP_IIP_DIR = Path(os.getenv("IIPBOP") or HERE / "BOP & IIP data")

# The economies kept by default, spelled as in the workbooks' sheet names
# (without the code in brackets). The IMF workbooks also hold non-SEACEN
# economies (e.g. Brazil, the G7), which are left out.
SEACEN_ECONOMIES = [
    "Brunei", "Cambodia", "China", "Chinese Taipei", "Hong Kong", "India", "Indonesia", "Korea",
    "Lao PDR", "Malaysia", "Mongolia", "Myanmar", "Nepal", "Papua New Guinea", "Philippines",
    "Singapore", "Sri Lanka", "Thailand", "Vietnam",
]

# Chinese Taipei isn't in the IMF workbooks; its data comes from the
# "Chinese Taipei (TW)" sheet of these central-bank workbooks instead. Its
# annual IIP also fills the quarterly IIP tables (each quarter shows that
# year's position).
TAIWAN_WORKBOOKS = {
    "BoP Quarterly": "CB BOP Quarterly",
    "BoP Annual": "CB BOP Annual",
    "IIP": "CB IIP Annual",
}

warnings.filterwarnings("ignore")


def check_paths() -> None:
    """Stops with a clear message if a path this module needs is missing from
    .env or doesn't exist."""
    paths = {"DATA": data_path, "AUTOMATION": automation_path}
    missing = [name for name, value in paths.items() if not value]
    if missing:
        raise SystemExit(f"Section 2: {', '.join(missing)} not set in {HERE / '.env'}")
    not_found = [name for name, value in paths.items() if not Path(value).exists()]
    if not BOP_IIP_DIR.is_dir():
        not_found.append(f"IIPBOP ({BOP_IIP_DIR})")
    if not_found:
        raise SystemExit(f"Section 2: the path in {', '.join(not_found)} doesn't exist")


def remove_non_seacen(sheet_names: list[str], economies: list[str] = SEACEN_ECONOMIES) -> list[str]:
    """The sheets (named like "Korea (KR)") whose economy is in `economies`;
    by default the SEACEN economies."""
    return [s for s in sheet_names if re.sub(r"\s*\([A-Z]{2}\)\s*$", "", s).strip() in economies]


def read_bop_iip(dataset: str, economies: list[str] = SEACEN_ECONOMIES) -> pd.DataFrame:
    """All economy sheets of one compiled workbook (e.g. "IMF BOP Quarterly"),
    stacked into one table, keeping only `economies` (by default the SEACEN
    economies). The first column, the item description, is named "Type"."""
    path = BOP_IIP_DIR / f"{dataset}.xlsx"
    sheets = remove_non_seacen(pd.ExcelFile(path).sheet_names, economies)
    if not sheets:
        raise SystemExit(f"Section 2: no SEACEN economy sheets found in {path}")
    frames = []
    for sheet in sheets:
        df = pd.read_excel(path, sheet_name=sheet)
        frames.append(df.rename(columns={df.columns[0]: "Type"}))
    return pd.concat(frames)


def read_taiwan(dataset: str) -> pd.DataFrame:
    """The Chinese Taipei sheet of the workbook TAIWAN_WORKBOOKS gives for
    `dataset`, as downloaded (the item description in the first column)."""
    path = BOP_IIP_DIR / f"{TAIWAN_WORKBOOKS[dataset]}.xlsx"
    sheets = remove_non_seacen(pd.ExcelFile(path).sheet_names, ["Chinese Taipei"])
    if not sheets:
        raise SystemExit(f"Section 2: no Chinese Taipei sheet found in {path}")
    return pd.read_excel(path, sheet_name=sheets[0])

def bop_quarterly(year_use:int = 2019) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    df_init = pd.read_excel(automation_path, sheet_name='BoP edit')

    df_all_annual = read_bop_iip('IMF BOP Quarterly')
    df_all_annual = df_all_annual[df_all_annual['Type'].isin(df_init['desc'])]


    df_all_annual_2 = df_all_annual[df_all_annual.columns[:4].to_list() + [dt for dt in df_all_annual.columns[4:].sort_values(ascending=True).to_list() if dt >= datetime(year_use, 1, 1)]]

    def sg_row(df:pd.DataFrame, main:bool = True) -> pd.DataFrame:
        var = df.columns[:4] if main else df.columns[4:]
        return df[var]

    sg_add = {
    'FDI Assets':{'FDI Equity Assets': 0.74,'FDI Debt Assets': 0.26},
    'FDI Liabilities': {'FDI Equity Liabilities': 0.85,'FDI Debt Liabilities': 0.15},
    'Portfolio Assets': {'Portfolio Equity Assets': 0.57, 'Portfolio Debt Assets': 0.43}, 
    'Portfolio Liabilities': {'Portfolio Equity Liabilities': 0.67, 'Portfolio Debt Liabilities': 0.33},
    'Other Investment Assets': {'Currency and Deposits Assets': 0.5, 'Loans Assets': 0.29, 'Insurance and Pension Assets': 0.02, 'Trade Credits and Advances Assets': 0.09, 'OI Others Assets': 0.1},
    'Other Investment Liabilities': {'Currency and Deposits Liabilities': 0.5, 'Loans Liabilities': 0.29, 'Insurance and Pension Liabilities': 0.02, 'Trade Credits and Advances Liabilities': 0.09, 'OI Others Liabilities': 0.1},
    }

    m = df_all_annual_2[(df_all_annual_2['Region']== 'Singapore')]['Type'].to_list()
    df_init[(df_init['desc'].isin(m))]['group'].unique().tolist()

    sg_item = [ 'BoP: Financial Account: Direct Investment: Assets',
                'BoP: Financial Account: Direct Investment: Liabilities',
                'BoP: Financial Account: Portfolio Investment: Assets',
                'BoP: Financial Account: Portfolio Investment: Liabilities',
                'BoP: Financial Account: Other Investment: Assets',
                'BoP: Financial Account: Other Investment: Liabilities']

    sg_df = pd.DataFrame(columns=df_all_annual_2.columns)

    # ensure the zip between sg_item and sg_add is correct
    for x, y in zip(sg_item, sg_add.values()):
        multiplier = list(y.values())
        row_type = list(y.keys())
        df_temp = df_all_annual_2[(df_all_annual_2['Type'] == x) &(df_all_annual_2['Region']== 'Singapore')]
        # print(multiplier, row_type)
        for n in range(len(row_type)):
            col1 = sg_row(df_temp, main=True)
            col1.iat[0, 0] = list(df_init[(df_init['short_title']==row_type[n])]['desc'])[0]
            col2 = sg_row(df_temp, main=False) * multiplier[n]
            merged = pd.concat([col1, col2], axis=1)
            sg_df = pd.concat([sg_df, merged], axis=0)

    df_init_tw = pd.read_excel(automation_path, sheet_name='Taiwan BoP')
    df_tw = read_taiwan("BoP Quarterly")
    df_tw.rename(columns={df_tw.columns[0]: 'Type'}, inplace=True)
    df_tw = df_tw[df_tw['Type'].isin(df_init_tw['desc'])]
    df_tw.replace('Taiwan (China)', 'Taiwan', inplace=True)
    for x, y in enumerate(df_tw['Type'].to_list()):
        df_tw.iat[x,0] = list(df_init_tw[(df_init_tw['desc'] == y)]['desc2'])[0]

    selected_columns = [col for col in df_tw.columns[4:] if pd.to_datetime(col) >= datetime(year_use, 1, 1)]
    final_columns = list(df_tw.columns[:4]) + selected_columns
    df_tw = df_tw[final_columns]

    # concat to the main df
    df_all_annual_2 = pd.concat([df_all_annual_2, sg_df, df_tw], axis=0)

    seacen_country = ['Papua New Guinea', 'Vietnam', 'Nepal', 'India', 'Indonesia', 'Laos', 'Sri Lanka', 'Hong Kong SAR (China)', 'Philippines', 'Taiwan', 'Malaysia', 'Mongolia', 'China', 'Cambodia', 'Thailand', 'Singapore', 'South Korea', 'Brunei', 'Myanmar']

    for x in df_all_annual_2['Type'].values:
        var1 = list(set(df_all_annual_2[(df_all_annual_2['Type'] == x)]['Region'].values).symmetric_difference(set(seacen_country)))
        if len(var1) > 0:
            for y in var1:
                df_temp1 = pd.DataFrame({"Type": [x], "Region": [y]})
                df_all_annual_2 = pd.concat([df_all_annual_2, df_temp1], axis=0)
        else:
            continue
    df_all_annual_2.replace('South Korea', 'Korea', inplace=True)
    df_all_annual_2.sort_values(by=['Type', 'Region'], inplace=True)
    df_all_annual_2.reset_index(drop=True, inplace=True)

    init_dict = df_init.to_dict()
    df_dict = dict()
    for x in init_dict['desc'].values():
        df_temp = df_all_annual_2[(df_all_annual_2[df_all_annual_2.columns[0]] == x)].reset_index(drop=True)
        df_dict[x] = df_temp

    # calculate the sum of the columns (1H and 2H) for each year
    num1, num2 = 0, 1
    df_col = df_all_annual_2.columns[4:][:-1]
    df_temp_col_1 = df_all_annual_2[df_all_annual_2.columns[:4]].reset_index(drop=True)

    for x in range(len(df_col)//2):
        df_col_temp = (df_all_annual_2[df_col[num1]] + df_all_annual_2[df_col[num1 + 1]]).div(1000)
        year = df_col[num1].year
        df_col_temp.reset_index(drop=True, inplace=True)
        df_temp_col_1[f'{num2}H{year}'] = df_col_temp
        num1 += 2
        num2 += 1
        if num2 == 3:
            num2 = 1

    df_temp_col_1.drop(columns=['Last Update Time', 'Unit'], inplace=True)

    calc_dict = {'Asia-Pacific': seacen_country,
             'Asian Advance Economies': ['Hong Kong SAR (China)', 'South Korea', 'Singapore', 'Taiwan'],
             'ASEAN5 Economies': ['Indonesia', 'Malaysia', 'Philippines', 'Thailand', 'Vietnam'],
             'Asian EDMEs': ['Brunei', 'Cambodia', 'Laos', 'Mongolia', 'Nepal', 'Sri Lanka', 'Papua New Guinea', 'Myanmar'],
             'China': ['China'],
             'India': ['India'],
             }

    ###################### this is for equity ######################

    asset_row = df_init.iloc[[0,4,5,6, 8, 9,10,11,12,13,14]]['desc'].to_list()
    asset_row2 = ['Direct Investment', 'Portfolio Equity', 'Portfolio Debt','Financial Derivatives', 'Other Equity','Currency & Deposits','Loans','Insurance, Pension & Standardized Guarantee Schemes'
                ,'Trade Credit & Advances','Other Accounts Receivable','Reserve Assets', ]
    sum_rows = df_init[(df_init['short_title'].isin(['OI Equity Assets','Insurance and Pension Assets','OI Others Assets',]))]['desc'].to_list()
    order = ['Direct Investment','Portfolio Equity','Portfolio Debt','Financial Derivatives','Currency & Deposits','Loans','Trade Credit & Advances','Other Investments','Reserve Assets',]
    df_by_region_main = pd.DataFrame(columns=df_temp_col_1.columns)

    for x,y in calc_dict.items():
        # print(x,y)
        df_by_region_temp = df_temp_col_1[(df_temp_col_1['Region'].isin(y)) & (df_temp_col_1['Type'].isin(asset_row))]
        df_by_region_temp = df_by_region_temp.groupby(['Type']).sum()
        df_by_region_temp.loc[str(x) + ' Total'] = df_by_region_temp.sum()
        df_by_region_temp.loc['Other Investments'] = df_by_region_temp.loc[sum_rows].sum()
        df_by_region_temp.reset_index(inplace = True)
        df_by_region_temp = df_by_region_temp[(~df_by_region_temp['Type'].isin(sum_rows))] #####  the symbol of '~' use similar like notin, the oposite of isin. this is build in python operator
        df_by_region_temp['Type'] = df_by_region_temp['Type'].replace(asset_row, asset_row2)
        df_by_region_temp = df_by_region_temp.sort_values(by='Type', key=lambda x: x.map({k:i for i, k in enumerate(order)}))
        df_by_region_temp['Region'] = x

        df_by_region_main = pd.concat([df_by_region_main, df_by_region_temp], axis = 0)  


    df_by_region_main.reset_index(inplace = True, drop=True)
    df_by_region_main.insert(loc=0, column='Group', value='Assets')

    ###################### this is for equity ######################

    ###################### this is for debt ######################

    debt_row = df_init.iloc[[20,24,25,26,28,29,30,31,32,33,34]]['desc'].to_list()
    debt_row2 = ['Direct Investment', 'Portfolio Equity', 'Portfolio Debt','Financial Derivatives', 'Other Equity','Currency & Deposits','Loans','Insurance, Pension & Standardized Guarantee Schemes'
                ,'Trade Credit & Advances','Other Accounts Payable','SDR Liabilities', ]
    sum_rows2 = df_init[(df_init['short_title'].isin(['OI Equity Liabilities', 'Insurance and Pension Liabilities', 'OI Others Liabilities', 'SDR Liabilities']))]['desc'].to_list()
    order2 = ['Direct Investment','Portfolio Equity','Portfolio Debt','Financial Derivatives','Currency & Deposits','Loans','Trade Credit & Advances','Other Investments']
    df_by_region_main2 = pd.DataFrame(columns=df_temp_col_1.columns)

    for x,y in calc_dict.items():
        # print(x,y)
        df_by_region_temp = df_temp_col_1[(df_temp_col_1['Region'].isin(y)) & (df_temp_col_1['Type'].isin(debt_row))]
        df_by_region_temp = df_by_region_temp.groupby(['Type']).sum()
        df_by_region_temp.loc[str(x) + ' Total'] = df_by_region_temp.sum()
        df_by_region_temp.loc['Other Investments'] = df_by_region_temp.loc[sum_rows2].sum()
        df_by_region_temp.reset_index(inplace = True)
        df_by_region_temp = df_by_region_temp[(~df_by_region_temp['Type'].isin(sum_rows2))] #####  the symbol of '~' use similar like notin, the oposite of isin. this is build in python operator
        df_by_region_temp['Type'] = df_by_region_temp['Type'].replace(debt_row, debt_row2)
        df_by_region_temp = df_by_region_temp.sort_values(by='Type', key=lambda x: x.map({k:i for i, k in enumerate(order2)}))
        df_by_region_temp['Region'] = x

        df_by_region_main2 = pd.concat([df_by_region_main2, df_by_region_temp], axis = 0)  

    df_by_region_main2.reset_index(inplace = True, drop=True)
    df_by_region_main2.insert(loc=0, column='Group', value='Liabilities')

    ###################### this is for debt ######################

    ###################### this is for caq data ######################

    caq_data = df_temp_col_1[(df_temp_col_1['Type'].isin(df_init[(df_init['group'] == 'Current Account')]['desc'].to_list()[1:]))]
    caq_data_calc = caq_data.groupby(['Type']).sum()
    caq_data_calc.loc['Current Account Balance'] = caq_data_calc.sum()
    caq_data_calc['Region'] = 'Asia-Pacific'
    caq_data_calc.reset_index(inplace=True)
    rename_to = ['Goods Trade', 'Services Trade','Primary Income','Secondary Income',]
    rename_from = ['BoP: Current Account: Goods','BoP: Current Account: Services','BoP: Current Account: Primary Income','BoP: Current Account: Secondary Income',]
    caq_data_calc['Type'] = caq_data_calc['Type'].replace(rename_from, rename_to)
    caq_data_calc = caq_data_calc.sort_values(by='Type', key=lambda x: x.map({k:i for i, k in enumerate(rename_to)}))
    caq_data_calc.reset_index(drop=True, inplace=True)
    caq_data_calc.insert(loc=0, column='Group', value='Current Account')

    ###################### this is for caq data ######################

    df_main = pd.concat([df_by_region_main, df_by_region_main2, caq_data_calc], axis = 0) ## combine
    df_main.reset_index(drop=True, inplace=True)
    df_all_annual_2.replace({k:v for (k,v) in zip(df_init['desc'].to_list(), df_init['short_title'].to_list())}, inplace=True)
    df_temp_col_1.replace({k:v for (k,v) in zip(df_init['desc'].to_list(), df_init['short_title'].to_list())}, inplace=True)

    return df_temp_col_1, df_all_annual_2, df_main

def bop_annual(year_use:int = 2019) -> tuple[pd.DataFrame, pd.DataFrame]:

    df_init = pd.read_excel(automation_path, sheet_name='BoP edit')

    df_all_annual = read_bop_iip('IMF BOP Annual')
    df_all_annual = df_all_annual[df_all_annual['Type'].isin(df_init['desc'])]


    df_all_annual_2 = df_all_annual[df_all_annual.columns[:4].to_list() + [dt for dt in df_all_annual.columns[4:].sort_values(ascending=True).to_list() if dt >= datetime(year_use, 1, 1)]]

    df_init_tw = pd.read_excel(automation_path, sheet_name='Taiwan BoP')
    df_tw = read_taiwan("BoP Annual")
    df_tw = df_tw.map(lambda x: x[:x.find(' (f')] if isinstance(x, str) and ' (f' in x else x)
    df_tw.rename(columns={df_tw.columns[0]: 'Type'}, inplace=True)
    # The annual sheet leaves the code off a few items ("BoP: CU: GS: Services"
    # for "BoP: CU: GS: Services (SV)"), so match those without it.
    def no_code(item):
        return re.sub(r"\s*\([A-Z]{2,}\)$", "", str(item).strip())
    by_no_code = {no_code(d): d for d in df_init_tw['desc']}
    df_tw['Type'] = [t if t in set(df_init_tw['desc']) else by_no_code.get(no_code(t), t) for t in df_tw['Type']]
    df_tw = df_tw[df_tw['Type'].isin(df_init_tw['desc'])]

    for x, y in enumerate(df_tw['Type'].to_list()):
        df_tw.iat[x,0] = list(df_init_tw[(df_init_tw['desc'] == y)]['desc2'])[0]

    selected_columns = [col for col in df_tw.columns[4:] if pd.to_datetime(col) >= datetime(year_use, 1, 1)]
    final_columns = list(df_tw.columns[:4]) + selected_columns
    df_tw = df_tw[final_columns]
    df_tw.replace('Taiwan (China)', 'Taiwan', inplace=True)

    # concat to the main df
    df_all_annual_2 = pd.concat([df_all_annual_2, df_tw], axis=0)

    seacen_country = ['Papua New Guinea', 'Vietnam', 'Nepal', 'India', 'Indonesia', 'Laos', 'Sri Lanka', 'Hong Kong SAR (China)', 'Philippines', 'Taiwan', 'Malaysia', 'Mongolia', 'China', 'Cambodia', 'Thailand', 'Singapore', 'South Korea', 'Brunei', 'Myanmar']

    for x in df_all_annual_2['Type'].values:
        var1 = list(set(df_all_annual_2[(df_all_annual_2['Type'] == x)]['Region'].values).symmetric_difference(set(seacen_country)))
        if len(var1) > 0:
            for y in var1:
                df_temp1 = pd.DataFrame({"Type": [x], "Region": [y]})
                df_all_annual_2 = pd.concat([df_all_annual_2, df_temp1], axis=0)
        else:
            continue
    df_all_annual_2.replace('South Korea', 'Korea', inplace=True)
    df_all_annual_2.sort_values(by=['Type', 'Region'], inplace=True)
    df_all_annual_2.reset_index(drop=True, inplace=True)
    df_all_annual_2.rename(columns={k:v for (k,v) in zip(df_all_annual_2.columns[4:], [x.year for x in df_all_annual_2.columns[4:]])}, inplace=True) ######### only for annnual data, dont use for quarterly data

    init_dict = df_init.to_dict()
    df_dict = dict()
    for x in init_dict['desc'].values():
        df_temp = df_all_annual_2[(df_all_annual_2[df_all_annual_2.columns[0]] == x)].reset_index(drop=True)
        df_dict[x] = df_temp

    calc_dict = {'Asia-Pacific': seacen_country,
                'Asian Advance Economies': ['Hong Kong SAR (China)', 'South Korea', 'Singapore', 'Taiwan'],
                'ASEAN5 Economies': ['Indonesia', 'Malaysia', 'Philippines', 'Thailand', 'Vietnam'],
                'Asian EDMEs': ['Brunei', 'Cambodia', 'Laos', 'Mongolia', 'Nepal', 'Sri Lanka', 'Papua New Guinea', 'Myanmar'],
                'China': ['China'],
                'India': ['India'],
                }

    ###################### this is for equity ######################

    asset_row = df_init.iloc[[0,4,5,6, 8, 9,10,11,12,13,14]]['desc'].to_list()
    asset_row2 = ['Direct Investment', 'Portfolio Equity', 'Portfolio Debt','Financial Derivatives', 'Other Equity','Currency & Deposits','Loans','Insurance, Pension & Standardized Guarantee Schemes'
                ,'Trade Credit & Advances','Other Accounts Receivable','Reserve Assets', ]
    sum_rows = df_init[(df_init['short_title'].isin(['OI Equity Assets','Insurance and Pension Assets','OI Others Assets',]))]['desc'].to_list()
    order = ['Direct Investment','Portfolio Equity','Portfolio Debt','Financial Derivatives','Currency & Deposits','Loans','Trade Credit & Advances','Other Investments','Reserve Assets',]
    df_by_region_main = pd.DataFrame(columns=df_all_annual_2.columns)

    for x,y in calc_dict.items():
        # print(x,y)
        df_by_region_temp = df_all_annual_2[(df_all_annual_2['Region'].isin(y)) & (df_all_annual_2['Type'].isin(asset_row))]
        df_by_region_temp.drop(columns=['Last Update Time', 'Region', 'Unit'], inplace=True)
        df_by_region_temp = df_by_region_temp.groupby(['Type']).sum()
        df_by_region_temp.loc[str(x) + ' Total'] = df_by_region_temp.sum()
        df_by_region_temp.loc['Other Investments'] = df_by_region_temp.loc[sum_rows].sum()
        df_by_region_temp.reset_index(inplace = True)
        df_by_region_temp = df_by_region_temp[(~df_by_region_temp['Type'].isin(sum_rows))] #####  the symbol of '~' use similar like notin, the oposite of isin. this is build in python operator
        df_by_region_temp['Type'] = df_by_region_temp['Type'].replace(asset_row, asset_row2)
        df_by_region_temp = df_by_region_temp.sort_values(by='Type', key=lambda x: x.map({k:i for i, k in enumerate(order)}))
        df_by_region_temp['Region'] = x

        df_by_region_main = pd.concat([df_by_region_main, df_by_region_temp], axis = 0)  


    df_by_region_main.reset_index(inplace = True, drop=True)
    df_by_region_main.insert(loc=0, column='Group', value='Assets')

    ###################### this is for equity ######################

    ###################### this is for debt ######################

    debt_row = df_init.iloc[[20,24,25,26,28,29,30,31,32,33,34]]['desc'].to_list()
    debt_row2 = ['Direct Investment', 'Portfolio Equity', 'Portfolio Debt','Financial Derivatives', 'Other Equity','Currency & Deposits','Loans','Insurance, Pension & Standardized Guarantee Schemes'
                ,'Trade Credit & Advances','Other Accounts Payable','SDR Liabilities', ]
    sum_rows2 = df_init[(df_init['short_title'].isin(['OI Equity Liabilities', 'Insurance and Pension Liabilities', 'OI Others Liabilities', 'SDR Liabilities']))]['desc'].to_list()
    order2 = ['Direct Investment','Portfolio Equity','Portfolio Debt','Financial Derivatives','Currency & Deposits','Loans','Trade Credit & Advances','Other Investments']
    df_by_region_main2 = pd.DataFrame(columns=df_all_annual_2.columns)

    for x,y in calc_dict.items():
        # print(x,y)
        df_by_region_temp = df_all_annual_2[(df_all_annual_2['Region'].isin(y)) & (df_all_annual_2['Type'].isin(debt_row))]
        df_by_region_temp.drop(columns=['Last Update Time', 'Region', 'Unit'], inplace=True)
        df_by_region_temp = df_by_region_temp.groupby(['Type']).sum()
        df_by_region_temp.loc[str(x) + ' Total'] = df_by_region_temp.sum()
        df_by_region_temp.loc['Other Investments'] = df_by_region_temp.loc[sum_rows2].sum()
        df_by_region_temp.reset_index(inplace = True)
        df_by_region_temp = df_by_region_temp[(~df_by_region_temp['Type'].isin(sum_rows2))] #####  the symbol of '~' use similar like notin, the oposite of isin. this is build in python operator
        df_by_region_temp['Type'] = df_by_region_temp['Type'].replace(debt_row, debt_row2)
        df_by_region_temp = df_by_region_temp.sort_values(by='Type', key=lambda x: x.map({k:i for i, k in enumerate(order2)}))
        df_by_region_temp['Region'] = x

        df_by_region_main2 = pd.concat([df_by_region_main2, df_by_region_temp], axis = 0)  

    df_by_region_main2.reset_index(inplace = True, drop=True)
    df_by_region_main2.insert(loc=0, column='Group', value='Liabilities')

    ###################### this is for debt ######################

    ###################### this is for caq data ######################

    caq_data = df_all_annual_2[(df_all_annual_2['Type'].isin(df_init[(df_init['group'] == 'Current Account')]['desc'].to_list()[1:]))]
    caq_data.drop(columns=['Last Update Time', 'Region', 'Unit'], inplace=True)
    caq_data_calc = caq_data.groupby(['Type']).sum()
    caq_data_calc.loc['Current Account Balance'] = caq_data_calc.sum()
    caq_data_calc['Region'] = 'Asia-Pacific'
    caq_data_calc.reset_index(inplace=True)
    rename_to = ['Goods Trade', 'Services Trade','Primary Income','Secondary Income',]
    rename_from = ['BoP: Current Account: Goods','BoP: Current Account: Services','BoP: Current Account: Primary Income','BoP: Current Account: Secondary Income',]
    caq_data_calc['Type'] = caq_data_calc['Type'].replace(rename_from, rename_to)
    caq_data_calc = caq_data_calc.sort_values(by='Type', key=lambda x: x.map({k:i for i, k in enumerate(rename_to)}))
    caq_data_calc.reset_index(drop=True, inplace=True)
    caq_data_calc.insert(loc=0, column='Group', value='Current Account')

    ###################### this is for caq data ######################

    df_main = pd.concat([df_by_region_main, df_by_region_main2, caq_data_calc], axis = 0) ## combine
    df_main.reset_index(drop=True, inplace=True)
    df_all_annual_2.replace({k:v for (k,v) in zip(df_init['desc'].to_list(), df_init['short_title'].to_list())}, inplace=True)

    return df_all_annual_2, df_main

def iip_quarterly(year_use:int = 2019) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    df_init = pd.read_excel(automation_path, sheet_name='IIP edit')

    df_all_annual = read_bop_iip('IMF IIP Quarterly')
    df_all_annual = df_all_annual[df_all_annual['Type'].isin(df_init['desc'])]


    df_all_annual_2 = df_all_annual[df_all_annual.columns[:4].to_list() + [dt for dt in df_all_annual.columns[4:].sort_values(ascending=True).to_list() if dt >= datetime(year_use, 1, 1)]]

    df_init_tw = pd.read_excel(automation_path, sheet_name='Taiwan IIP')
    df_tw = read_taiwan("IIP")
    df_tw.rename(columns={df_tw.columns[0]: 'Type'}, inplace=True)
    df_tw = df_tw[df_tw['Type'].isin(df_init_tw['desc'])]

    for x, y in enumerate(df_tw['Type'].to_list()):
        df_tw.iat[x,0] = list(df_init_tw[(df_init_tw['desc'] == y)]['desc2'])[0]

    selected_columns = [col for col in df_tw.columns[4:] if pd.to_datetime(col) >= datetime(year_use, 1, 1)]
    final_columns = list(df_tw.columns[:4]) + selected_columns
    df_tw = df_tw[final_columns]
    df_tw.replace('Taiwan (China)', 'Taiwan', inplace=True)
    df_tw2 = pd.DataFrame(columns=df_all_annual_2.columns)
    df_tw2 = pd.concat([df_tw2, df_tw], axis=0)

    for y in df_tw.columns[4:]:
        for x in df_tw2.columns[4:]:
            if y.year == x.year:
                df_tw2[x] = df_tw[y]  

    # concat to the main df
    df_all_annual_2 = pd.concat([df_all_annual_2, df_tw2], axis=0)

    seacen_country = ['Papua New Guinea', 'Vietnam', 'Nepal', 'India', 'Indonesia', 'Laos', 'Sri Lanka', 'Hong Kong SAR (China)', 'Philippines', 'Taiwan', 'Malaysia', 'Mongolia', 'China', 'Cambodia', 'Thailand', 'Singapore', 'South Korea', 'Brunei', 'Myanmar']

    for x in df_all_annual_2['Type'].values:
        var1 = list(set(df_all_annual_2[(df_all_annual_2['Type'] == x)]['Region'].values).symmetric_difference(set(seacen_country)))
        if len(var1) > 0:
            for y in var1:
                df_temp1 = pd.DataFrame({"Type": [x], "Region": [y]})
                df_all_annual_2 = pd.concat([df_all_annual_2, df_temp1], axis=0)
        else:
            continue
    
    df_all_annual_2.replace('South Korea', 'Korea', inplace=True)
    df_all_annual_2.sort_values(by=['Type', 'Region'], inplace=True)
    df_all_annual_2.reset_index(drop=True, inplace=True)

    init_dict = df_init.to_dict()
    df_dict = dict()
    for x in init_dict['desc'].values():
        df_temp = df_all_annual_2[(df_all_annual_2[df_all_annual_2.columns[0]] == x)].reset_index(drop=True)
        df_dict[x] = df_temp
    
    df_temp_col_1 = df_all_annual_2.copy()

    # 1H = June, 2H = December
    rename1 = [x for x in df_temp_col_1.columns[4:] if x.month in [6, 12]]
    rename2 = ['1H'+f'{x.year}' if x.month == 6 else '2H'+f'{x.year}' for x in rename1]
    dict_rename = {k:v for (k,v) in zip(rename1, rename2)}
    df_temp_col_1.rename(columns=dict_rename, inplace=True)
    df_temp_col_1 = df_temp_col_1[[*df_temp_col_1.columns[:4].to_list(), *rename2]]

    df_temp_col_1.drop(columns=['Last Update Time', 'Unit'], inplace=True)

    calc_dict = {'Asia-Pacific': seacen_country,
             'Asian Advance Economies': ['Hong Kong SAR (China)', 'South Korea', 'Singapore', 'Taiwan'],
             'ASEAN5 Economies': ['Indonesia', 'Malaysia', 'Philippines', 'Thailand', 'Vietnam'],
             'Asian EDMEs': ['Brunei', 'Cambodia', 'Laos', 'Mongolia', 'Nepal', 'Sri Lanka', 'Papua New Guinea', 'Myanmar'],
             'China': ['China'],
             'India': ['India'],
             }

    ###################### this is for equity ######################

    asset_row = df_init.iloc[[0,4,5,6, 8, 9,10,11,12,13,14]]['desc'].to_list()
    asset_row2 = ['Direct Investment', 'Portfolio Equity', 'Portfolio Debt','Financial Derivatives', 'Other Equity','Currency & Deposits','Loans','Insurance, Pension & Standardized Guarantee Schemes'
                ,'Trade Credit & Advances','Other Accounts Receivable','Reserve Assets', ]
    sum_rows = df_init[(df_init['short_title'].isin(['OI Equity Assets','Insurance and Pension Assets','OI Others Assets',]))]['desc'].to_list()
    order = ['Direct Investment','Portfolio Equity','Portfolio Debt','Financial Derivatives','Currency & Deposits','Loans','Trade Credit & Advances','Other Investments','Reserve Assets',]
    df_by_region_main = pd.DataFrame(columns=df_temp_col_1.columns)

    for x,y in calc_dict.items():
        # print(x,y)
        df_by_region_temp = df_temp_col_1[(df_temp_col_1['Region'].isin(y)) & (df_temp_col_1['Type'].isin(asset_row))]
        df_by_region_temp = df_by_region_temp.groupby(['Type']).sum()
        df_by_region_temp.loc[str(x) + ' Total'] = df_by_region_temp.sum()
        df_by_region_temp.loc['Other Investments'] = df_by_region_temp.loc[sum_rows].sum()
        df_by_region_temp.reset_index(inplace = True)
        df_by_region_temp = df_by_region_temp[(~df_by_region_temp['Type'].isin(sum_rows))] #####  the symbol of '~' use similar like notin, the oposite of isin. this is build in python operator
        df_by_region_temp['Type'] = df_by_region_temp['Type'].replace(asset_row, asset_row2)
        df_by_region_temp = df_by_region_temp.sort_values(by='Type', key=lambda x: x.map({k:i for i, k in enumerate(order)}))
        df_by_region_temp['Region'] = x

        df_by_region_main = pd.concat([df_by_region_main, df_by_region_temp], axis = 0)  


    df_by_region_main.reset_index(inplace = True, drop=True)
    df_by_region_main.insert(loc=0, column='Group', value='Assets')

    ###################### this is for equity ######################

    ###################### this is for debt ######################

    debt_row = df_init.iloc[[16,20,21,22,24,25,26,27,28,29,30]]['desc'].to_list()
    debt_row2 = ['Direct Investment', 'Portfolio Equity', 'Portfolio Debt','Financial Derivatives', 'Other Equity','Currency & Deposits','Loans','Insurance, Pension & Standardized Guarantee Schemes'
                ,'Trade Credit & Advances','Other Accounts Payable','SDR Liabilities', ]
    sum_rows2 = df_init[(df_init['short_title'].isin(['OI Equity Liabilities', 'Insurance and Pension Liabilities', 'OI Others Liabilities', 'SDR Liabilities']))]['desc'].to_list()
    order2 = ['Direct Investment','Portfolio Equity','Portfolio Debt','Financial Derivatives','Currency & Deposits','Loans','Trade Credit & Advances','Other Investments']
    df_by_region_main2 = pd.DataFrame(columns=df_temp_col_1.columns)

    for x,y in calc_dict.items():
        # print(x,y)
        df_by_region_temp = df_temp_col_1[(df_temp_col_1['Region'].isin(y)) & (df_temp_col_1['Type'].isin(debt_row))]
        df_by_region_temp = df_by_region_temp.groupby(['Type']).sum()
        df_by_region_temp.loc[str(x) + ' Total'] = df_by_region_temp.sum()
        df_by_region_temp.loc['Other Investments'] = df_by_region_temp.loc[sum_rows2].sum()
        df_by_region_temp.reset_index(inplace = True)
        df_by_region_temp = df_by_region_temp[(~df_by_region_temp['Type'].isin(sum_rows2))] #####  the symbol of '~' use similar like notin, the oposite of isin. this is build in python operator
        df_by_region_temp['Type'] = df_by_region_temp['Type'].replace(debt_row, debt_row2)
        df_by_region_temp = df_by_region_temp.sort_values(by='Type', key=lambda x: x.map({k:i for i, k in enumerate(order2)}))
        df_by_region_temp['Region'] = x

        df_by_region_main2 = pd.concat([df_by_region_main2, df_by_region_temp], axis = 0)  

    df_by_region_main2.reset_index(inplace = True, drop=True)
    df_by_region_main2.insert(loc=0, column='Group', value='Liabilities')

    ###################### this is for debt ######################

    ###################### this is for caq data ######################

    caq_data = df_temp_col_1[(df_temp_col_1['Type'].isin(df_init[(df_init['group'] == 'Current Account')]['desc'].to_list()[1:]))]
    caq_data_calc = caq_data.groupby(['Type']).sum()
    caq_data_calc.loc['Current Account Balance'] = caq_data_calc.sum()
    caq_data_calc['Region'] = 'Asia-Pacific'
    caq_data_calc.reset_index(inplace=True)
    rename_to = ['Goods Trade', 'Services Trade','Primary Income','Secondary Income',]
    rename_from = ['BoP: Current Account: Goods','BoP: Current Account: Services','BoP: Current Account: Primary Income','BoP: Current Account: Secondary Income',]
    caq_data_calc['Type'] = caq_data_calc['Type'].replace(rename_from, rename_to)
    caq_data_calc = caq_data_calc.sort_values(by='Type', key=lambda x: x.map({k:i for i, k in enumerate(rename_to)}))
    caq_data_calc.reset_index(drop=True, inplace=True)
    caq_data_calc.insert(loc=0, column='Group', value='Current Account')

    ###################### this is for caq data ######################

    df_main = pd.concat([df_by_region_main, df_by_region_main2, caq_data_calc], axis = 0) ## combine
    df_main.reset_index(drop=True, inplace=True)
    df_all_annual_2.replace({k:v for (k,v) in zip(df_init['desc'].to_list(), df_init['short_title'].to_list())}, inplace=True)
    df_temp_col_1.replace({k:v for (k,v) in zip(df_init['desc'].to_list(), df_init['short_title'].to_list())}, inplace=True)

    return df_temp_col_1, df_all_annual_2, df_main

def iip_annual(year_use:int = 2019) -> tuple[pd.DataFrame, pd.DataFrame]:

    df_init = pd.read_excel(automation_path, sheet_name='IIP edit')

    df_all_annual = read_bop_iip('IMF IIP Annual')
    df_all_annual = df_all_annual[df_all_annual['Type'].isin(df_init['desc'])]


    df_all_annual_2 = df_all_annual[df_all_annual.columns[:4].to_list() + [dt for dt in df_all_annual.columns[4:].sort_values(ascending=True).to_list() if dt >= datetime(year_use, 1, 1)]]

    df_init_tw = pd.read_excel(automation_path, sheet_name='Taiwan IIP')
    df_tw = read_taiwan("IIP")
    df_tw.rename(columns={df_tw.columns[0]: 'Type'}, inplace=True)
    df_tw = df_tw[df_tw['Type'].isin(df_init_tw['desc'])]

    for x, y in enumerate(df_tw['Type'].to_list()):
        df_tw.iat[x,0] = list(df_init_tw[(df_init_tw['desc'] == y)]['desc2'])[0]

    selected_columns = [col for col in df_tw.columns[4:] if pd.to_datetime(col) >= datetime(year_use, 1, 1)]
    final_columns = list(df_tw.columns[:4]) + selected_columns
    df_tw = df_tw[final_columns]
    df_tw.replace('Taiwan (China)', 'Taiwan', inplace=True)

    # concat to the main df
    df_all_annual_2 = pd.concat([df_all_annual_2, df_tw], axis=0)

    seacen_country = ['Papua New Guinea', 'Vietnam', 'Nepal', 'India', 'Indonesia', 'Laos', 'Sri Lanka', 'Hong Kong SAR (China)', 'Philippines', 'Taiwan', 'Malaysia', 'Mongolia', 'China', 'Cambodia', 'Thailand', 'Singapore', 'South Korea', 'Brunei', 'Myanmar']

    for x in df_all_annual_2['Type'].values:
        var1 = list(set(df_all_annual_2[(df_all_annual_2['Type'] == x)]['Region'].values).symmetric_difference(set(seacen_country)))
        if len(var1) > 0:
            for y in var1:
                df_temp1 = pd.DataFrame({"Type": [x], "Region": [y]})
                df_all_annual_2 = pd.concat([df_all_annual_2, df_temp1], axis=0)
        else:
            continue
    
    df_all_annual_2.replace('South Korea', 'Korea', inplace=True)
    df_all_annual_2.sort_values(by=['Type', 'Region'], inplace=True)
    df_all_annual_2.reset_index(drop=True, inplace=True)
    df_all_annual_2.rename(columns={k:v for (k,v) in zip(df_all_annual_2.columns[4:], [x.year for x in df_all_annual_2.columns[4:]])}, inplace=True) ######### only for annnual data, dont use for quarterly data

    init_dict = df_init.to_dict()
    df_dict = dict()
    for x in init_dict['desc'].values():
        df_temp = df_all_annual_2[(df_all_annual_2[df_all_annual_2.columns[0]] == x)].reset_index(drop=True)
        df_dict[x] = df_temp

    calc_dict = {'Asia-Pacific': seacen_country,
             'Asian Advance Economies': ['Hong Kong SAR (China)', 'South Korea', 'Singapore', 'Taiwan'],
             'ASEAN5 Economies': ['Indonesia', 'Malaysia', 'Philippines', 'Thailand', 'Vietnam'],
             'Asian EDMEs': ['Brunei', 'Cambodia', 'Laos', 'Mongolia', 'Nepal', 'Sri Lanka', 'Papua New Guinea', 'Myanmar'],
             'China': ['China'],
             'India': ['India'],
             }

    ###################### this is for equity ######################

    asset_row = df_init.iloc[[0,4,5,6, 8, 9,10,11,12,13,14]]['desc'].to_list()
    asset_row2 = ['Direct Investment', 'Portfolio Equity', 'Portfolio Debt','Financial Derivatives', 'Other Equity','Currency & Deposits','Loans','Insurance, Pension & Standardized Guarantee Schemes'
                ,'Trade Credit & Advances','Other Accounts Receivable','Reserve Assets', ]
    sum_rows = df_init[(df_init['short_title'].isin(['OI Equity Assets','Insurance and Pension Assets','OI Others Assets',]))]['desc'].to_list()
    order = ['Direct Investment','Portfolio Equity','Portfolio Debt','Financial Derivatives','Currency & Deposits','Loans','Trade Credit & Advances','Other Investments','Reserve Assets',]
    df_by_region_main = pd.DataFrame(columns=df_all_annual_2.columns)

    for x,y in calc_dict.items():
        # print(x,y)
        df_by_region_temp = df_all_annual_2[(df_all_annual_2['Region'].isin(y)) & (df_all_annual_2['Type'].isin(asset_row))]
        df_by_region_temp = df_by_region_temp.drop(columns=['Last Update Time', 'Unit']) ############## this is new 
        df_by_region_temp = df_by_region_temp.groupby(['Type']).sum()
        df_by_region_temp.loc[str(x) + ' Total'] = df_by_region_temp.sum()
        df_by_region_temp.loc['Other Investments'] = df_by_region_temp.loc[sum_rows].sum()
        df_by_region_temp.reset_index(inplace = True)
        df_by_region_temp = df_by_region_temp[(~df_by_region_temp['Type'].isin(sum_rows))] #####  the symbol of '~' use similar like notin, the oposite of isin. this is build in python operator
        df_by_region_temp['Type'] = df_by_region_temp['Type'].replace(asset_row, asset_row2)
        df_by_region_temp = df_by_region_temp.sort_values(by='Type', key=lambda x: x.map({k:i for i, k in enumerate(order)}))
        df_by_region_temp['Region'] = x

        df_by_region_main = pd.concat([df_by_region_main, df_by_region_temp], axis = 0)  


    df_by_region_main.reset_index(inplace = True, drop=True)
    df_by_region_main.insert(loc=0, column='Group', value='Assets')

    ###################### this is for debt ######################

    debt_row = df_init.iloc[[16,20,21,22,24,25,26,27,28,29,30]]['desc'].to_list()
    debt_row2 = ['Direct Investment', 'Portfolio Equity', 'Portfolio Debt','Financial Derivatives', 'Other Equity','Currency & Deposits','Loans','Insurance, Pension & Standardized Guarantee Schemes'
                ,'Trade Credit & Advances','Other Accounts Payable','SDR Liabilities', ]
    sum_rows2 = df_init[(df_init['short_title'].isin(['OI Equity Liabilities', 'Insurance and Pension Liabilities', 'OI Others Liabilities', 'SDR Liabilities']))]['desc'].to_list()
    order2 = ['Direct Investment','Portfolio Equity','Portfolio Debt','Financial Derivatives','Currency & Deposits','Loans','Trade Credit & Advances','Other Investments']
    df_by_region_main2 = pd.DataFrame(columns=df_all_annual_2.columns)

    for x,y in calc_dict.items():
        # print(x,y)
        df_by_region_temp = df_all_annual_2[(df_all_annual_2['Region'].isin(y)) & (df_all_annual_2['Type'].isin(debt_row))]
        df_by_region_temp = df_by_region_temp.drop(columns=['Last Update Time', 'Unit']) ############## this is new 
        df_by_region_temp = df_by_region_temp.groupby(['Type']).sum()
        df_by_region_temp.loc[str(x) + ' Total'] = df_by_region_temp.sum()
        df_by_region_temp.loc['Other Investments'] = df_by_region_temp.loc[sum_rows2].sum()
        df_by_region_temp.reset_index(inplace = True)
        df_by_region_temp = df_by_region_temp[(~df_by_region_temp['Type'].isin(sum_rows2))] #####  the symbol of '~' use similar like notin, the oposite of isin. this is build in python operator
        df_by_region_temp['Type'] = df_by_region_temp['Type'].replace(debt_row, debt_row2)
        df_by_region_temp = df_by_region_temp.sort_values(by='Type', key=lambda x: x.map({k:i for i, k in enumerate(order2)}))
        df_by_region_temp['Region'] = x

        df_by_region_main2 = pd.concat([df_by_region_main2, df_by_region_temp], axis = 0)  

    df_by_region_main2.reset_index(inplace = True, drop=True)
    df_by_region_main2.insert(loc=0, column='Group', value='Liabilities')

    ###################### this is for debt ######################

    df_main = pd.concat([df_by_region_main, df_by_region_main2], axis = 0) ## combine
    df_main.reset_index(drop=True, inplace=True)
    df_all_annual_2.replace({k:v for (k,v) in zip(df_init['desc'].to_list(), df_init['short_title'].to_list())}, inplace=True)

    return df_all_annual_2, df_main

def cb_bop_quarterly(year_use:int = 2019) -> pd.DataFrame:
    df_init = pd.read_excel(data_path, sheet_name='BOP')
    df_all_annual = read_bop_iip('CB BOP Quarterly')
    df_all_annual = df_all_annual[df_all_annual['Type'].isin(df_init['Item'])]

    df_all_annual = df_all_annual[df_all_annual['Region'].isin(df_init['Country'])]
    # df_all_annual

    df_all_annual_2 = df_all_annual[df_all_annual.columns[:4].to_list() + [dt for dt in df_all_annual.columns[4:].sort_values(ascending=True).to_list() if dt >= datetime(year_use, 1, 1)]]

    seacen_country = ['Papua New Guinea', 'Vietnam', 'Nepal', 'India', 'Indonesia', 'Laos', 'Sri Lanka', 'Hong Kong SAR (China)', 'Philippines', 'Taiwan', 'Malaysia', 'Mongolia', 'China', 'Cambodia', 'Thailand', 'Singapore', 'South Korea', 'Brunei', 'Myanmar']

    to_replace = {k:v for k,v in zip(df_init['Item'], df_init['short_title'])}

    df_all_annual_2.replace(to_replace, inplace=True)

    for x in df_all_annual_2['Type'].values:
        var1 = list(set(df_all_annual_2[(df_all_annual_2['Type'] == x)]['Region'].values).symmetric_difference(set(seacen_country)))
        if len(var1) > 0:
            for y in var1:
                df_temp1 = pd.DataFrame({"Type": [x], "Region": [y]})
                df_all_annual_2 = pd.concat([df_all_annual_2, df_temp1], axis=0)
        else:
            continue
    df_all_annual_2.replace('South Korea', 'Korea', inplace=True)

    df_all_annual_2 = df_all_annual_2[(df_all_annual_2['Region'].isin(df_init['Country'].unique()) & ~(df_all_annual_2['Unit'].isna()))] #remove this if needed
    df_np = df_all_annual_2[df_all_annual_2['Region'] == 'Nepal']
    df_all_annual_2 = df_all_annual_2[df_all_annual_2['Region'] != 'Nepal']
    dt_ch = [dt for dt in df_all_annual_2.columns[4:].sort_values(ascending=True).to_list() if dt.month in [4,7,10,1]]
    new_dt = [dt - relativedelta(months=1) for dt in dt_ch]
    df_np = df_np[df_np.columns[:4].to_list() + dt_ch]
    df_all_annual_2 = df_all_annual_2[df_all_annual_2.columns[:4].to_list() + new_dt]
    df_np.rename(columns={k:v for k,v in zip(dt_ch, new_dt)}, inplace=True)

    df_all_annual_2 = pd.concat([df_all_annual_2, df_np], axis=0)
    df_all_annual_2.sort_values(by=['Type', 'Region'], inplace=True)
    df_all_annual_2.reset_index(drop=True, inplace=True)

    return df_all_annual_2

def cb_iip_quarterly(year_use:int = 2019) -> pd.DataFrame:
    df_init = pd.read_excel(data_path, sheet_name='IIP')
    df_all_annual = read_bop_iip('CB IIP Quarterly')
    df_all_annual = df_all_annual[df_all_annual['Type'].isin(df_init['Item'])]

    df_all_annual = df_all_annual[df_all_annual['Region'].isin(df_init['Country'])]


    df_all_annual_2 = df_all_annual[df_all_annual.columns[:4].to_list() + [dt for dt in df_all_annual.columns[4:].sort_values(ascending=True).to_list() if dt >= datetime(year_use, 1, 1)]]

    seacen_country = ['Papua New Guinea', 'Vietnam', 'Nepal', 'India', 'Indonesia', 'Laos', 'Sri Lanka', 'Hong Kong SAR (China)', 'Philippines', 'Taiwan', 'Malaysia', 'Mongolia', 'China', 'Cambodia', 'Thailand', 'Singapore', 'South Korea', 'Brunei', 'Myanmar']

    to_replace = {k:v for k,v in zip(df_init['Item'], df_init['short_title'])}

    df_all_annual_2.replace(to_replace, inplace=True)

    for x in df_all_annual_2['Type'].values:
        var1 = list(set(df_all_annual_2[(df_all_annual_2['Type'] == x)]['Region'].values).symmetric_difference(set(seacen_country)))
        if len(var1) > 0:
            for y in var1:
                df_temp1 = pd.DataFrame({"Type": [x], "Region": [y]})
                df_all_annual_2 = pd.concat([df_all_annual_2, df_temp1], axis=0)
        else:
            continue
    df_all_annual_2.replace('South Korea', 'Korea', inplace=True)

    df_all_annual_2 = df_all_annual_2[(df_all_annual_2['Region'].isin(df_init['Country'].unique()) & ~(df_all_annual_2['Unit'].isna()))] #remove this if needed

    df_all_annual_2.sort_values(by=['Type', 'Region'], inplace=True)
    df_all_annual_2.reset_index(drop=True, inplace=True)
    
    return df_all_annual_2

def to_combine_bopq(year_use:int = 2019):
    x1, y1, z1 = bop_quarterly()
    y2 = cb_bop_quarterly()
    df_temp2 = pd.DataFrame(columns=y1.columns)
    temp1 = pd.DataFrame(columns=y1.columns)
    index_list = []
    date_use = {'Laos': [datetime(2025, 9, 1)],
                'Malaysia': [datetime(2025, 3, 1), datetime(2025, 6, 1), datetime(2025, 9, 1)],
                'Mongolia': [datetime(2025, 6, 1), datetime(2025, 9, 1)],
                'Nepal': [datetime(2025, 9, 1)],
                'Sri Lanka': [datetime(2025, 6, 1), datetime(2025, 9, 1)],
                }

    for x,y in zip(y2['Type'].values, y2['Region'].values):
        if y not in date_use.keys():
            continue
        temp1 = y1[(y1['Type'] == x) & (y1['Region'] == y)]
        index_list.append(temp1.index[0])
        temp1.sort_values(by=['Type', 'Region'], inplace=True)
        temp1[date_use[y]] = y2[(y2['Type'] == x) & (y2['Region'] == y)][date_use[y]].values
        df_temp2 = pd.concat([df_temp2, temp1], axis=0)
        
    y1.drop(index_list, inplace=True, axis=0) ########################## this line is new
    df_temp2 = pd.concat([y1, df_temp2], axis=0) ########################## this line is new
    df_temp2.sort_values(by=['Type', 'Region'], inplace=True)
    df_temp2.reset_index(drop=True, inplace=True)
    return df_temp2

def to_combine_iipq(year_use:int = 2019):
    x1, y1, z1 = iip_quarterly()
    y2 = cb_iip_quarterly()
    df_temp2 = pd.DataFrame(columns=y1.columns)
    temp1 = pd.DataFrame(columns=y1.columns)
    index_list = []
    date_use = {
                'Malaysia': [datetime(2025, 3, 1), datetime(2025, 6, 1), datetime(2025, 9, 1)],                
                'Sri Lanka': [datetime(2025, 6, 1), datetime(2025, 9, 1), ],
                }

    for x,y in zip(y2['Type'].values, y2['Region'].values):
        if y not in date_use.keys():
            continue
        temp1 = y1[(y1['Type'] == x) & (y1['Region'] == y)]
        index_list.append(temp1.index[0])
        temp1.sort_values(by=['Type', 'Region'], inplace=True)
        temp1[date_use[y]] = y2[(y2['Type'] == x) & (y2['Region'] == y)][date_use[y]].values
        df_temp2 = pd.concat([df_temp2, temp1], axis=0)

    y1.drop(index_list, inplace=True, axis=0) ########################## this line is new
    df_temp2 = pd.concat([y1, df_temp2], axis=0) ########################## this line is new
    df_temp2.sort_values(by=['Type', 'Region'], inplace=True)
    df_temp2.reset_index(drop=True, inplace=True)
    return df_temp2

def main(year_use:int = 2019) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    check_paths()
    df1_1, df1_2, df1_3 = iip_quarterly(year_use)
    df2_1, df2_2 = iip_annual(year_use)
    df3_1, df3_2, df3_3 = bop_quarterly(year_use)
    df4_1, df4_2 = bop_annual(year_use)

    df1_1.insert(loc=0, column='Group2', value='IIP Quarterly')
    df1_2.insert(loc=0, column='Group2', value='IIP Quarterly')
    df1_3.insert(loc=0, column='Group2', value='IIP Quarterly')
    df3_1.insert(loc=0, column='Group2', value='BoP Quarterly')
    df3_2.insert(loc=0, column='Group2', value='BoP Quarterly')
    df3_3.insert(loc=0, column='Group2', value='BoP Quarterly')

    df2_1.insert(loc=0, column='Group2', value='IIP Annual')
    df2_2.insert(loc=0, column='Group2', value='IIP Annual')
    df4_1.insert(loc=0, column='Group2', value='BoP Annual')
    df4_2.insert(loc=0, column='Group2', value='BoP Annual')

    quarterly_type_half = pd.concat([df1_1, df3_1], axis=0)
    quarterly_type_quarter = pd.concat([df1_2, df3_2], axis=0)
    quarterly_region_half = pd.concat([df1_3, df3_3], axis=0)

    annual_type_full = pd.concat([df2_1, df4_1], axis=0)
    annual_region_full = pd.concat([df2_2, df4_2], axis=0)
    annual_region_full.drop(columns=['Last Update Time', 'Unit'], inplace=True)

    return quarterly_type_half, quarterly_type_quarter, quarterly_region_half, annual_type_full, annual_region_full