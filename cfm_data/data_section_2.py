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

# The economies in each regional group, spelled as in the tables at the point
# the groups are summed (after "South Korea" is renamed "Korea"). Brunei,
# Myanmar and Papua New Guinea stay in the by-economy tables but are left out
# of every group, including the Asia-Pacific total and the current account, as
# in the SEACEN Capital Flows Monitor.
ASIAN_ADVANCED = ['Hong Kong SAR (China)', 'Korea', 'Singapore', 'Taiwan']
ASEAN5 = ['Indonesia', 'Malaysia', 'Philippines', 'Thailand', 'Vietnam']
ASIAN_EDMES = ['Cambodia', 'Laos', 'Mongolia', 'Nepal', 'Sri Lanka']
REGION_GROUPS = {
    'Asia-Pacific': ASIAN_ADVANCED + ASEAN5 + ASIAN_EDMES + ['China', 'India'],
    'Asian Advance Economies': ASIAN_ADVANCED,
    'ASEAN5 Economies': ASEAN5,
    'Asian EDMEs': ASIAN_EDMES,
    'China': ['China'],
    'India': ['India'],
}

warnings.filterwarnings("ignore")


def half_year_sums(df: pd.DataFrame, label_cols: int = 4) -> pd.DataFrame:
    """Quarterly flows (US$ million, one column per quarter) -> half-yearly
    sums in US$ billion ("1H2026" = Q1 + Q2, "2H2026" = Q3 + Q4), with the
    label columns first. Quarters are paired by their dates, not their
    positions. A half-year is built only once both of its quarters are in the
    table, and an economy missing either quarter gets a blank for that
    half-year rather than a one-quarter sum."""
    out = df[df.columns[:label_cols]].reset_index(drop=True)
    values = df[df.columns[label_cols:]].apply(pd.to_numeric, errors="coerce").reset_index(drop=True)
    by_quarter = {(d.year, (d.month - 1) // 3 + 1): d for d in values.columns}
    for year in sorted({year for year, _ in by_quarter}):
        for half, (first, second) in ((1, (1, 2)), (2, (3, 4))):
            if (year, first) in by_quarter and (year, second) in by_quarter:
                total = values[by_quarter[(year, first)]] + values[by_quarter[(year, second)]]
                out[f"{half}H{year}"] = total / 1000
    return out


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


def _quarter_label(date) -> str:
    """2026-06-01 -> "2Q2026"."""
    return f"{(date.month - 1) // 3 + 1}Q{date.year}"


IIP_MISSING_COLUMNS = ["Economy", "Item", "Last available", "Value used", "Quarters filled", "Count"]


def fill_to_latest_quarter(df: pd.DataFrame, report: list | None = None, label_cols: int = 4) -> pd.DataFrame:
    """Carries each row's last value forward to the latest quarter any
    economy has data for, so an economy that hasn't reported the newest
    quarter yet (e.g. Cambodia's 2Q2026) shows its latest position instead
    of a gap. Gaps before a row's last value, and rows with no data at all,
    are left alone. Each filled row is added to `report` (columns
    IIP_MISSING_COLUMNS), and each economy filled is printed."""
    periods = df.columns[label_cols:]
    values = df[periods].apply(pd.to_numeric, errors="coerce")
    with_data = values.notna().any()
    latest = with_data[with_data].index.max()
    filled = {}
    for i, row in values.iterrows():
        last = row.last_valid_index()
        if last is None or last >= latest:
            continue
        gap = [p for p in periods if last < p <= latest]
        df.loc[i, gap] = row[last]
        filled.setdefault((df.at[i, "Region"], last), []).append(df.at[i, "Type"])
        if report is not None:
            report.append({
                "Economy": df.at[i, "Region"],
                "Item": df.at[i, "Type"],
                "Last available": _quarter_label(last),
                "Value used": row[last],
                "Quarters filled": ", ".join(_quarter_label(p) for p in gap),
                "Count": len(gap),
            })
    for (region, last), types in sorted(filled.items(), key=lambda kv: (str(kv[0][0]), kv[0][1])):
        print(f"IIP quarterly: {region} carried forward from {_quarter_label(last)} "
              f"to {_quarter_label(latest)} ({len(types)} items)")
    return df


CB_FILLED_COLUMNS = ["Dataset", "Economy", "Item", "Quarters", "Count", "Sign"]
# How the central-bank sign convention is checked against the IMF's, per
# economy and side (assets, liabilities, current account): of the quarters
# where both have a value (above US$1 million either way), at least this share
# must agree in sign (values used as they are) or disagree (values reversed),
# and there must be at least CB_MIN_OVERLAP such quarters.
CB_SIGN_AGREEMENT = 0.9
CB_MIN_OVERLAP = 8


def _side(item: str) -> str:
    if item.endswith("Assets"):
        return "Assets"
    if item.endswith("Liabilities"):
        return "Liabilities"
    return "Current account"


def fill_from_central_bank(df: pd.DataFrame, cb: pd.DataFrame, dataset: str, short_titles: dict,
                           report: list | None = None, label_cols: int = 4) -> pd.DataFrame:
    """Fills blank cells of the IMF quarterly table `df` from the central-bank
    table `cb` (same economy, item and quarter). Only blanks are filled: IMF
    values are never replaced, and quarters the IMF table doesn't have aren't
    added. `df` names its items by their long description and `cb` by their
    short title; `short_titles` maps one to the other.

    Some central banks record BoP assets with the opposite sign to the IMF
    (e.g. Cambodia, Malaysia and Vietnam). So for each economy and side
    (assets, liabilities, current account), the quarters where both have a
    value decide the sign: values are used as they are if at least
    CB_SIGN_AGREEMENT of them agree, reversed if as many disagree, and that
    side isn't filled otherwise (or with fewer than CB_MIN_OVERLAP such
    quarters). Every item with blanks the central bank has values for is added
    to `report` (columns CB_FILLED_COLUMNS): its quarters, and in "Sign"
    whether they were filled as reported, reversed, or not filled."""
    periods = [p for p in df.columns[label_cols:] if p in set(cb.columns[label_cols:])]
    if not periods:
        return df
    items = df["Type"].map(lambda t: short_titles.get(t, t))
    keys = list(zip(df["Region"], items))
    imf = df[periods].apply(pd.to_numeric, errors="coerce")
    cbv = cb.set_index(["Region", "Type"])[periods].apply(pd.to_numeric, errors="coerce")
    cbv = cbv[~cbv.index.duplicated()].reindex(keys)
    cbv.index = df.index

    side = pd.Series([f"{region}|{_side(str(item))}" for region, item in keys], index=df.index)
    both = imf.notna() & cbv.notna() & (imf.abs() > 1) & (cbv.abs() > 1)
    agree = both & ((imf > 0) == (cbv > 0))
    n_both = both.sum(axis=1).groupby(side).sum()
    n_agree = agree.sum(axis=1).groupby(side).sum()
    sign = {}
    for key in n_both.index:
        total, same = n_both[key], n_agree[key]
        if total >= CB_MIN_OVERLAP and same >= CB_SIGN_AGREEMENT * total:
            sign[key] = 1
        elif total >= CB_MIN_OVERLAP and total - same >= CB_SIGN_AGREEMENT * total:
            sign[key] = -1

    filled = {}
    for i in df.index:
        gap = imf.loc[i].isna() & cbv.loc[i].notna()
        quarters = [p for p in periods if gap[p]]
        if not quarters:
            continue
        factor = sign.get(side[i])
        if factor is not None:
            df.loc[i, quarters] = cbv.loc[i, quarters] * factor
            filled[df.at[i, "Region"]] = filled.get(df.at[i, "Region"], 0) + 1
        if report is not None:
            report.append({
                "Dataset": dataset,
                "Economy": df.at[i, "Region"],
                "Item": items[i],
                "Quarters": ", ".join(_quarter_label(p) for p in quarters),
                "Count": len(quarters),
                "Sign": {1: "as reported", -1: "reversed"}.get(
                    factor, "not filled: its sign convention can't be matched to the IMF's"),
            })
    for region, count in sorted(filled.items()):
        print(f"{dataset} quarterly: {region} filled from its central bank ({count} items)")
    return df


def bop_quarterly(year_use:int = 2019, report: list | None = None) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """The quarterly BoP tables: (by half-year, by quarter, by regional group).
    Blanks in the IMF data are filled from the central-bank workbook first;
    what was filled is added to `report` (see fill_from_central_bank)."""
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

    # blanks in the IMF data filled from the central-bank workbook, before any sums
    df_all_annual_2 = fill_from_central_bank(df_all_annual_2, cb_bop_quarterly(year_use), "BoP",
                                             dict(zip(df_init['desc'], df_init['short_title'])), report)

    init_dict = df_init.to_dict()
    df_dict = dict()
    for x in init_dict['desc'].values():
        df_temp = df_all_annual_2[(df_all_annual_2[df_all_annual_2.columns[0]] == x)].reset_index(drop=True)
        df_dict[x] = df_temp

    # half-yearly sums (1H = Q1 + Q2, 2H = Q3 + Q4), in US$ billion
    df_temp_col_1 = half_year_sums(df_all_annual_2)

    df_temp_col_1.drop(columns=['Last Update Time', 'Unit'], inplace=True)

    calc_dict = REGION_GROUPS

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

    caq_data = df_temp_col_1[(df_temp_col_1['Type'].isin(df_init[(df_init['group'] == 'Current Account')]['desc'].to_list()[1:]))
                             & df_temp_col_1['Region'].isin(REGION_GROUPS['Asia-Pacific'])]
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

    calc_dict = REGION_GROUPS

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

    caq_data = df_all_annual_2[(df_all_annual_2['Type'].isin(df_init[(df_init['group'] == 'Current Account')]['desc'].to_list()[1:]))
                               & df_all_annual_2['Region'].isin(REGION_GROUPS['Asia-Pacific'])]
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

def iip_quarterly(year_use:int = 2019, report: list | None = None,
                  cb_report: list | None = None) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """The quarterly IIP tables: (by half-year, by quarter, by regional group).
    Blanks in the IMF data are first filled from the central-bank workbook
    (added to `cb_report`, see fill_from_central_bank); then rows still short
    of the latest quarter are carried forward (added to `report`, see
    fill_to_latest_quarter), with their items named as in the tables."""
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

    # blanks in the IMF data filled from the central-bank workbook, before the carry-forward
    df_all_annual_2 = fill_from_central_bank(df_all_annual_2, cb_iip_quarterly(year_use), "IIP",
                                             dict(zip(df_init['desc'], df_init['short_title'])), cb_report)

    # positions are stocks, so an economy missing the newest quarter(s) keeps its latest position
    filled = []
    df_all_annual_2 = fill_to_latest_quarter(df_all_annual_2, filled)
    if report is not None:
        short_titles = dict(zip(df_init['desc'], df_init['short_title']))
        report.extend({**row, "Item": short_titles.get(row["Item"], row["Item"])} for row in filled)

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

    calc_dict = REGION_GROUPS

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

    caq_data = df_temp_col_1[(df_temp_col_1['Type'].isin(df_init[(df_init['group'] == 'Current Account')]['desc'].to_list()[1:]))
                             & df_temp_col_1['Region'].isin(REGION_GROUPS['Asia-Pacific'])]
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

    calc_dict = REGION_GROUPS

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

def main(year_use:int = 2019) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """The Section 2 tables: quarterly_type_half, quarterly_type_quarter,
    quarterly_region_half, annual_type_full, annual_region_full, iip_missing
    (the quarterly IIP rows filled up to the latest quarter) and cb_filled
    (the quarterly BoP and IIP blanks filled from the central-bank workbooks)."""
    check_paths()
    iip_missing = []
    cb_filled = []
    df1_1, df1_2, df1_3 = iip_quarterly(year_use, iip_missing, cb_filled)
    df2_1, df2_2 = iip_annual(year_use)
    df3_1, df3_2, df3_3 = bop_quarterly(year_use, cb_filled)
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

    # one block per economy, its items in the tables' order
    iip_missing = pd.DataFrame(iip_missing, columns=IIP_MISSING_COLUMNS).sort_values("Economy", kind="stable")
    iip_missing.reset_index(drop=True, inplace=True)
    cb_filled = pd.DataFrame(cb_filled, columns=CB_FILLED_COLUMNS).sort_values(["Dataset", "Economy"], kind="stable")
    cb_filled.reset_index(drop=True, inplace=True)

    return (quarterly_type_half, quarterly_type_quarter, quarterly_region_half, annual_type_full, annual_region_full,
            iip_missing, cb_filled)