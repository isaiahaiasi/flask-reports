import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# * Helpers
def get_cell(col, row):
    alpha = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    col_alpha = alpha[col % len(alpha)]
    # TODO: handle columns > Z
    return f"{col_alpha}{row}"

def fmt_time(t_raw):
    try:
        numeric_str = re.sub(r'h|m', '', t_raw)     # '3h 15m' -> '3 15'
        str_h, str_m = numeric_str.split(" ")       # '3 15' -> ('3', '15')
        h, m = float(str_h), float(str_m)           # ('3', '15') -> (3.0, 15.0)
        return ((h * 60) + m)/60                    # (3.0, 15.0) -> 3.25
    except:
        return None


# * CSV -> DATAFRAME fns
def get_grouped_dfs(input_file):
    df_raw = pd.read_csv(input_file)

    # strip unnecessary columns
    df_clean = df_raw.loc[:, 'Date':'Full Name']
    df_clean["Hours Worked"] = df_raw['Hours Worked'].map(fmt_time)

    # split by employee and format xlsx fragment for each person
    dfdict_group = {x: y.drop("Full Name", axis=1)
                    for x, y in df_clean.groupby('Full Name')}

    return dfdict_group


# * dict<DATAFRAME> -> XLSX Fns
# write individual timesheet:
# - SORT:
#   - Top-level sort should be by Break Type (so they can be more easily ignored)
# - COLS:
#   - "Hours worked": translation of og Hours Worked/Worked Hours (?) to number format
#   - "REG": Hours worked, minus UNPAID, minus OT
#   - "OT": Empty
#   - "SICK": Empty
#   - "PTO": Empty
#   - "HOLIDAY": Empty
# - SUMS:
#   - UNPAID/"Break Type": need SUM to subtract from Hours Worked for REG hours
#   - All added cols need sums: Hours Worked, REG, OT, SICK, PTO, HOLIDAY

def write_individual_timesheet(workbook, name, raw_df):
    worksheet = workbook.create_sheet(name)
    worksheet[get_cell(0, 1)] = name

    # write contents of dataframe, including headers
    for r in dataframe_to_rows(raw_df, index=False, header=True):
        worksheet.append(r)

    # add column headings for SICK, PTO, HOLIDAY
    worksheet[get_cell(3, 2)] = "SICK"
    worksheet[get_cell(4, 2)] = "PTO"
    worksheet[get_cell(5, 2)] = "HOLIDAY"

    # write formulae for totals under df
    r = len(raw_df.index) + 3
    worksheet[get_cell(0, r)] = 'Totals:'
    totals_count = 4
    for x in range(totals_count):
        c = x + 2
        sum_range_start = get_cell(c, 3)
        sum_range_end = get_cell(c, r - 1)
        worksheet[get_cell(c, r)] = f"=SUM({sum_range_start}:{sum_range_end})"

    # write "grand total" underneath other totals
    r = r + 1
    worksheet[get_cell(0, r)] = "Grand total:"
    totals_start = get_cell(2, r - 1)
    totals_end = get_cell(2 + totals_count - 1, r - 1)
    worksheet[get_cell(2, r)] = f"=SUM({totals_start}:{totals_end})"


def get_xlsx_from_df_group(df_group):
    wb = Workbook()
    for [name, df] in df_group.items():
        write_individual_timesheet(wb, name, df)

    # delete default sheet
    wb.remove(wb["Sheet"])

    return wb


# takes raw csv input_file and returns formatted xlsx file
def get_formatted_timesheet(input_file):
    df_group = get_grouped_dfs(input_file)
    return get_xlsx_from_df_group(df_group)
