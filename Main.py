import gspread
import pandas as pd
import numpy as np
import datetime
from us_state_abbrev import abbrev_us_state
import re
import sys
from xlsxwriter.utility import xl_cell_to_rowcol
from gspread.urls import SPREADSHEETS_API_V4_BASE_URL
from tqdm import tqdm

gc = gspread.oauth()

print('Reading Data...')
team_entry_sheet = gc.open(
    "Mountains-Midwest COVID-19 by County (TEAM ENTRY) 9-15-2020 to present").sheet1
nyt_comparison_sheet = gc.open(
    "QA TEAM ONLY - COVID-19 by County Comparison Data (NYT)").sheet1

team_entry = pd.DataFrame(team_entry_sheet.get_all_records(head=2))
team_entry = team_entry.replace(r'^\s*$', np.nan, regex=True)
nyt_comparison = pd.DataFrame(nyt_comparison_sheet.get_all_records())
nyt_comparison = nyt_comparison.replace(r'^\s*$', np.nan, regex=True)

JHU_USAF_Cases_comparison_sheets = gc.open(
    "COVID-19 by County Comparison Data (JHU/USAF) ")

USAF_Cases_comparison_sheet = JHU_USAF_Cases_comparison_sheets.sheet1
USAF_Deaths_comparison_sheet = JHU_USAF_Cases_comparison_sheets.get_worksheet(
    1)
JHU_comparison_sheet = JHU_USAF_Cases_comparison_sheets.get_worksheet(2)

USAF_Cases_comparison = pd.DataFrame(
    USAF_Cases_comparison_sheet.get_all_records())
USAF_Deaths_comparison = pd.DataFrame(
    USAF_Deaths_comparison_sheet.get_all_records())
JHU_comparison = pd.DataFrame(JHU_comparison_sheet.get_all_records())

USAF_Cases_comparison = USAF_Cases_comparison.replace(
    r'^\s*$', np.nan, regex=True)
USAF_Deaths_comparison = USAF_Deaths_comparison.replace(
    r'^\s*$', np.nan, regex=True)
JHU_comparison = JHU_comparison.replace(r'^\s*$', np.nan, regex=True)

print('Done')


print('Pouring functions...')
# given a data frame and state, find out the latest updated time, formated in US time format


def lastTimeUpdated(state):
    print(f'For {state}, last updated on:')

    # team entry
    yourstateDF = team_entry[team_entry.state == state]
    nototalDF = yourstateDF.filter(regex=f"\d$", axis=1).iloc[1:]
    first_col_empty = pd.isna(nototalDF).any().idxmax()
#     print(first_col_empty) # show which column has empty values
    nextdate_str = re.compile("\d.*").findall(first_col_empty)[0]
    nextdate = datetime.datetime.strptime(nextdate_str, '%m%d%y')
    te_date = nextdate - datetime.timedelta(days=1)
    print(f"Team entry: {te_date.strftime('%m/%d/%y')}")

    # nyt comparison dataset
    nyt_date_str = nyt_comparison.date.iloc[-1]
    nyt_date = datetime.datetime.strptime(nyt_date_str, '%Y-%m-%d')
    print(f"NYT comparison dataset: {nyt_date.strftime('%m/%d/%y')}")

    # USAF
    usaf_date_str = USAF_Cases_comparison.columns[-1]
    usaf_date = datetime.datetime.strptime(usaf_date_str, '%m/%d/%Y')
    print(f"USAF comparison dataset: {usaf_date.strftime('%m/%d/%y')}")

    # JHU
    jhu_date_str = JHU_comparison.columns[-1][1:-2]
    jhu_date = datetime.datetime.strptime(jhu_date_str, '%Y%m%d')
    print(f"JHU comparison dataset: {jhu_date.strftime('%m/%d/%y')}")

    earliest_date = min([te_date, nyt_date, usaf_date, jhu_date])
    print(f"Date you can work on: {earliest_date.strftime('%m/%d/%y')}")

    return earliest_date

# given a state and a date (module datetime), output part of data frame you need to work on.


def partTeamEntry(state, date):
    yourstateDF = team_entry[team_entry.state == state]
    infoDF = yourstateDF.iloc[:, 0:5]
    covidDF = yourstateDF.filter(regex=f"{date.strftime('%m%d%y')}$", axis=1)

    return pd.concat([infoDF, covidDF], axis=1).iloc[1:, 5:]
#     return infoDF


def partNYT(state, date):
    date_str = date.strftime('%Y-%m-%d')
    return nyt_comparison[(nyt_comparison.date == date_str) & (nyt_comparison.state == abbrev_us_state[state])]


def partUSAFCases(state, date):
    date_str = date.strftime('%-m/%-d/%Y')
    return USAF_Cases_comparison[USAF_Cases_comparison.State == state].filter(items=[date_str])


def partUSAFDeaths(state, date):
    date_str = date.strftime('%-m/%-d/%Y')
    return USAF_Deaths_comparison[USAF_Deaths_comparison.State == state].filter(items=[date_str])


def partJHU(state, date):
    date_str = [f"x{date.strftime('%Y%m%d')}_c",
                f"x{date.strftime('%Y%m%d')}_m"]
    return JHU_comparison[JHU_comparison.state_ab == state].filter(items=date_str)


def letter_to_num(letter):
    return [ord(char) - 97 for char in letter.lower()][0]


def inputData(comp_xlsx, state, date):
    from openpyxl import load_workbook

    book = load_workbook(comp_xlsx)
    writer = pd.ExcelWriter(comp_xlsx, engine='openpyxl')
    writer.book = book

    # ExcelWriter for some reason uses writer.sheets to access the sheet.
    # If you leave it empty it will not know that sheet Main is already there
    # and will create a new sheet.

    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

    partTeamEntry(state, date).to_excel(writer,
                                        startcol=letter_to_num('C'),
                                        startrow=3,
                                        header=None, index=False,
                                        sheet_name=list(writer.sheets.keys())[1])
    partJHU(state, date).to_excel(writer,
                                  startcol=letter_to_num('I'),
                                  startrow=3,
                                  header=None, index=False,
                                  sheet_name=list(writer.sheets.keys())[1])
    partNYT(state, date).to_excel(writer,
                                  startcol=letter_to_num('L'),
                                  startrow=3,
                                  header=None, index=False,
                                  sheet_name=list(writer.sheets.keys())[1])
    partUSAFCases(state, date).to_excel(writer,
                                        startcol=letter_to_num('V'),
                                        startrow=3,
                                        header=None, index=False,
                                        sheet_name=list(writer.sheets.keys())[1])
    partUSAFDeaths(state, date).to_excel(writer,
                                         startcol=26,
                                         startrow=3,
                                         header=None, index=False,
                                         sheet_name=list(writer.sheets.keys())[1])

    writer.save()


def insert_note(worksheet, label, note):
    """
    Insert note ito the google worksheet for a certain cell.

    Compatible with gspread.
    """
    spreadsheet_id = worksheet.spreadsheet.id
    worksheet_id = worksheet.id

    row, col = tuple(label)      # [0, 0] is A1

    url = f"{SPREADSHEETS_API_V4_BASE_URL}/{spreadsheet_id}:batchUpdate"
    payload = {
        "requests": [
            {
                "updateCells": {
                    "range": {
                        "sheetId": worksheet_id,
                        "startRowIndex": row,
                        "endRowIndex": row + 1,
                        "startColumnIndex": col,
                        "endColumnIndex": col + 1
                    },
                    "rows": [
                        {
                            "values": [
                                {
                                    "note": note
                                }
                            ]
                        }
                    ],
                    "fields": "note"
                }
            }
        ]
    }
    worksheet.spreadsheet.client.request("post", url, json=payload)


print('Done')

if __name__ == '__main__':
    today = datetime.date.today()
    print("Today's date is", today)

    state = input('Enter a state you working on: ').upper()

    lastTimeUpdated(state)

    status = True
    while status:
        # custom date
        # last: 2020, 10 ,23
        print('Please input the date you want to work on: ')
        YY = int(input('Year: '))
        MM = int(input('Month: '))
        DD = int(input('Date: '))
        date = datetime.datetime(YY, MM, DD)

        pre = './comparison sheets/'
        if state == 'MN':
            filename = "Minnesota - Mountains - County Comparison (with probable)"
            inputData(f'{pre+filename}.xlsx', state, date)
        if state == 'ND':
            filename = "North Dakota - Mountains - County Comparison (with probable)"
            inputData(f'{pre+filename}.xlsx', state, date)
        if state == 'TX':
            filename = "Texas - Mountains - County Comparison (with probable)"
            inputData(f'{pre+filename}.xlsx', state, date)

        print('Please manually save your .xlsx file to .csv.')
        input("Press Enter to continue...")

        print('Processing...')
        read_comp_again = pd.read_csv(f'{pre+filename}.csv', header=1)
        read_comp_again = read_comp_again.dropna(how='all')
        part_read_again = read_comp_again.loc[:, [
            'County', 'State', 'Unnamed: 21', 'Unnamed: 22']]

        # change the order as team entry
        part_read_again_wo_total = part_read_again.loc[1:, :]
        unknown_row = part_read_again_wo_total.iloc[0, :]

        part_read_again_wo_total = part_read_again_wo_total.shift(-1)
        part_read_again_wo_total.iloc[-1] = unknown_row.squeeze()

        part_reference = pd.concat(
            [pd.DataFrame(part_read_again.loc[0, :]).transpose(), part_read_again_wo_total])
        part_reference.columns = ['County', 'State',
                                  'Case Comments', 'Death Comments']

        case_comments = part_reference.loc[:, 'Case Comments'].to_list()
        death_comments = part_reference.loc[:, 'Death Comments'].to_list()

        start_col_idx = [i for i, col in enumerate(
            team_entry.columns) if col.endswith(date.strftime('%m%d%y'))][::2]
        # start and end only for tx!
        start_row_idx = 469
        end_row_idx = 724
        print('Double checking if the column number is correct...')

        CELL_INDEX = input('Please input the cell index (e.g. FJ470): ')
        if start_col_idx[0] != xl_cell_to_rowcol(CELL_INDEX)[1]:
            print('Not syncing indices! Aborting the program.')
            sys.exit()
        else:
            print('Success syncing indices.')

        print('Writing cases columns...')
        # Case
        for i, ridx in enumerate(tqdm(range(start_row_idx, end_row_idx + 1))):
            if i == 0:
                continue
            if pd.isna(case_comments[i]):
                continue
            insert_note(team_entry_sheet,
                        (ridx, start_col_idx[0]), case_comments[i])
        print('Done')

        print('Writing deaths columns...')
        # Death
        for i, ridx in enumerate(tqdm(range(start_row_idx, end_row_idx + 1))):
            if i == 0:
                continue
            if pd.isna(death_comments[i]):
                continue
            insert_note(team_entry_sheet,
                        (ridx, start_col_idx[1]), death_comments[i])
        print('Done')

        print('Printing STATE TOTALS for you to copy and paste:')
        print('Case Comment for STATE TOTALS: ', case_comments[0])
        print('Death Comment for STATE TOTALS: ', death_comments[0])

        status_yn = input('Do you want to enter another date? (y/n)')
        if status_yn.lower() == 'y':
            status = True
        else:
            status = False
