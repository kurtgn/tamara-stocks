import xlrd, pdb
from datetime import datetime
import pickle
from utils import find_single_ticker, get_historical_data, append_excel_sheet, Report
import sys


# if frompickle is in argv, get array from pickled file
# (to avoid long excel loads)
if len(sys.argv) == 2 and sys.argv[1] == 'frompickle':
    with open('arr.pickle', 'rb') as outp:
        arr = pickle.load(outp)

# otherwise open excel file and pickle it for future use
else:
    file = xlrd.open_workbook('input.xlsx')
    sheet = file.sheet_by_name('Final')
    arr = []
    for row in range(1, sheet.nrows-1):
        arr.append(sheet.row_values(row))
    with open('arr.pickle', 'wb') as inp:
        pickle.dump(arr, inp)

r = Report()


for row in arr[0:3]:
    aquiror_name = row[3]
    target_name = row[2]
    deal_date = row[13]
    deal_date = xlrd.xldate_as_tuple(deal_date, 0)[0:5]
    deal_date = datetime(*deal_date)

    # get aquiror and target tickers
    aquiror_ticker, msg1 = find_single_ticker(aquiror_name)
    target_ticker, msg2 = find_single_ticker(target_name)
    #print('\n\n'+ msg1 + '\n' + msg2)
    print(aquiror_ticker, target_ticker, deal_date)

    # get historical data for the ticker if supplied, or [] otherwise
    if aquiror_ticker:
        aquiror_data = get_historical_data(aquiror_ticker, deal_date, 4, 5)
    else:
        aquiror_data = []

    # get historical data for the ticker if supplied, or [] otherwise
    if target_ticker:
        target_data = get_historical_data(target_ticker, deal_date, 4, 5)
    else:
        target_data = []

    report_row1 = r.write_company(aquiror_name, aquiror_ticker, aquiror_data, deal_date)
    report_row2 = r.write_company(target_name, target_ticker, target_data, deal_date)

    new_row = max(report_row1, report_row2) + 3
    r.set_baseline_row(new_row)


r.save_file('results.xls')