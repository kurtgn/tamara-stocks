import xlrd, pdb
from datetime import datetime
import pickle
from utils import find_single_ticker, get_historical_data, append_excel_sheet
import sys
from xlwt import Workbook


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

w = Workbook()
sheet = w.add_sheet('исторические данные')
sheet.write(0, 4, 'AQUIROR')
sheet.write(0, 10, 'TARGET')
sheet.write(2, 3, 'Date')
sheet.write(2, 4, 'Open')
sheet.write(2, 5, 'Closed')
sheet.write(2, 9, 'Date')
sheet.write(2, 10, 'Open')
sheet.write(2, 11, 'Closed')

sheet_row = 4


for row in arr:
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
        aquiror_data = get_historical_data(aquiror_ticker, deal_date, 4, 365)
    else:
        aquiror_data = []

    # get historical data for the ticker if supplied, or [] otherwise
    if target_ticker:
        target_data = get_historical_data(target_ticker, deal_date, 4, 365)
    else:
        target_data = []

    # write a neat table to the excel sheet
    sheet, sheet_row = append_excel_sheet(aquiror_name, aquiror_ticker, aquiror_data,
                                    target_name, target_ticker, target_data,
                                    sheet, sheet_row, deal_date)


w.save('results.xls')