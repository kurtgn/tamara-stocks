import xlrd, pdb
from datetime import datetime
import pickle
from utils import find_single_ticker, get_historical_data, append_excel_sheet
import sys
from xlwt import Workbook

'''
arr = [1,2,3,4,5]
with open('arr.pickle', 'wb') as inp:
    pickle.dump(arr, inp)

quit()

'''
#pdb.set_trace()
if len(sys.argv) == 2 and sys.argv[1] == 'frompickle':
    with open('arr.pickle', 'rb') as outp:
        arr = pickle.load(outp)

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

counter = 0

for row in arr:
    aquiror_name = row[3]
    target_name = row[2]
    deal_date = row[13]
    deal_date = xlrd.xldate_as_tuple(deal_date, 0)[0:5]
    deal_date = datetime(*deal_date)
    aquiror_ticker, msg1 = find_single_ticker(aquiror_name)
    target_ticker, msg2 = find_single_ticker(target_name)
    print('\n\n'+ msg1 + '\n' + msg2)
    print(aquiror_ticker, target_ticker, deal_date)
    if aquiror_ticker:
        aquiror_data = get_historical_data(aquiror_ticker, deal_date, 4, 365)
    else:
        aquiror_data = []

    if target_ticker:
        target_data = get_historical_data(target_ticker, deal_date, 4, 365)
    else:
        target_data = []



    sheet, sheet_row = append_excel_sheet(aquiror_name, aquiror_ticker, aquiror_data,
                                    target_name, target_ticker, target_data,
                                    sheet, sheet_row, deal_date)
    print('counter:',counter)
    counter += 1

w.save('results.xls')