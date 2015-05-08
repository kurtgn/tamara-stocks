import xlrd, pdb
from datetime import datetime
import pickle
from utils import find_single_ticker
import sys


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



for row in arr:
    aquiror = row[3]
    target = row[2]
    date_int = row[13]
    date_tuple = xlrd.xldate_as_tuple(date_int, 0)[0:5]
    date_normal = datetime(*date_tuple)
    aquiror_ticker, msg1 = find_single_ticker(aquiror)
    target_ticker, msg2 = find_single_ticker(target)
    print('\n\n'+ msg1 + '\n' + msg2)
    print(aquiror_ticker, target_ticker, date_normal)

