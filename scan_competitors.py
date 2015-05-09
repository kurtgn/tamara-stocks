import xlrd
from datetime import datetime
import pickle
from utils import get_historical_data, Report
import sys


#####################################
###########   Тамаре    #############
#####################################
#
# лаг вперед и лаг назад
# изменяются здесь
FUTURE_LAG = 4
PAST_LAG = 365





# if frompickle is in argv, get array from pickled file
# (to avoid long excel loads)
if len(sys.argv) == 2 and sys.argv[1] == 'frompickle':
    with open('arr.pickle', 'rb') as outp:
        arr = pickle.load(outp)

# otherwise open excel file and pickle it for future use
else:
    file = xlrd.open_workbook('competitors.xlsx')
    sheet = file.sheet_by_name('Competitors list')
    arr = []
    for row in range(1, sheet.nrows-1):
        arr.append(sheet.row_values(row))
    with open('arr.pickle', 'wb') as inp:
        pickle.dump(arr, inp)

r = Report()

deal_num = int(arr[0][0])
end_row_list = []




for row in arr:

    current_deal_num = int(row[0])
    ticker = row[2]
    # get part before dot and strip spaces
    ticker = ticker.split('.')[0].strip()
    company_name = row[3]
    deal_date = row[1]
    deal_date = xlrd.xldate_as_tuple(deal_date, 0)[0:5]
    deal_date = datetime(*deal_date)
    print(deal_num, ticker, deal_date)

    data = get_historical_data(ticker, deal_date, FUTURE_LAG, PAST_LAG)

    end_row = r.write_company(company_name, ticker, data,
                              deal_date, deal_num=current_deal_num)
    end_row_list.append(end_row)

    if current_deal_num != deal_num:
        r.set_baseline_row(max(end_row_list) + 3)
        deal_num = current_deal_num
        end_row_list = []

r.save_file('res_competitors.xls')