import xlrd
from datetime import datetime
import pickle
from utils import get_historical_data, Report, extract_columns_add_close_return
import sys
import numpy as np
from scipy.stats import linregress


#####################################
##########   Параметры   ############
#####################################
#
# лаг вперед и лаг назад
# изменяются здесь
FUTURE_LAG = 4
PAST_LAG = 500
EVENT_WINDOW = 100





# if frompickle is in argv, get array from pickled file
# (to avoid long excel loads)
if len(sys.argv) == 2 and sys.argv[1] == 'frompickle':
    with open('arr.pickle', 'rb') as outp:
        arr = pickle.load(outp)

# otherwise open excel file and pickle it for future use
else:


    if len(sys.argv) != 2:
        print('Enter file to scan.\n\nExample:\npython scan.py final.xlsx')
        quit()

    file = sys.argv[1]
    try:
        file = xlrd.open_workbook(file)
    except FileNotFoundError:
        print('File {} not found.'.format(file))
        quit()

    try:
        sheet = file.sheet_by_name('Competitors list')
    except:
        print('Sheet "Competitors list" not found.')
        quit()
    arr = []
    for row in range(1, sheet.nrows-1):
        arr.append(sheet.row_values(row))
    with open('arr.pickle', 'wb') as inp:
        pickle.dump(arr, inp)

r = Report()





# remember first deal number
deal_num = int(arr[0][0])

# iterate over data array
for row in arr:
    # check current deal number
    current_deal_num = int(row[0])
    if current_deal_num != deal_num:
        r.start_new_row()
        deal_num = current_deal_num

    # get ticker value
    ticker = row[2]
    # get part before dot and strip spaces
    ticker = ticker.split('.')[0].strip()

    # get other params
    company_name = row[3]
    deal_date = row[1]
    deal_date = xlrd.xldate_as_tuple(deal_date, 0)[0:5]
    deal_date = datetime(*deal_date)
    print(deal_num, ticker, deal_date)

    # get historical data from yahoo.finance
    company_data = get_historical_data(ticker, deal_date, FUTURE_LAG, PAST_LAG)
    market_data = get_historical_data('^GSPC', deal_date, FUTURE_LAG, PAST_LAG)

    # clean it up a bit, add 'return' column
    company_data = extract_columns_add_close_return(company_data)
    market_data = extract_columns_add_close_return(market_data)

    # make those sets of the same size in case one is bigger than other
    normalized_length = min(len(company_data), len(market_data))
    company_data = company_data[:normalized_length]
    market_data = market_data[:normalized_length]

    # extract 'return' rows
    company_return = [row[2] for row in company_data]
    market_return = [row[2] for row in market_data]

    # extract estimation period: all observations that were earlier than the event window
    comp_est_period = company_return[EVENT_WINDOW:]
    market_est_period = market_return[EVENT_WINDOW:]

    # calculate linear regression over the estimation period
    beta, alpha, r_value, p_value, std_err = linregress(market_est_period, comp_est_period)

    # extrapolate the regression into the event window:
    # put market_data through the regression values and
    # calculate expected company return
    company_expected_return = []
    for idx, val in enumerate(company_return[:EVENT_WINDOW]):
        exp_val = market_return[idx] * beta + alpha
        company_expected_return.append(exp_val)

    # put data into numpy format
    company_expected_return_event_window = np.array(company_expected_return)
    company_return_event_window = np.array(company_return[:EVENT_WINDOW])

    # calculate abnormal return
    abnornal_return = company_return_event_window - company_expected_return_event_window

    # and calculate related values
    standard_deviation = np.std(abnornal_return)
    CAR = sum(abnornal_return)
    t_stat = CAR / ((EVENT_WINDOW * standard_deviation ** 2) ** 0.5)

    data_for_print = {
        'company_data': company_data,
        'market_data': market_data,
        'beta': beta,
        'alpha': alpha,
        'r_value': r_value,
        'p_value': p_value,
        'std_err': std_err,
        'company_expected_return': company_expected_return,
        'abnormal_return': abnornal_return,
        'standard_deviation': standard_deviation,
        'CAR': CAR,
        't_stat': t_stat
    }

    r.write_data(company_name, ticker, data_for_print,
                 deal_date, deal_num=current_deal_num)



r.save_file('res_competitors.xls')