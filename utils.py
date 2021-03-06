
import requests
import json
from urllib.request import urlopen
from urllib.parse import urlencode, quote
from urllib.error import HTTPError
from datetime import timedelta, datetime
import pdb
from xlwt import Workbook, easyxf

def get_historical_data(ticker, deal_date, future_lag, past_lag):
    '''
    given a ticker and date and lags, return list of historical data
    '''
    yUrl = 'http://real-chart.finance.yahoo.com/table.csv?'
    endDate = add_workdays(deal_date, future_lag)
    startDate = deal_date - timedelta(days=past_lag)

    queryString = urlencode(
        {'s': ticker,
        'a': startDate.month - 1 if startDate.month != 0 else 11,
        'b': startDate.day,
        'c': startDate.year,
        'd': endDate.month - 1 if endDate.month != 0 else 11,
        'e': endDate.day,
        'f': endDate.year,
        'g': 'd', # daily values
        'ignore': '.csv' # just stolen from yahoo example
        }
    )
    queryString = yUrl + queryString
    print(queryString)
    try:
        resp = urlopen(queryString)
    except HTTPError:
        return '404'

    result_list = resp.readlines()[1:]
    return result_list

def write_file(result_list, deal_date, company_name, filename):
    '''
    given a list, put it into excel file.
    deal_date specifies a string which will be rendered as bold
    company_name and filename are self-explanatory
    '''
    w = Workbook()
    sheet = w.add_sheet(company_name)
    row = 2
    boldfont = easyxf(strg_to_parse='font: bold on')
    normalfont = easyxf(strg_to_parse='')
    sheet.write(0, 0, company_name)
    sheet.write(1, 0, 'Date')
    sheet.write(1, 1, 'Open')
    sheet.write(1, 2, 'Close')

    for line in result_list:
        elements = line.decode('utf8').split(',')
        date_string = elements[0]
        open_value = elements[1]
        close_value = elements[4]
        if datetime.strptime(date_string, '%Y-%m-%d') == deal_date:
            style = boldfont
        else:
            style = normalfont
        sheet.write(row, 0, date_string, style)
        sheet.write(row, 1, open_value, style)
        sheet.write(row, 2, close_value, style)
        row += 1

        print(date_string, open_value, close_value)
    w.save(filename)




def add_workdays(date, days_amt):
    '''
    increment days in a given date while skipping weekends
    '''
    while days_amt > 0:
        date += timedelta(days=1)
        if date.weekday() not in [5, 6] and not (date.month == 12 and date.day == 25):
            days_amt -= 1
    return date





def find_single_ticker(name):
    '''
    finds ticker by name
    returns ticker, msg
    ticker = None if several options are found and cannot be narrowed down
    mgs is a comment for flashing
    '''

    tickers = get_list_from_yahoo(name)
    # if nothing found try some other name options
    if len(tickers) == 0:
        options = gen_name_options(name)

        for option in options:
            ticker_options = get_list_from_yahoo(option)

            # if there are several options, try to narrow down
            # by removing tickers with dots
            if len(ticker_options) > 1:
                # get just symbols
                names = [t['symbol'] for t in ticker_options]
                # remove dots
                names = [n for n in names if '.' not in n]
                # if length is 1 now - then this is what we need
                if len(names) == 1:
                    t = names[0]
                    msg = 'Нашелся тикер: {}'.format(t)
                    return t, msg

                # if there are still several tickers left, return None
                msg = 'По названию "{}" найдено несколько компаний: <br/>'
                msg = msg.format(name)
                for el in ticker_options:
                    pair = ' {} : {}<br/> '.format(el['name'], el['symbol'])
                    msg += pair
                return None, msg

            elif len(ticker_options) == 1:
                # yay, we got what we needed!
                t = ticker_options[0]['symbol']
                msg = 'Нашелся тикер: {}'.format(t)
                return t, msg

            else:
                pass

        # nothing helped. Returning with empty hands
        options = '; '.join(options)
        msg = 'По названию "{}" найдено 0 компаний. Испробованы варианты: {}'
        msg = msg.format(name, options)
        return None, msg

    # if we got several tickers at the first try
    # first we narrow it down by removing tickers with dots
    elif len(tickers) > 1:
        # get just symbols
        names = [t['symbol'] for t in tickers]
        # remove dots
        names = [n for n in names if '.' not in n]
        # if length is 1 now - then this is what we need
        if len(names) == 1:
            t = names[0]
            msg = 'Нашелся тикер: {}'.format(t)
            return t, msg

        # if above didnt help narrow our list down
        msg = 'По названию "{}" найдено несколько компаний:<br/>'
        msg = msg.format(name)
        for el in tickers:
            pair = ' {} : {} <br/>'.format(el['name'], el['symbol'])
            msg += pair
        return None, msg

    # if we found jus what we needed from the first try
    else:
        t = tickers[0]['symbol']
        msg = 'Нашелся тикер: {}'.format(t)
        return t, msg



def get_list_from_yahoo(name):
    '''
    given a company name, perform a query on a yahoo API
    to retrieve possible tickers
    '''
    searchUrl = 'http://d.yimg.com/autoc.finance.yahoo.com/autoc?' \
                    'query={0}&callback=YAHOO.Finance.SymbolSuggest.ssCallback'
    escapedName = quote(name)
    searchUrl = searchUrl.format(escapedName)
    response = requests.get(searchUrl)
    results = response.text.split('Callback')[1]
    results = results[1:-1]
    jsonResponse = json.loads(results)
    resultList = jsonResponse['ResultSet']['Result']
    return resultList


def gen_name_options(name):
    '''
    add some name options
    i.e. Company Inc -> Company, Inc   Company Inc.    Company, Inc.
    '''
    if 'inc' in name.lower():
        beforeInc = name.lower().split(' inc')[0]
        options = [
            beforeInc + ', inc',
            beforeInc + ' inc.',
            beforeInc + ', inc',
            beforeInc + ' inc',
            beforeInc
        ]
        return options
    if 'corp' in name.lower():
        beforeInc = name.lower().split(' corp')[0]
        options = [
            beforeInc + ', corp',
            beforeInc + ' corp.',
            beforeInc + ', corp',
            beforeInc + ' corp',
            beforeInc
        ]
        return options
    return [name,]


def append_excel_sheet(aquiror_name, aquiror_ticker, aquiror_data,
                       sheet, row, col, deal_date):

    '''
    given info about a company,
    write it into excel file
    and return file object and new row+col number
    '''
    sheet.write(row, 3, aquiror_ticker)
    sheet.write(row, 4, aquiror_name)

    row += 2


    # if yahoo didnt respond to the data request with 404
    if aquiror_data != '404':


        for d in aquiror_data:
            d = d.decode('utf8').split(',')
            date_string = d[0]
            open_value = d[1]
            close_value = d[4]
            # put a DEAL string at the point where the deal happened
            if datetime.strptime(date_string, '%Y-%m-%d') == deal_date:
                sheet.write(row, col, 'DEAL')
            sheet.write(row, col+1, date_string)
            sheet.write(row, col+2, open_value)
            sheet.write(row, col+3, close_value)
            row += 1

    col += 5

    return sheet, row, col


class Report(object):
    def __init__(self):
        self.w = Workbook()
        self.sheet = self.w.add_sheet('results')
        self.row = 0
        self.col = 2
        self.baseline_row = 0
        self.boldfont = easyxf(strg_to_parse='font: bold on')
        self.normalfont = easyxf(strg_to_parse='')
        self.row_history = []

    def write_company(self, name, ticker, data, deal_date, deal_num=None):
        '''
        given info about a company,
        write a nice column of data
        and anvance 5 columns right
        '''
        self.row = self.baseline_row

        self.sheet.write(self.row, self.col, ticker)
        self.sheet.write(self.row, self.col+1, name)

        self.row += 2

        # if yahoo didnt respond to the data request with 404
        if data != '404':
            for d in data:
                d = d.decode('utf8').split(',')
                date_string = d[0]
                open_value = d[1]
                close_value = d[4]
                # put a DEAL string at the point where the deal happened
                if datetime.strptime(date_string, '%Y-%m-%d') == deal_date:
                    style = self.boldfont
                else:
                    style = self.normalfont
                self.sheet.write(self.row, self.col, date_string, style)
                self.sheet.write(self.row, self.col+1, open_value, style)
                self.sheet.write(self.row, self.col+2, close_value, style)

                # a quick and dirty workaround
                # to write deal number to the first column
                # we catch exceptions here because
                # we want to enable overwriting

                if deal_num:
                    try:
                        self.sheet.write(self.row, 0, deal_num)
                    except Exception:
                        pass

                self.row += 1

        self.col+=4
        return self.row


    def write_data(self, name, ticker, data_for_print, deal_date, deal_num=None):
        '''
        given info about a company,
        write a nice column of data
        and anvance 5 columns right
        '''
        self.row = self.baseline_row

        self.sheet.write(self.row, self.col, ticker)
        self.sheet.write(self.row, self.col+1, name)

        # start commenting here if needed short version
        # '''

        self.sheet.write(self.row+1, self.col, 'date')
        self.sheet.write(self.row+1, self.col+1, 'adj close')
        self.sheet.write(self.row+1, self.col+2, 'com return')
        self.sheet.write(self.row+1, self.col+3, 'mrkt return')

        self.row += 2

        # print regression values
        self.sheet.write(self.row, self.col+4, 'beta', self.boldfont)
        self.sheet.write(self.row+1, self.col+4, data_for_print['beta'])
        self.sheet.write(self.row+2, self.col+4, 'alpha', self.boldfont)
        self.sheet.write(self.row+3, self.col+4, data_for_print['alpha'])
        self.sheet.write(self.row+4, self.col+4, 'r_value', self.boldfont)
        self.sheet.write(self.row+5, self.col+4, data_for_print['r_value'])
        self.sheet.write(self.row+6, self.col+4, 'p_value', self.boldfont)
        self.sheet.write(self.row+7, self.col+4, data_for_print['p_value'])
        self.sheet.write(self.row+8, self.col+4, 'std_err', self.boldfont)
        self.sheet.write(self.row+9, self.col+4, data_for_print['std_err'])

        for idx, d in enumerate(data_for_print['company_data']):
            date = d[0].strftime('%d.%m.%Y')
            adj_close = d[1]
            ret_value = d[2]

            # put bold font at the point where the deal happened
            if d[0] == deal_date:
                style = self.boldfont
            else:
                style = self.normalfont
            self.sheet.write(self.row, self.col, date, style)
            self.sheet.write(self.row, self.col+1, adj_close, style)
            self.sheet.write(self.row, self.col+2, ret_value, style)

            market_ret_value = data_for_print['market_data'][idx][2]
            self.sheet.write(self.row, self.col+3, market_ret_value, style)

            # a quick and dirty workaround
            # to write deal number to the first column
            # we catch exceptions here because
            # we want to enable overwriting cells

            if deal_num:
                try:
                    self.sheet.write(self.row, 0, deal_num)
                except Exception:
                    pass

            self.row += 1


        self.row_history.append(self.row)

        self.col += 5


        self.row = self.baseline_row + 2

        self.sheet.write(self.row-1, self.col, 'comp exp ret')
        for idx, val in enumerate(data_for_print['company_expected_return']):
            self.sheet.write(self.row+idx, self.col, val)

        self.col += 1
        self.sheet.write(self.row-1, self.col, 'abnormal return')
        for idx, val in enumerate(data_for_print['abnormal_return']):
            self.sheet.write(self.row+idx, self.col, val)


        # end commenting here if needed short version
        # '''
        # self.row += 2

        self.col += 1
        self.sheet.write(self.row, self.col, 'std deviation', self.boldfont)
        self.sheet.write(self.row+1, self.col, data_for_print['standard_deviation'])
        self.sheet.write(self.row, self.col+1, 'CAR', self.boldfont)
        self.sheet.write(self.row+1, self.col+1, data_for_print['CAR'])
        self.sheet.write(self.row, self.col+2, 't stat', self.boldfont)
        self.sheet.write(self.row+1, self.col+2, data_for_print['t_stat'])

        self.col += 4

        self.row_history.append(15)

    def set_baseline_row(self, val):
        '''
        something like carriage return.
        Set new baseline row and reset column number to starting value.
        '''
        self.baseline_row = val
        self.col = 2

    def start_new_row(self):
        '''
        take maximum row from row history
        set it as new baseline row
        reset self.col to starting value
        '''
        self.baseline_row = max(self.row_history) + 3
        self.col = 2

    def save_file(self, filename):
        self.w.save(filename)


def extract_columns_add_close_return(data):
    new_data = []
    # if yahoo didnt respond to the data request with 404
    if data != '404':
        for idx, row in enumerate(data):
            row = row.decode('utf8').split(',')
            date_string = row[0]
            date_string = datetime.strptime(date_string, '%Y-%m-%d')
            adj_close_value = float(row[6])
            try:
                next_row = data[idx+1]
                next_row = next_row.decode('utf8').split(',')
                older_adj_close_value = float(next_row[6])
                ret_value = adj_close_value / older_adj_close_value - 1
            except:
                ret_value = 0
            new_data.append([date_string, adj_close_value, ret_value])
    return new_data
