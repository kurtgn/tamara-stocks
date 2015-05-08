
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
        if date.weekday() not in [5, 6]:
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
                     target_name, target_ticker, target_data,
                     sheet, row, deal_date):

    '''
    given info about aquiror and target,
    write it into excel file
    and return file object and new row number
    '''
    sheet.write(row, 3, aquiror_ticker)
    sheet.write(row, 4, aquiror_name)
    sheet.write(row, 9, target_ticker)
    sheet.write(row, 10, target_name)

    row += 2
    aq_row = row

    # if yahoo didnt respond to the data request with 404
    if aquiror_data != '404':


        for d in aquiror_data:
            d = d.decode('utf8').split(',')
            date_string = d[0]
            open_value = d[1]
            close_value = d[4]
            # put a DEAL string at the point where the deal happened
            if datetime.strptime(date_string, '%Y-%m-%d') == deal_date:
                sheet.write(aq_row, 2, 'DEAL')
            sheet.write(aq_row, 3, date_string)
            sheet.write(aq_row, 4, open_value)
            sheet.write(aq_row, 5, close_value)
            aq_row += 1

    tar_row = row

    # if yahoo didnt respond to the data request with 404
    if target_data != '404':
        for d in target_data:
            d = d.decode('utf8').split(',')
            date_string = d[0]
            open_value = d[1]
            close_value = d[4]
            # put a DEAL string at the point where the deal happened
            if datetime.strptime(date_string, '%Y-%m-%d') == deal_date:
                sheet.write(tar_row, 8, 'DEAL')
            sheet.write(tar_row, 9, date_string)
            sheet.write(tar_row, 10, open_value)
            sheet.write(tar_row, 11, close_value)
            tar_row += 1

    row = max(tar_row, aq_row) + 3

    return sheet, row