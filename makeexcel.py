from urllib.request import urlopen
from urllib.parse import urlencode
from urllib.error import HTTPError
from datetime import timedelta, datetime
import pdb
from xlwt import Workbook, easyxf

def get_historical_data(ticker, deal_date, future_lag, past_lag):
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
    increment days while skipping weekends
    '''
    while days_amt > 0:
        date += timedelta(days=1)
        if date.weekday() not in [5, 6]:
            days_amt -= 1
    return date