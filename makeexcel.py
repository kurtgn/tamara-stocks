from urllib.request import urlopen
from urllib.parse import urlencode
from datetime import timedelta, datetime
import pdb
from xlwt import Workbook, easyxf

def makeexcel(ticker, deal_date, future_lag, past_lag):
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
    resp = urlopen(queryString)

    w = Workbook()
    sheet = w.add_sheet(ticker)
    row = 1
    boldfont = easyxf(strg_to_parse='font: bold on')
    normalfont = easyxf(strg_to_parse='')

    sheet.write(0, 0, 'Date')
    sheet.write(0, 1, 'Open')
    sheet.write(0, 2, 'Adj Close')

    for line in resp.readlines()[1:]:
        elements = line.decode('utf8').split(',')
        date_string = elements[0]
        open_value = elements[1]
        close_value = elements[6]
        if datetime.strptime(date_string, '%Y-%m-%d') == deal_date:
            style = boldfont
        else:
            style = normalfont
        sheet.write(row, 0, date_string, style)
        sheet.write(row, 1, open_value, style)
        sheet.write(row, 2, close_value, style)
        row += 1

        print(date_string, open_value, close_value)

    filename = ticker+'.xls'
    w.save(filename)
    return filename



def add_workdays(date, days_amt):
    '''
    increment days while skipping weekends
    '''
    while days_amt > 0:
        date += timedelta(days=1)
        if date.weekday() not in [5, 6]:
            days_amt -= 1
    return date