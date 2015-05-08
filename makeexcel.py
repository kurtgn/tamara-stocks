from urllib.request import urlopen
from urllib.parse import urlencode
from datetime import timedelta


def makeexcel(ticker, deal_date, future_lag, past_lag):
    yUrl = 'http://real-chart.finance.yahoo.com/table.csv?'
    endDate = deal_date + timedelta(days=future_lag)
    startDate = deal_date - timedelta(days=past_lag)
    startD = 1
    startM = 1
    startY = 2000
    endD = 1
    endM = 1
    endY = 2010
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
    #pdb.set_trace()
    resp = urlopen(queryString)
    for line in resp.readlines():
        print(line.decode('utf-8'))
