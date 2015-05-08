from flask import Flask, render_template, request, redirect, url_for, flash
from urllib.parse import urlencode
from urllib.request import urlopen, quote
import csv
import requests
import json
from datetime import datetime
from makeexcel import makeexcel

app = Flask(__name__)
app.config['DEBUG'] = True
app.secret_key='sjkhdkjashdiashpd98ahsduh'
import pdb



@app.route("/", methods=['GET','POST'])
def hello():
    print(request.files)
    return render_template('uploader.html')


@app.route("/onecompany/", methods=['GET','POST'])
def onecompany():
    if request.method == 'GET':
        return render_template('enter_name_or_code.html')
    c_name = request.form.get('c_name')
    c_ticker = request.form.get('c_ticker')

    if not c_name and not c_ticker:
        flash('Введите название или тикер')
        return redirect(url_for('onecompany'))

    if not c_ticker:
        t, msg = find_single_ticker(c_name)
        if not t:
            flash(msg)
            return redirect(url_for('onecompany'))
        c_ticker = t



    deal_date = request.form.get('deal_date')
    future_lag = int(request.form.get('future_lag'))
    past_lag = int(request.form.get('past_lag'))
    if not deal_date or not future_lag or not past_lag:
        flash('Введите дату, лаг вперед и лаг назад.')
        return redirect(url_for('onecompany'))

    try:
        deal_date = datetime.strptime(deal_date, '%d.%m.%Y')
    except ValueError:
        flash('Некорректный формат даты.')
        return redirect(url_for('onecompany'))

    makeexcel(c_ticker, deal_date, future_lag, past_lag)
    return 'ok'



def find_single_ticker(name):
    tickers = search_ticker_by_name(name)
    if len(tickers) == 0:
        print('result length 0. generating options...')
        options = gen_name_options(name)

        for option in options:
            print('option: ', option)
            ticker_options = search_ticker_by_name(option)

            if len(ticker_options) > 1:
                print('found more than 1 result')
                msg = 'По названию "{}" найдено несколько компаний: <br/>'
                msg = msg.format(name)
                for el in tickers:
                    pair = ' {}:{}<br/> '.format(el['name'], el['symbol'])
                    msg += pair
                return None, msg

            elif len(ticker_options) == 1:
                t = ticker_options[0]['symbol']
                print('found exactly 1 result')
                msg = 'Нашелся тикер: {}'.format(t)
                return t, msg

            else:
                print('result length 0')
                pass

        options = '; '.join(options)
        msg = 'По названию "{}" найдено 0 компаний. Испробованы варианты: {}'
        msg = msg.format(name, options)
        return None, msg

    elif len(tickers) > 1:
        msg = 'По названию "{}" найдено несколько компаний:<br/>'
        msg = msg.format(name)
        for el in tickers:
            pair = ' {}:{} <br/>'.format(el['name'], el['symbol'])
            msg += pair
        return None, msg

    else:
        t = tickers[0]['symbol']
        msg = 'Нашелся тикер: {}'.format(t)
        return t, msg





def search_ticker_by_name(name):
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
    if 'inc' in name.lower():
        beforeInc = name.lower().split(' inc')[0]
        options = [
            beforeInc + ', inc',
            beforeInc + ' inc.',
            beforeInc + ', inc',
            beforeInc + ' inc'
        ]
        return options
    return [name,]



if __name__ == "__main__":
    app.run()



