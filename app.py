from flask import Flask, render_template, request, redirect, url_for, flash, send_file

from datetime import datetime
from utils import get_historical_data, write_file, find_single_ticker


app = Flask(__name__)
app.config['DEBUG'] = True
app.secret_key='thisisnotasecret'


@app.route("/", methods=['GET','POST'])
def main():
    if request.method == 'GET':
        return render_template('enter_name_or_code.html')

    # get args
    company_name = request.form.get('c_name')
    ticker = request.form.get('c_ticker')

    if not company_name and not ticker:
        flash('Введите название или тикер')
        return redirect(url_for('main'))

    # if ticker is not provided, find it by name.
    # If not successful, flash error and redirect
    if not ticker:
        t, msg = find_single_ticker(company_name)
        if not t:
            flash(msg)
            return redirect(url_for('main'))
        ticker = t

    # get date and lag args
    deal_date = request.form.get('deal_date')
    future_lag = int(request.form.get('future_lag'))
    past_lag = int(request.form.get('past_lag'))

    # validate existence
    if not deal_date or not future_lag or not past_lag:
        flash('Введите дату, лаг вперед и лаг назад.')
        return redirect(url_for('main'))

    # validate date format
    try:
        deal_date = datetime.strptime(deal_date, '%d.%m.%Y')
    except ValueError:
        flash('Некорректный формат даты.')
        return redirect(url_for('main'))

    # validate weekdays
    if deal_date.weekday() in [5, 6]:
        msg = '{} - выходной день. В этот день не было торгов. Выберите рабочий день.'
        flash(msg.format(deal_date.strftime('%d.%m.%Y')))
        return redirect(url_for('main'))

    hist_data = get_historical_data(ticker, deal_date, future_lag, past_lag)

    # get_historical_data() will return '404' if the request is unsuccessful
    # i.e. wrong ticker specified
    if hist_data == '404':
        flash('Такого тикера не существует.')
        return redirect(url_for('main'))

    if not company_name:
        company_name = ticker
    filename = ticker + '.xls'
    write_file(hist_data, deal_date, company_name, filename)

    return send_file(filename,
                     as_attachment=True,
                     attachment_filename=filename)

if __name__ == "__main__":
    app.run()



