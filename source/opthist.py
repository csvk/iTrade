"""
Created on Jan 11, 2017

@author: Souvik
"""

import sqlite3
from datetime import datetime
from dateutil.relativedelta import relativedelta


class OptionsHist:
    """ Historical options data"""

    # variables

    columns = ["symbol", "expiry", "type", "date", "strikeprice", "open", "high", "low", "close", "ltp",
               "settleprice", "number_of_contracts", "turnover", "premium_turnover", "open_interest",
               "change_in_oi", "underlying"]

    def __init__(self, db):

        # variables
        self.conn = None

        print('OptionsHist.__init__', db)
        self.conn = sqlite3.connect(db)

    def read(self, symbol, **kwargs):
        """ Return data based on input """

        qry = '''SELECT * FROM opthist WHERE symbol={}'''.format(symbol)

        for field, val in kwargs.items():
            print('OptionsHist.read', field, ": ", val)
            qry += ' AND {} = {}'.format(field, val)
            # qry = qry + ' AND {} = {}'.format(field, val)

        c = self.conn.cursor()
        c.execute(qry)
        rows = c.fetchall()

    def expiry_month(self, date):

        expiry = datetime.strptime(date, '%Y-%m-%d')

        if expiry.day > 10:
            nxtExpiry = expiry.replace(day=1)
            nxtExpiry = nxtExpiry + relativedelta(months=1)
        else:
            nxtExpiry = expiry

        return nxtExpiry.strftime('%Y-%m-%d')


    def expiry_month_new(self, date, n, midmonth_cutoff=10):

        expiry = datetime.strptime(date, '%Y-%m-%d')
        exp_month_start, exp_month_end = expiry.replace(day=1), expiry.replace(day=1)
        if expiry.day > midmonth_cutoff and n == 0:
            exp_month_start = exp_month_start + relativedelta(months=1)
            exp_month_end = exp_month_end + relativedelta(months=2)
        else:
            exp_month_start = exp_month_start + relativedelta(months=n)
            exp_month_end = exp_month_end + relativedelta(months=n+1)

        return [exp_month_start.strftime('%Y-%m-%d'), exp_month_end.strftime('%Y-%m-%d')]


    def immediate_opt(self, symbol, **kwargs):
        """ Return immediate immediate option data """

        qry = '''SELECT * FROM opthist WHERE symbol={}'''.format(symbol)

        orderby = None

        for field, val in kwargs.items():
            #print("# ", field, val)

            if field == "date":
                qry += ' AND {} > {}'.format("date", val)
                qry += r" AND {} > '{}'".format("expiry", self.expiry_month(val[1:11]))
            elif field == "close":
                if kwargs['type'] == r"'CE'":
                    compare, orderby = '>=', 'ASC'
                    """
                    if val % 100 < 50:
                        val = val - val % 100 + 50
                    else:
                        val = val - val % 100 + 100
                    """
                    val = val - val % 100 + 100
                else:
                    compare, orderby = '<=', 'DESC'
                    """
                    if val % 100 < 50:
                        val -= val % 100
                    else:
                        val = val - val % 100 + 50
                    """
                    val = val - val % 100 + 50

                qry += ' AND {} {} {}'.format("strikeprice", compare, val)
            else:
                qry += ' AND {} = {}'.format(field, val)

        qry += ' ORDER BY expiry ASC, date ASC, strikeprice {} LIMIT 1'.format(orderby)

        c = self.conn.cursor()
        c.execute(qry)
        rows = c.fetchall()

        return dict(zip(OptionsHist.columns, rows[0]))


    def nth_opt(self, nth_expiry, nth_strike, midmonth_cutoff, entrytime, **kwargs):
        """ Return nth option data """

        qry = '''SELECT * FROM opthist WHERE'''

        for field, val in kwargs.items():
            #print("# ", field, val)

            if field == "symbol":
                qry += ' {} = {}'.format(field, val)
            elif field == "date":
                compare = '=' if entrytime=='eod' else '>'
                qry += ' AND {} {} {}'.format("date", compare, val)
                qry += r" AND {} BETWEEN '{}' AND '{}'".format(
                    "expiry",
                    self.expiry_month_new(val[1:11], nth_expiry, midmonth_cutoff)[0],
                    self.expiry_month_new(val[1:11], nth_expiry, midmonth_cutoff)[1]
                )
            elif field == "close":
                if kwargs['type'] == r"'CE'":
                    val = val - val % 100 + 100 * (nth_strike + 1)
                else:
                    val = val - val % 100 - 100 * nth_strike
                qry += ' AND strikeprice = {}'.format(val)
            else:
                qry += ' AND {} = {}'.format(field, val)

        qry += ' ORDER BY date ASC, expiry ASC  LIMIT 1'

        c = self.conn.cursor()
        c.execute(qry)
        rows = c.fetchall()

        return dict(zip(OptionsHist.columns, rows[nth_expiry]))

    def option_exit_data(self, option_entry, exit_date):

        qry = """SELECT * FROM opthist WHERE symbol = '{}' AND expiry = '{}' AND type = '{}' AND date > '{}' AND """ \
              """strikeprice = {} ORDER BY date ASC LIMIT 1""".format(option_entry['symbol'], option_entry['expiry'],
                                                              option_entry['type'], exit_date,
                                                              option_entry['strikeprice'])

        c = self.conn.cursor()
        c.execute(qry)
        rows = c.fetchall()

        return dict(zip(OptionsHist.columns, rows[0]))

    def option_data_between_entry_exit(self, trade):

        qry = """SELECT * FROM opthist WHERE symbol = '{}' AND expiry = '{}' AND type = '{}' AND date > '{}' AND """ \
              """date < '{}' AND strikeprice = {} ORDER BY date ASC""".format(trade['entry']['symbol'],
                                                                            trade['entry']['expiry'],
                                                                            trade['entry']['type'],
                                                                            trade['entry']['date'],
                                                                            trade['exit']['date'],
                                                                            trade['entry']['strikeprice'])

        c = self.conn.cursor()
        c.execute(qry)
        rows = c.fetchall()

        transposed_rows = list(zip(*rows))

        d = dict(zip(OptionsHist.columns, transposed_rows))

        return d


    def option_data_post_entry(self, trade):
        """Return option data for all dates till expiry starting from entry date + 1
                Return format: {'symbol': (day1, day2, ..., expiryday),
                                'expiry': (day1, day2, ..., expiryday),
                                'type': (), 'strikeprice': (), 'open': (), 'high': (), 'low': (), 'close': (),'ltp': (),
                                . . . . ., 'position': {day1 in YYYY-MM-DD: position index,
                                                        day2 in YYYY-MM-DD: position index,
                                                        .....,
                                                        expiryday in YYYY-MM-DD: position index}
        """

        qry = """SELECT * FROM opthist WHERE symbol = '{}' AND expiry = '{}' AND type = '{}' AND date > '{}' AND """ \
              """strikeprice = {} ORDER BY date ASC""".format(trade['entry']['symbol'],
                                                              trade['entry']['expiry'],
                                                              trade['entry']['type'],
                                                              trade['entry']['date'],
                                                              trade['entry']['strikeprice'])

        c = self.conn.cursor()
        c.execute(qry)
        rows = c.fetchall()

        transposed_rows = list(zip(*rows))

        d = dict(zip(OptionsHist.columns, transposed_rows))

        return d

