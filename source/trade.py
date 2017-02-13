"""
Created on Jan 15, 2017
◘
@author: Souvik
"""

import os
import openpyxl as xl
from datetime import datetime
from shutil import copy2
import timelog
import numpy as np
from opthist import OptionsHist

class TradeTest:
    """Main Class for performing a trading backtest"""

    def __init__(self, root, dataFile, optionsDB):

        # variables

        self.nth_expiry, self.nth_strike, self.midmonth_cutoff, self.entrytime = 0, 0, 10, 'sod'

        self.dataFile = dataFile
        self.testFile = self.dataFile.replace(".xlsx", "_Test_" + str(datetime.now().strftime('%Y-%m-%d %H.%M.%S')) +
                                              ".xlsx")
        self.optionsFile = optionsDB

        # init

        # starttime = timelog.now()

        os.chdir(root)
        copy2(self.dataFile, self.testFile) # Create test file for analysis

        print('TradeTest.__init__', self.testFile)

        self.dataHist = DataHist(self.testFile)

        # print(data.bar(len(data.date) - 1)['date'])

        self.oh = OptionsHist(optionsDB)


    def signals(self, latitude=0):
        """ Return trading signals
        Sample Return Format: ["Buy", None, None, "Sell", None, ....]
        Call Format: signals(percentage latitude to compare close against lower high an higher low)
        """



        buy = np.logical_and(self.dataHist.value['BarUpDown'] == 1,
                              self.dataHist.value['Close'] > self.dataHist.value['LH Value'] * (1 + latitude))
        sell = np.logical_and(self.dataHist.value['BarUpDown'] == -1,
                              self.dataHist.value['Close'] < self.dataHist.value['HL Value'] * (1 - latitude))

        signals = []
        bought, sold = False, False

        for i in range(0, self.dataHist.barcount):
            if buy[i]:
                if not bought:
                    signals.append("Buy")
                    bought, sold = True, False
                else:
                    signals.append(None)
            elif sell[i]:
                if not sold:
                    signals.append("Sell")
                    sold, bought = True, False
                else:
                    signals.append(None)
            else:
                signals.append(None)

        return signals


    def trades(self, signals):
        """ Return trades
        Return Format: [{'Symbol' : <>, 'Date' : <>, 'Close' : <>, 'Signal' : <>}, {...}, {...}, ....]
        Call Format: trades(returned list from signals() method)
        """

        trades = []

        for i in range(0,self.dataHist.barcount):
            if signals[i] in ("Buy", "Sell"):
                trades.append(
                    {
                        'Symbol': self.dataHist.value['Symbol'][i],
                        'Date': self.dataHist.value['Date'][i],
                        'Close': self.dataHist.value['Close'][i],
                        'Signal': signals[i]
                    }
                )

        return trades


    def sel_options(self, trades):

        optTrades = []

        for i in range(0, len(trades)):

            optSymbol = 'NIFTY' if trades[i]['Symbol'] == 'NSENIFTY' else trades[i]['Symbol']
            optType = "CE" if trades[i]['Signal'] == "Buy" else "PE"

            # print(trades[i]['Date'], trades[i]['Close'], optType)
            optTrades.append(
                self.oh.immediate_opt(
                    r"'{}'".format(optSymbol),
                    type=r"'{}'".format(optType),
                    date=r"'{}'".format(trades[i]['Date']),
                    close=trades[i]['Close']
                )
            )


    def signal_data(self, bar, signal):
        """ Returns signal data (similar to trades function
        Return Format: {'Symbol' : <>, 'Date' : <>, 'Close' : <>, 'Signal' : <>}
        Call Format: signal_data(bar_index, signal)
        """

        return {
            'Symbol': self.dataHist.value['Symbol'][bar],
            'Date': self.dataHist.value['Date'][bar],
            'Close': self.dataHist.value['Close'][bar],
            'Signal': signal
        }


    def sel_option(self, signal_data):
        """Return selected option based on signal_data"""

        optSymbol = 'NIFTY' if signal_data['Symbol'] == 'NSENIFTY' else signal_data['Symbol']
        optType = "CE" if signal_data['Signal'] == "CE Buy" else "PE"

        # print('TradeTest.sel_option', trade['Date'], trade['Close'], optType)

        optiondata = self.oh.nth_opt(self.nth_expiry, self.nth_strike, self.midmonth_cutoff, self.entrytime,
            symbol=r"'{}'".format(optSymbol),
            type=r"'{}'".format(optType),
            date=r"'{}'".format(signal_data['Date']),
            close=signal_data['Close']
        )

        # print('TradeTest.sel_option', optiontrade)

        return optiondata


    def option_exit_data(self, i, option_entry):
        """Return option data on exit date based on option_entry"""

        exit_date = self.dataHist.bar(i)['Date']
        exit_data = self.oh.option_exit_data(option_entry, exit_date)

        return exit_data


    def backtest(self, start=0, end=0, latitude=0, exittarget=2, nth_expiry=0, nth_strike=0, midmonth_cutoff=10,
                 entrytime='sod'):
        """Return historical trades
        Return format: {entry date in YYYY-MM-DD: {"entry": {option data from sel_option()},
                                                   "exit": {option data from option_exit_data()}
                                 },
                        entry date in YYYY-MM-DD: {"entry": {option data from sel_option()},
                                                   "exit": {option data from option_exit_data()}
                                 },......
                        }
        Call format: backtest(start=0,              start date for backtest in YYYY-MM-DD
                              end=0,                end date for backtest in YYYY-MM-DD
                              latitude=0,           additional percentage difference required for buy/sell signals
                              exittarget=2,         exit target in multiple of entry price
                              nth_expiry=0,         option expiry month to be selected
                              nth_strike=0,         option strike price to be selected
                              midmonth_cutoff=10    cut-off date to be used to select immediate expiry or next
                              entrytime='sod'       entry on same end of day ('eod') or next day start of day ('sod')
                            )
        """

        if start == 0 or start < self.dataHist.value['Date'][0]:
            start = self.dataHist.value['Date'][0]
        if end == 0 or end > self.dataHist.value['Date'][self.dataHist.barcount - 1]:
            end = self.dataHist.value['Date'][self.dataHist.barcount - 1]

        print('Running backtest from', start, 'to', end, '. . . .')
        #print('###', self.dataHist.date_bar(start), self.dataHist.barindex[start],
        #      self.dataHist.date_bar(end), self.dataHist.barindex[end])

        self.nth_expiry, self.nth_strike, self.midmonth_cutoff, self.entrytime = \
            nth_expiry, nth_strike, midmonth_cutoff, entrytime

        buy = np.logical_and(self.dataHist.value['BarUpDown'] == 1,
                              self.dataHist.value['Close'] > self.dataHist.value['LH Value'] * (1 + latitude))
        sell = np.logical_and(self.dataHist.value['BarUpDown'] == -1,
                              self.dataHist.value['Close'] < self.dataHist.value['HL Value'] * (1 - latitude))

        trades = {}
        CE_bought, CE_sold, PE_bought, PE_sold = False, False, False, False
        selected_option = None
        CE_entry_date, PE_entry_date = None, None

        for i in range(self.dataHist.barindex[start], self.dataHist.barindex[end]):
            # print(i, buy[i], sell[i])
            if buy[i]:
                if PE_bought:
                    exit_option = self.option_exit_data(i, trades[PE_entry_date]["entry"]) # Exit PE Option
                    trades[PE_entry_date]['exit'] = exit_option
                    PE_bought, PE_sold = False, True
                if not CE_bought:
                    selected_option = self.sel_option(self.signal_data(i, "CE Buy")) # Entry CE Option
                    CE_entry_date = selected_option['date']
                    trades[CE_entry_date] = {}
                    trades[CE_entry_date]['entry'] = selected_option
                    CE_bought, CE_sold = True, False
            elif sell[i]:
                if CE_bought:
                    exit_option = self.option_exit_data(i, trades[CE_entry_date]["entry"]) # Exit CE Option
                    trades[CE_entry_date]['exit'] = exit_option
                    CE_bought, CE_sold = False, True
                if not PE_bought:
                    selected_option = self.sel_option(self.signal_data(i, "PE Buy")) # Entry PE Option
                    PE_entry_date = selected_option['date']
                    trades[PE_entry_date] = {}
                    trades[PE_entry_date]['entry'] = selected_option
                    PE_bought, PE_sold = True, False
            else:
                pass


        return trades



class DataHist:
    """Class to access historical data"""

    def __init__(self, dataFile):
        """Populate all columns and bars"""

        self.fileName = dataFile
        self.wb = xl.load_workbook(self.fileName)

        # Read column headers

        self.wsc = self.wb.get_sheet_by_name('columns')  # columns tab

        self.header = {}  # format: {'column name': column number, ..... ]

        col = 1
        while self.wsc.cell(row=1, column=col).value is not None:
            # self.header.append([self.wsc.cell(row=1, column=col).value, col-1])
            self.header[self.wsc.cell(row=1, column=col).value] = col - 1
            col += 1

        # Populate all data

        self.wsd = self.wb.get_sheet_by_name('data') # data tab

        self.value = {}  # format: {'column name' : [row1_val, row2_val....], ......}

        for name, position in self.header.items():
            self.value[name] = self.xl_column(position)

        self.barindex = {}

        if 'Date' in self.header:
            for i in range(0, len(self.value['Date'])):
                self.barindex[self.value['Date'][i]] = i


        self.barcount = len(self.xl_column(0)) # number of bars in data tab


    def xl_column(self, col):
        """Return numpy for each column"""

        return np.array([r[col].value for r in self.wsd.iter_rows()])


    def bar(self, i):
        """ Return all columns in a single bar
        Return Format: {'column name' : value, .....}
        Call Format: bar(bar_index)
        """

        barvalues = {}

        for item in self.value:
            barvalues[item] = self.value[item][i]

        return barvalues

    def date_bar(self, date):
        """ Return all columns in a single bar
        Return Format: {'column name' : value, .....}
        Call Format: date_bar(date in YYYY-MM-DD)
        """
        return self.bar(self.barindex[date])

    def add_column_names(self, *columns):
        """Add new column names in columns tab"""

        next_column = len(self.header) + 1

        for column in columns:
            self.wsc.cell(row=1, column=next_column).value = column
            self.header[column] = next_column - 1
            next_column += 1

        self.wb.save(self.fileName)


    def add_data_by_bar(self, bar, data):

        for column, value in data.items():
            self.wsd.cell(row=bar, column=self.header[column] + 1).value = value

        self.wb.save(self.fileName)

    def add_data_by_date(self, date, data):

        for column, value in data.items():
            self.wsd.cell(row=self.barindex[date] + 1, column=self.header[column] + 1).value = value

        self.wb.save(self.fileName)


