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

        self.wb = xl.load_workbook(self.testFile)

        self.dataHist = DataHist(self.wb)

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
            # print(i, buy[i], sell[i])
            if buy[i]:  # and buy[i] != prevBuy:
                if not bought:
                    signals.append("Buy")
                    bought, sold = True, False
                else:
                    signals.append(None)
            elif sell[i]:  # and sell[i] != prevSell:
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
                # print(trades[len(trades)-1]['Date'], trades[len(trades)-1]['Signal'])

        # print(len(trades))
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
            # print(optTrades[len(optTrades)-1])


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

        exit_date = self.dataHist.bar(i)['Date']
        exit_data = self.oh.option_exit_data(option_entry, exit_date)

        return exit_data


    def backtest(self, start=0, end=0, latitude=0, exittarget=2, nth_expiry=0, nth_strike=0, midmonth_cutoff=10,
                 entrytime='sod'):

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

        option_entries = []
        CE_bought, CE_sold, PE_bought, PE_sold = False, False, False, False
        selected_option, option_expiry = None, None

        for i in range(self.dataHist.barindex[start], self.dataHist.barindex[end]):
            # print(i, buy[i], sell[i])
            if buy[i]:
                if PE_bought:
                    exit_option = self.option_exit_data(i, option_entries[len(option_entries) - 1]) # Exit PE Option
                    # print('Tradetest.backtest exit option', exit_option)
                    option_entries.append(exit_option)
                    PE_bought, PE_sold = False, True
                    print('Tradetest.backtest PE Exit', self.dataHist.bar(i)['Date'], exit_option['type'],
                          exit_option['expiry'])
                if not CE_bought:
                    selected_option = self.sel_option(self.signal_data(i, "CE Buy"))
                    selected_option['signal'] = 'CE Entry'
                    # print('Tradetest.backtest selected option', selected_option)
                    option_entries.append(selected_option)
                    option_expiry = selected_option['expiry']
                    CE_bought, CE_sold = True, False
                    print('Tradetest.backtest CE Entry', self.dataHist.bar(i)['Date'], selected_option['type'],
                          selected_option['expiry'])
            elif sell[i]:
                if CE_bought:
                    exit_option = self.option_exit_data(i, option_entries[len(option_entries) - 1]) # Exit CE Option
                    # print('Tradetest.backtest exit option', exit_option)
                    option_entries.append(exit_option)
                    CE_bought, CE_sold = False, True
                    print('Tradetest.backtest CE Exit', self.dataHist.bar(i)['Date'], exit_option['type'],
                          exit_option['expiry'])
                if not PE_bought:
                    selected_option = self.sel_option(self.signal_data(i, "PE Buy"))
                    selected_option['signal'] = 'PE Entry'
                    # print('Tradetest.backtest selected option', selected_option)
                    option_entries.append(selected_option)
                    option_expiry = selected_option['expiry']
                    PE_bought, PE_sold = True, False
                    print('TradeTest.backtest PE Entry', self.dataHist.bar(i)['Date'], selected_option['type'],
                          selected_option['expiry'])
            else:
                pass
            # print('TradeTest.backtest', self.dataHist.bar(i)['Date'], bought, sold, option_entries[len(option_entries) - 1]['expiry'])
            # if (bought or sold) and self.dataHist.bar(i)['Date'] == option_entries[len(option_entries) - 1]['expiry']:
            #     print('Tradetest.backtest Exit', self.dataHist.bar(i)['Date'])


        return option_entries



class DataHist:
    """Class to access historical data"""

    def __init__(self, wb):
        """Populate all columns and bars"""

        # Read column headers

        wsc = wb.get_sheet_by_name('columns')  # columns tab

        self.header = []  # format: [['column name', column number], ..... ]

        col = 1
        while wsc.cell(row=1, column=col).value is not None:
            self.header.append([wsc.cell(row=1, column=col).value, col-1])
            col += 1

        # Populate all data

        wsd = wb.get_sheet_by_name('data') # data tab

        self.value = {}  # format: {'column name' : [row1_val, row2_val....], ......}

        for item in self.header:
            self.value[item[0]] = self.xl_column(wsd, item[1])

        self.barindex = {}

        if 'Date' in [column[0] for column in self.header]:
            for i in range(0, len(self.value['Date'])):
                self.barindex[self.value['Date'][i]] = i


        self.barcount = len(self.xl_column(wsd, 0)) # number of bars in data tab

        # print(self.data['DATE'][len(self.data['DATE'])-1])


    def xl_column(self, wsd, col):
        """Return numpy for each column"""

        return np.array([r[col].value for r in wsd.iter_rows()])


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












