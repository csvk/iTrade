"""
Created on Jan 11, 2017

@author: Souvik
"""

import os
import opthist
import trade

root = "C:/Users/Souvik/OneDrive/Python/iTrade" # Laptop
# root = "C:/Users/SVK/OneDrive/Python/iTrade"     # Desktop
dataFile = "data/NIFTY.xlsx"
optionsDB = "data/NIFTY_OPTDB_Fresh.db"


os.chdir(root)

tt = trade.TradeTest(root, dataFile, optionsDB)

s = tt.signals()
t = tt.trades(s)
# tt.sel_options(t)
# tt.backtest(start='2016-01-04',end='2016-06-14')
# print(tt.backtest(start='2007-02-01', end='2007-03-30', exit_target=2))
# print(tt.backtest(start='2008-08-14',end='2008-11-20'))
# print(tt.backtest(start='2010-02-01',end='2010-10-29'))
# print(tt.backtest())
print(tt.backtest(exit_target=2))

