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
print(tt.backtest(start='2016-01-04',end='2016-06-14'))
tt.dataHist.add_column_names('Entry Type', 'Expiry', 'Strike Price', 'Entry Date', 'Entry Open', 'Entry High',
                             'Entry Low', 'Entry Close', 'Entry LTP', 'Exit Date', 'In_Market Expiry',
                             'In_Market Strike Price', 'In_Market Open', 'In_Market High', 'In_Market Low',
                             'In_Market Close', 'In_Market LTP')
print(tt.dataHist.header)
data = {'Entry Type':'PE', 'Expiry':'test', 'Strike Price':10}
tt.dataHist.add_data_by_bar(10, data)
tt.dataHist.add_data_by_date('2006-12-11', data)




# print('bars : ', test.dataHist.barcount, 'signals : ', len(s))

"""
for i in range(0, test.dataHist.barcount):
    # print(i)
    # print(i, test.dataHist.value['DATE'][i], 'LH Val : ', test.dataHist.value['LH Value'][i], 'HL Val : ',
    #      test.dataHist.value['HL Value'][i], 'Close : ', test.dataHist.value['Close'][i], s[i])
    print(i, test.dataHist.value['Date'][i],s[i] )

"""