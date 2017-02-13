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

"""
oh = opthist.OptionsHist(optionsDB)

oh.immediate_opt(r"'NIFTY'", optiontype=r"'CE'", datevalue=r"'2007-09-02'", close=4400.15)
oh.immediate_opt(r"'NIFTY'", optiontype=r"'CE'", datevalue=r"'2007-09-02'", close=5406.15)
oh.immediate_opt(r"'NIFTY'", optiontype=r"'PE'", datevalue=r"'2007-02-02'", close=4086.15)
selOpt = oh.immediate_opt(r"'NIFTY'", optiontype=r"'PE'", date=r"'2007-02-02'", close=4006.15)

print(selOpt)
"""
tt = trade.TradeTest(root, dataFile, optionsDB)

#print(test.dataHist.value['DATE'][1])
#print(test.dataHist.bar(1))
s = tt.signals()
t = tt.trades(s)
# tt.sel_options(t)
# tt.backtest(start='2016-01-04',end='2016-06-14')
print(tt.backtest(start='2016-01-04',end='2016-06-14'))
tt.dataHist.add_column_names('Entry Type', 'Expiry', 'Strike Price', 'Entry Date', 'Entry Open', 'Entry High',
                             'Entry Low', 'Entry Close', 'Entry LTP', 'Exit Date', 'Exit Open', 'Exit High',
                             'Exit Low', 'Exit Close', 'Exit LTP')
print(tt.dataHist.header)
data = {'Entry Type':'PE', 'Expiry':'test', 'Strike Price':10}
tt.dataHist.add_data_by_bar(10, data)
tt.dataHist.add_data_by_date('2006-12-11', data)
#print(tt.backtest())




# print('bars : ', test.dataHist.barcount, 'signals : ', len(s))

"""
for i in range(0, test.dataHist.barcount):
    # print(i)
    # print(i, test.dataHist.value['DATE'][i], 'LH Val : ', test.dataHist.value['LH Value'][i], 'HL Val : ',
    #      test.dataHist.value['HL Value'][i], 'Close : ', test.dataHist.value['Close'][i], s[i])
    print(i, test.dataHist.value['Date'][i],s[i] )

"""