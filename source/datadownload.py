"""
Created on Jan 28, 2017
@author: Souvik
@Program Function: Download NSENIFTY Options data


"""

from nsepy import get_history, get_index_pe_history
import os
import openpyxl as xl
from trade import DataHist
from datetime import datetime, date
from dateutil.relativedelta import relativedelta

root = "C:/Users/Souvik/OneDrive/Python/Trade1"      # Laptop
#root = "C:/Users/SVK/OneDrive/Python/Trade1"        # Desktop
NIFTY = "data/NIFTY_All 2005 onwards.xlsx"        # NIFTY Data
expFile = "data/NIFTY_OPT_Expiries - 2004 onwards.xlsx"             # expiry dates, should start with first entry's month as per NIFTY


#  optionList Data Format:
#  {
#   expiry1: {'CE': ([symbol, startdate, enddate, strike1], [..., strike2], ...[]),
#             'PE': ([symbol, startdate, enddate, strike1], [..., strike2], ...[])},
#   expiry2: {'CE': ([symbol, startdate, enddate, strike1], [..., strike2], ...[]),
#             'PE': ([symbol, startdate, enddate, strike1], [..., strike2], ...[])},
#   expiry3: {'CE': (),
#             'PE': ()}
#  }

# optionList['2016-12-28'] = {'CE': [['NIFTY', '2016-10-01', '2016-12-28', 9000],
#                                  ['NIFTY', '2016-10-01', '2016-12-28', 9100]],
#                            'PE': [['NIFTY', '2016-10-01', '2016-12-28', 9000],
#                                   ['NIFTY', '2016-10-01', '2016-12-28', 9100]]
#                            }

optionList = {}


def expiry_date_ranges(expiryList):
    """Return date ranges for option data for set of expiries
    Input format: [expiry1, expiry2,....] i.e., [YYYY-MM-DD, YYYY-MM-DD,.....]
    Return Format: {expmonth1 : [start date, end date], expmonth2 : [],....}
                   i.e., {YYYY-MM : [YYYY-MM-DD, YYYY-MM-DD], YYYY-MM : [],....}
    """

    expiryDateRanges = {}

    for expiry in expiryList:
        expiryDate = datetime.strptime(expiry, '%Y-%m-%d').date()
        start = expiryDate.replace(day=1)
        start = start + relativedelta(months=-3)
        expiryDateRanges[expiryDate.strftime('%Y-%m')] = [start.strftime('%Y-%m-%d'), expiryDate.strftime('%Y-%m-%d')]
        #print("$$$", expiryDate, expiryDateRanges[expiryDate])

    #print('$$$',type(date(2006,1,10)))

    return expiryDateRanges


def expiry_months(date):
    """Return 3 expiry months for a given date
    Input format: date i.e., YYYY-MM-DD
    Return Format: [expmonth1, expmonth2, expmonth3]
                   i.e., [YYYY-MM-DD, YYYY-MM-DD, YYYY-MM-DD]
    """

    barDate = datetime.strptime(date, '%Y-%m-%d').date()

    return [barDate.strftime('%Y-%m'), (barDate + relativedelta(months=1)).strftime('%Y-%m'),
           (barDate + relativedelta(months=2)).strftime('%Y-%m')]


def cepe_options(date, close, expMonths, expiryDateRanges):
    """Updates optionList"""

    fs = close - close % 100 + 100

    strikesCE = [fs, fs+100, fs+200, fs+300, fs+400, fs+500, fs+600, fs+700, fs+800, fs+900]
    strikesPE = [fs-100, fs-200, fs-300, fs-400, fs-500, fs-600, fs-700, fs-800, fs-900, fs-1000]

    # print(close, strikesCE, strikesPE)

    #options = {}
    for m in expMonths:
        #options[m] = []
        for s in strikesCE:
            new = [expiryDateRanges[m][0], expiryDateRanges[m][1], s]
            if new not in optionList[m]['CE']:
                optionList[m]['CE'].append(new)

        for s in strikesPE:
            new = [expiryDateRanges[m][0], expiryDateRanges[m][1], s]
            if new not in optionList[m]['PE']:
                optionList[m]['PE'].append(new)

        #print(date, close, m, options[m])
        #print("$$$", m, optionList[m])

########################################################################################################################


os.chdir(root)

wbExp = xl.load_workbook(expFile)
wbData = xl.load_workbook(NIFTY)
expHist = DataHist(wbExp)
dataHist = DataHist(wbData)

# set expiry date ranges based on expiry dates in expFile

expiryDateRanges = expiry_date_ranges(expHist.value['expiry'])

for item, dates in expiryDateRanges.items():
    print(item, dates[0], dates[1])
#print(expiryDateRanges)
#print(dataHist.value['Close'][0])

for item, val in expiryDateRanges.items():
    optionList[item] = {'CE': [], 'PE': []}
# print(optionList)


for i in range(0, dataHist.barcount):
    expMonths = expiry_months(dataHist.value['Date'][i])
    # print(dataHist.value['Date'][i], expMonths)
    cepe_options(dataHist.value['Date'][i], int(dataHist.value['Close'][i]), expMonths, expiryDateRanges)

"""
optCount = 0
#print(optionList)
for expiry, val in optionList.items():
    for otype, data in val.items():
        for opt in data:
            extract = get_history(symbol="NIFTY",
                                  start=datetime.strptime(opt[0], '%Y-%m-%d').date(),
                                  end=datetime.strptime(opt[1], '%Y-%m-%d').date(),
                                  index=True,
                                  option_type=otype,
                                  strike_price=opt[2],
                                  expiry_date=datetime.strptime(opt[1], '%Y-%m-%d').date())
            if len(extract) > 0:
                filename = 'data/extracts/{}_{}_{}.csv'.format(expiry, otype, opt[2])
                extract.to_csv(filename, sep=',')
                print(expiry, otype, opt, len(extract), "{} written".format(filename))
            else:
                print(expiry, otype, opt, len(extract), "No records")
"""
