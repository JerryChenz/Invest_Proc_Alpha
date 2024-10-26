from yfinance import Ticker, download
import datetime as dt
from smart_value.stock import *
import time

'''
Available yfinance features:
attrs = [
    'info', 'financials', 'quarterly_financials', 'major_holders',
    'institutional_holders', 'balance_sheet', 'quarterly_balance_sheet',
    'cashflow', 'quarterly_cashflow', 'earnings', 'quarterly_earnings',
    'sustainability', 'recommendations', 'calendar'
]
'''


def get_forex(report_symbol, price_symbol):
    """ Return the forex quote
    Known bug: If there is no quote available, the fxRate will be 0.

    :param report_symbol: String forex report symbol
    :param price_symbol: String forex price symbol
    :return: Float average of last 3 day's forex quote
    """

    start_date = dt.datetime.today() - dt.timedelta(days=7)
    end_date = dt.datetime.today()
    # print(report_symbol + " " + price_symbol)
    if report_symbol == price_symbol:
        return 1
    elif report_symbol == "MOP" and price_symbol == "HKD":
        # Known missing quote on yahoo finance
        return 0.98
    else:
        forex_code = f"{report_symbol}{price_symbol}=X"
        # print(forex_code)
        # return the average of the last 3 forex quote
        return download(forex_code, start_date, end_date).tail(3)['Adj Close'].mean()


def get_quote(symbol, option):
    """ Get the updated quote from yfinance fast_info

    :param symbol: string stock symbol
    :param option: fast_info argument
    Below available
    'currency', 'dayHigh', 'dayLow', 'exchange', 'fiftyDayAverage', 'lastPrice', 'lastVolume',
    'marketCap', 'open', 'previousClose', 'quoteType', 'regularMarketPreviousClose', 'shares', 'tenDayAverageVolume',
    'threeMonthAverageVolume', 'timezone', 'twoHundredDayAverage', 'yearChange', 'yearHigh', 'yearLow'
    :return: quote given option
    """

    result = None

    tries = 2
    for i in range(tries):
        try:
            company = Ticker(symbol).fast_info
            result = company[option]
        except:  # random API error
            if i < tries - 1:  # i is zero indexed
                wait = 60
                print(f"Possible API error, wait {wait} seconds before retry...")
                time.sleep(wait)
                continue
            else:
                return result
        else:
            return result


def get_info(symbol, option):
    """ Get the updated quote from yfinance info

    :param symbol: string stock symbol
    :param option: info argument
    Below available
    'sector', 'fullTimeEmployees', 'longBusinessSummary', 'city', 'phone', 'country', 'companyOfficers', 'website',
    'maxAge', 'address1', 'industry', 'address2', 'ebitdaMargins', 'profitMargins', 'grossMargins', 'operatingCashflow',
     'revenueGrowth', 'operatingMargins', 'ebitda', 'targetLowPrice', 'recommendationKey', 'grossProfits',
     'freeCashflow', 'targetMedianPrice', 'earningsGrowth', 'currentRatio', 'returnOnAssets',
     'numberOfAnalystOpinions', 'targetMeanPrice', 'debtToEquity', 'returnOnEquity', 'targetHighPrice', 'totalCash',
     'totalDebt', 'totalRevenue', 'totalCashPerShare', 'financialCurrency', 'revenuePerShare', 'quickRatio',
     'recommendationMean', 'shortName', 'longName', 'isEsgPopulated', 'gmtOffSetMilliseconds', 'messageBoardId',
     'market', 'annualHoldingsTurnover', 'enterpriseToRevenue', 'beta3Year', 'enterpriseToEbitda',
     'morningStarRiskRating', 'forwardEps', 'revenueQuarterlyGrowth', 'sharesOutstanding', 'fundInceptionDate',
     'annualReportExpenseRatio', 'totalAssets', 'bookValue', 'sharesShort', 'sharesPercentSharesOut', 'fundFamily',
     'lastFiscalYearEnd', 'heldPercentInstitutions', 'netIncomeToCommon', 'trailingEps', 'lastDividendValue',
     'SandP52WeekChange', 'priceToBook', 'heldPercentInsiders', 'nextFiscalYearEnd', 'yield', 'mostRecentQuarter',
     'shortRatio', 'sharesShortPreviousMonthDate', 'floatShares', 'beta', 'enterpriseValue', 'priceHint',
     'threeYearAverageReturn', 'lastSplitDate', 'lastSplitFactor', 'legalType', 'lastDividendDate',
     'morningStarOverallRating', 'earningsQuarterlyGrowth', 'priceToSalesTrailing12Months', 'dateShortInterest',
     'pegRatio', 'ytdReturn', 'forwardPE', 'lastCapGain', 'shortPercentOfFloat', 'sharesShortPriorMonth',
     'impliedSharesOutstanding', 'category', 'fiveYearAverageReturn', 'trailingAnnualDividendYield', 'payoutRatio',
     'navPrice', 'trailingAnnualDividendRate', 'toCurrency', 'expireDate', 'algorithm', 'dividendRate',
     'exDividendDate', 'circulatingSupply', 'startDate', 'trailingPE', 'lastMarket', 'maxSupply', 'openInterest',
     'volumeAllCurrencies', 'strikePrice', 'ask', 'askSize', 'fromCurrency', 'fiveYearAvgDividendYield', 'bid',
     'tradeable', 'dividendYield', 'bidSize', 'coinMarketCapLink', 'preMarketPrice', 'logo_url', 'trailingPegRatio']

    :return: quote given option
    """

    company = Ticker(symbol).info
    return company[option]


def update_data(data):
    """Update the price and currency info"""

    ticker = data.index.values.tolist()[0]
    data['price'] = get_quote(ticker, 'last_price')
    data['priceCurrency'] = get_quote(ticker, 'currency')
    # Forex
    data['fxRate'] = get_forex(data['financialCurrency'].values[:1][0], data['priceCurrency'].values[:1][0])
    # exclude minor countries without forex data
    # data = data[data['fxRate'].notna()]
    return data


class YfData(Stock):
    """Retrieves the data from yfinance"""

    def __init__(self, symbol):
        """
        :param symbol: string ticker of the stock
        """
        super().__init__(symbol)
        # yfinance
        try:
            self.stock_data = Ticker(self.symbol)
        except KeyError:
            print("Check your stock ticker")
        # self.avg_gross_margin = None
        # self.avg_ebit_margin = None
        # self.avg_net_margin = None
        # self.avg_sales_growth = None
        # self.avg_ebit_growth = None
        # self.avg_ni_growth = None
        # self.years_of_data = None
        self.load_attributes()

    def load_attributes(self):

        self.name = self.stock_data.info['shortName']
        self.price = [self.stock_data.fast_info['last_price'], self.stock_data.fast_info['currency']]
