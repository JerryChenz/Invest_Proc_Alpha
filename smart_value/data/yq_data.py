from smart_value.stock import *
from yahooquery import Ticker
from forex_python.converter import CurrencyRates
import time


def get_quote(symbol):
    """ Get the updated quote from yfinance fast_info

    :param symbol: string stock symbol
    :return: quote given option
    """

    price = None
    price_currency = None
    report_currency = None

    tries = 2
    for i in range(tries):
        try:
            company = Ticker(symbol)

            price = company.financial_data[symbol]['currentPrice']
            price_currency = company.price[symbol]['currency']
            report_currency = company.financial_data[symbol]['financialCurrency']
        except:  # random API error
            if i < tries - 1:  # i is zero indexed
                wait = 60
                print(f"Possible API error, wait {wait} seconds before retry...")
                time.sleep(wait)
                continue
            else:
                return price, price_currency, report_currency
        else:
            return price, price_currency, report_currency


def update_data(data):
    """Update the price and currency info"""

    ticker = data.index.values.tolist()[0]
    quote = get_quote(ticker)
    data['price'] = quote[0]
    data['priceCurrency'] = quote[1]
    # Forex
    data['fxRate'] = get_forex(data['financialCurrency'].values[:1][0], data['priceCurrency'].values[:1][0])

    return data


def get_forex(buy, sell):
    """get exchange rate, buy means ask and sell means bid

    :param buy: report currency
    :param sell: price currency
    """

    try:
        if buy != sell:
            c = CurrencyRates()
            return c.get_rate(buy, sell)
        elif buy == "HKD" and sell == "MOP":
            return 0.97  # MOP to HKD not available
        else:
            return 1
    except:
        print("fx_rate error: Currency Rates Not Available")
        # will return None


class YqData(Stock):
    """Retrieves the data from yahooquery"""

    def __init__(self, symbol):
        """
        :param symbol: string ticker of the stock
        """
        super().__init__(symbol)
        try:
            self.stock_data = Ticker(self.symbol)
        except KeyError:
            print("Check your stock ticker")
        self.load_attributes()

    def load_attributes(self):

        self.name = self.stock_data.quote_type[self.symbol]['shortName']
        self.price = [self.stock_data.financial_data[self.symbol]['currentPrice'],
                      self.stock_data.price[self.symbol]['currency']]
