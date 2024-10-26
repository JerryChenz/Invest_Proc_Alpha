from yahooquery import Ticker
import time


def get_quote(symbol):
    """ Get the updated quote from yfinance fast_info

    :param symbol: string stock symbol
    :return: quote given option
    """

    price = None
    price_currency = None

    tries = 2
    for i in range(tries):
        try:
            company = Ticker(symbol)

            price = company.financial_data[symbol]['currentPrice']
            price_currency = company.price[symbol]['currency']
        except:  # random API error
            if i < tries - 1:  # i is zero indexed
                wait = 60
                print(f"Possible API error, wait {wait} seconds before retry...")
                time.sleep(wait)
                continue
            else:
                return price, price_currency
        else:
            return price, price_currency
