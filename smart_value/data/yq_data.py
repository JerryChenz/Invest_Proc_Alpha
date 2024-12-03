from yahooquery import Ticker
import time


def get_quotes(ticker_list):
    """ Get the updated stock info from yahooquery

    :param ticker_list: list of stock symbols
    :return: a dictionary containing stock data in the last
    """

    stock_dict = {}

    print(f"obtaining the prices for {len(ticker_list)} stocks...")
    tries = 2
    for i in range(tries):
        try:
            all_symbols = " ".join(ticker_list)
            stock_info = Ticker(all_symbols)
            stock_dict = stock_info.price
        except:  # random API error
            if i < tries - 1:  # i is zero indexed
                wait = 60
                print(f"Possible API error, wait {wait} seconds before retry...")
                time.sleep(wait)
                continue
            else:
                return stock_dict
        else:
            return stock_dict


def get_price(ticker, stock_dict):
    """ Get the updated stock info from yahooquery

    :param ticker: string stock symbol
    :param stock_dict: a dictionary containing stock data in the last
    :return: the stock price of this symbol
    """

    return stock_dict[ticker]['regularMarketPrice'], stock_dict[ticker]['currency']

