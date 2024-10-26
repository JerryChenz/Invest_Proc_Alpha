import xlwings
import pathlib
import shutil
import os
from datetime import datetime
import re

import smart_value.tools.marco_monitor
from smart_value.data import yf_data as yf
from smart_value.data import yq_data as yq
from smart_value.stock import Stock


def update_model(ticker, model_name, model_path, source):
    """Update the model.

    :param ticker: the string ticker of the stock
    :param model_name: the model file name
    :param model_path: the model file path
    :param source: String data source selector
    """

    print(f'Updating {model_name}...')
    company = StockModel(ticker, source)  # uses yahoo finance data by default

    with xlwings.App(visible=False) as app:
        model_xl = app.books.open(model_path)
        update_dashboard(model_xl.sheets('Dashboard'), company)
        model_xl.save(model_path)
        model_xl.close()


def update_dashboard(dash_sheet, stock):
    """Update the Dashboard sheet.

    :param dash_sheet: the xlwings object of the dashboard sheet in the model
    :param stock: the Stock object
    """

    marco = smart_value.tools.marco_monitor.read_marco()
    dash_sheet.range('C10').value = marco.us_riskfree

    dash_sheet.range('I4').value = stock.price[0]
    dash_sheet.range('I12').value = stock.fx_rate


class StockModel(Stock):
    """Stock model class"""

    def __init__(self, symbol, source):
        """
        :param symbol: string ticker of the stock
        :param source: data source selector
        """
        super().__init__(symbol)
        self.source = source
        self.load_attributes()

    def load_attributes(self):
        """data source selector."""
        try:
            if self.source == "yf":
                ticker_data = yf.YfData(self.symbol)
                self.load_data(ticker_data)
                self.fx_rate = yf.get_forex(self.report_currency, self.price[1])

            elif self.source == "yf_quote":
                market_price = yf.get_quote(self.symbol, "last_price")
                price_currency = yf.get_quote(self.symbol, "currency")
                report_currency = yf.get_info(self.symbol, "financialCurrency")
                self.load_quote(market_price, price_currency, report_currency)
                self.fx_rate = yf.get_forex(report_currency, price_currency)

            elif self.source == "yq":
                ticker_data = yq.YqData(self.symbol)
                self.load_data(ticker_data)
                self.fx_rate = yf.get_forex(self.report_currency, self.price[1])  # Use the better yfinance Forex

            elif self.source == "yq_quote":
                quote = yq.get_quote(self.symbol)
                market_price = quote[0]
                price_currency = quote[1]
                report_currency = quote[2]
                self.load_quote(market_price, price_currency, report_currency)
                self.fx_rate = yf.get_forex(report_currency, price_currency)  # Use the better yfinance Forex

            elif self.source == "fmp":
                pass

            else:
                raise KeyError(f"The source keyword {self.source} is invalid!")
        except KeyError as error:
            print('Caught this error: ' + repr(error))

    def load_data(self, ticker_data):
        """Scrap the financial_data from yahoo finance

        :param ticker_data: financial data object
        """

        self.name = ticker_data.name
        self.sector = ticker_data.sector
        self.price = ticker_data.price
        self.exchange = ticker_data.exchange
        self.shares = ticker_data.shares
        self.report_currency = ticker_data.report_currency
        self.last_dividend = ticker_data.last_dividend
        self.buyback = ticker_data.buyback
        self.annual_bs = ticker_data.annual_bs
        self.quarter_bs = ticker_data.quarter_bs
        self.cf_df = ticker_data.cf_df
        self.is_df = ticker_data.is_df
        self.last_fy = ticker_data.annual_bs.columns[0]
        self.most_recent_quarter = ticker_data.most_recent_quarter

    def load_quote(self, market_price, price_currency, report_currency):
        """Get the quick market quote

        :param market_price: current or delayed market price
        :param price_currency: currency symbol of the market_price
        :param report_currency: currency symbol of the financial statements
        """

        self.price = [market_price, price_currency]
        self.fx_rate = yf.get_forex(report_currency, price_currency)
