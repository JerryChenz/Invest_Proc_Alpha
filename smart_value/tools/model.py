import smart_value.tools.macro_monitor as macro_monitor
from smart_value.data import yf_data as yf
from smart_value.data import yq_data as yq
from smart_value.stock import Stock

model_pos = {
    # Left side of the company info
    "symbol": "C3",
    "name": "C4",
    "last_revision": "C5",
    "next_review": "D6",
    "watchlist": "C7",
    "comp_group": "D7",
    # Right side of the company info
    "price": "G3",
    "price_currency": "H3",
    "shares_outstanding": "G4",
    "report_currency": "G6",
    "fx_rate": "G7",
    # US market yields
    "us_riskfree": "C10",
    "us_bbb_yield": "C11",
    "us_required_return": "C12",
    # China market yields
    "cn_riskfree": "C14",
    "cn_on_bbb_yield": "C15",
    "cn_off_bbb_yield": "C16",
    "cn_required_return": "C17",
    # HK and other market yields
    "hk_required_return": "D12",
    "other_required_return": "D17",
    # Normalized Cost Structure
    "cogs": 'C20',
    "op_exp_less_da": 'C21',
    "interest": 'C22',
    "mcx": 'C23',
    "change_of_wc": 'C24',
    "non_controlling_interests": 'C25',
    "pre_tax_profit": 'C26',
    # Business Indicators
    "pre_tax_roa": 'G20',
    "after_tax_growth": 'G21',
    # Price Indicators
    "pb_ratio": 'G23',
    "ep_ratio": 'G24',
    "payout_ratio": 'G25',
    "dp_ratio": 'G26',
    # Valuation
    "lower_value": 'C29',
    "upper_value": 'D29',
    "value": 'F29'
}


def update_dashboard(dash_sheet, stock):
    """Update the Dashboard sheet.

    :param dash_sheet: the xlwings object of the dashboard sheet in the model
    :param stock: the Stock object
    """

    marco = macro_monitor.read_marco()
    dash_sheet.range(model_pos["us_riskfree"]).value = marco.us_riskfree
    dash_sheet.range(model_pos["us_bbb_yield"]).value = marco.us_bbb_yield
    dash_sheet.range(model_pos["us_required_return"]).value = marco.us_required_return
    dash_sheet.range(model_pos["cn_riskfree"]).value = marco.cn_riskfree
    dash_sheet.range(model_pos["cn_on_bbb_yield"]).value = marco.cn_on_bbb_yield
    dash_sheet.range(model_pos["cn_off_bbb_yield"]).value = marco.cn_off_bbb_yield
    dash_sheet.range(model_pos["cn_required_return"]).value = marco.cn_required_return
    dash_sheet.range(model_pos["hk_required_return"]).value = marco.hk_required_return
    dash_sheet.range(model_pos["other_required_return"]).value = marco.other_required_return
    # tbc on marco update
    dash_sheet.range(model_pos["price"]).value = stock.price[0]
    dash_sheet.range(model_pos["fx_rate"]).value = stock.fx_rate


class StockModel(Stock):
    """Stock model class"""

    def __init__(self, symbol, report_currency, source):
        """
        :param symbol: string ticker of the stock
        :param source: data source selector
        """
        super().__init__(symbol)
        self.report_currency = report_currency
        self.fx_rate = None
        self.source = source
        self.load_attributes()

    def load_attributes(self):
        """data source selector."""
        try:
            if self.source == "yf":
                market_price = yf.get_quote(self.symbol, "last_price")
                price_currency = yf.get_quote(self.symbol, "currency")
                self.price = [market_price, price_currency]
                self.fx_rate = yf.get_forex(self.report_currency, price_currency)

            else:
                quote = yq.get_quote(self.symbol)
                market_price = quote[0]
                price_currency = quote[1]
                self.price = [market_price, price_currency]
                self.fx_rate = yf.get_forex(self.report_currency, price_currency)  # Use the better yfinance Forex

        except KeyError as error:
            raise KeyError(f"The source keyword {self.source} is invalid!")
