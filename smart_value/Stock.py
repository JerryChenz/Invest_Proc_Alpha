import pandas as pd

def __init__(self, symbol):
    """
    :param symbol: string ticker of the stock
    """
    super().__init__(symbol)

    self.shares = None
    self.price = None
    self.valuation = None
