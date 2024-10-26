class Stock:
    """a type of Assets"""

    def __init__(self, symbol):
        """
        :param symbol: string ticker of the stock
        """

        self.symbol = symbol
        self.name = None
        self.price = None
        self.price_currency = None
        self.shares_outstanding = None
