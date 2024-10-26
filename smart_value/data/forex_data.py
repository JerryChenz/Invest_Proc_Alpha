from forex_python.converter import CurrencyRates


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
