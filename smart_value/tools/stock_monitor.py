import xlwings
import re
from smart_value.tools import *
from smart_value.tools.find_docs import models_folder_path
from smart_value.tools.find_docs import macro_monitor_file_path
from smart_value.tools.find_docs import stock_monitor_file_path
from smart_value.tools.model import model_pos


def update_monitor(quick=False):
    """Update the Monitor file

    :param quick: the option to update the monitor sheet without updating the valuations
    """

    opportunities = []

    # Step 1: Update the marco Monitor
    macro_monitor.update_marco(macro_monitor_file_path, "Free")

    # Step 2: load and update the new valuation xlsx
    for opportunities_path in find_docs.get_model_paths():
        print(f"Working with {opportunities_path}...")
        op = read_opportunity(opportunities_path, quick)  # load and update the new valuation xlsx
        opportunities.append(op)

    # Step 3: Update the stock monitor.
    print("Updating Monitor...")
    with xlwings.App(visible=False) as app:
        s_monitor = app.books.open(stock_monitor_file_path)
        update_opportunities(s_monitor, opportunities)
        s_monitor.save(models_folder_path)
        s_monitor.close()
    print("Update completed")


def read_opportunity(opportunities_path, quick=False):
    """Read all the opportunities at the opportunities_path.

    :param opportunities_path: path of the model in the opportunities' folder
    :param quick: the option to skip updating the models
    :return: an Asset object
    """

    r_stock = re.compile(".*_Stock_Valuation")
    # get the formula results using xlwings because openpyxl doesn't evaluate formula
    with xlwings.App(visible=False) as app:
        xl_book = app.books.open(opportunities_path)
        dash_sheet = xl_book.sheets('Dashboard')
        if r_stock.match(str(opportunities_path)):
            company = model.StockModel(dash_sheet.range(model_pos["symbol"]).value, "yq_quote")
            if quick is False:
                model.update_dashboard(dash_sheet, company)  # Update
            xl_book.save(opportunities_path)  # xls must be saved to update the values
            op = MonitorStock(dash_sheet)  # the MonitorStock object representing an opportunity
        else:
            print(f"'{opportunities_path}' is incorrect")
        xl_book.close()

    return op


def update_opportunities(s_monitor, op_list):
    """Update the opportunities sheet in the stock_monitor file

    :param op_list: list of stock objects
    :param s_monitor: xlwings stock monitor file object
    """

    monitor_sheet = s_monitor.sheets('Opportunities')
    monitor_sheet.range('B5:S400').clear_contents()

    r = 5
    for op in op_list:
        monitor_sheet.range((r, 2)).value = op.symbol
        monitor_sheet.range((r, 3)).value = op.name
        # Price section
        monitor_sheet.range((r, 4)).value = op.price
        monitor_sheet.range((r, 5)).value = op.price_currency
        monitor_sheet.range((r, 6)).value = op.ep_ratio
        monitor_sheet.range((r, 7)).value = op.dp_ratio
        # Valuation section
        monitor_sheet.range((r, 8)).value = op.value
        monitor_sheet.range((r, 9)).value = op.lower_value
        monitor_sheet.range((r, 10)).value = op.upper_value
        # Miscellaneous section
        monitor_sheet.range((r, 11)).value = op.watchlist
        monitor_sheet.range((r, 12)).value = op.comp_group
        monitor_sheet.range((r, 13)).value = op.last_revision
        monitor_sheet.range((r, 14)).value = op.next_review
        # Cost Structure section
        monitor_sheet.range((r, 15)).value = op.cogs
        monitor_sheet.range((r, 16)).value = op.op_exp_less_da
        monitor_sheet.range((r, 17)).value = op.interest
        monitor_sheet.range((r, 18)).value = op.change_of_wc
        monitor_sheet.range((r, 19)).value = op.non_controlling_interests
        monitor_sheet.range((r, 20)).value = op.pre_tax_profit
        r += 1
    print(f"Total {len(op_list)} opportunities Updated")


class MonitorStock:
    """Monitor class

    Defines what data can be extracted from the valuation model and used in the Monitor.
    """

    def __init__(self, dash_sheet):
        self.symbol = dash_sheet.range(model_pos["symbol"]).value
        self.name = dash_sheet.range(model_pos["name"]).value
        # Price section
        self.price = dash_sheet.range(model_pos["price"]).value
        self.price_currency = dash_sheet.range(model_pos["price_currency"]).value
        self.ep_ratio = dash_sheet.range(model_pos["ep_ratio"]).value
        self.dp_ratio = dash_sheet.range(model_pos["dp_ratio"]).value
        # Valuation section
        self.value = dash_sheet.range(model_pos["value"]).value
        self.lower_value = dash_sheet.range(model_pos["lower_value"]).value
        self.upper_value = dash_sheet.range(model_pos["upper_value"]).value
        # Miscellaneous section
        self.watchlist = dash_sheet.range(model_pos["watchlist"]).value
        self.comp_group = dash_sheet.range(model_pos["comp_group"]).value
        self.last_revision = dash_sheet.range(model_pos["last_revision"]).value
        self.next_review = dash_sheet.range(model_pos["next_review"]).value
        # Cost Structure section
        self.cogs = dash_sheet.range(model_pos["cogs"]).value
        self.op_exp_less_da = dash_sheet.range(model_pos["op_exp_less_da"]).value
        self.interest = dash_sheet.range(model_pos["interest"]).value
        self.change_of_wc = dash_sheet.range(model_pos["change_of_wc"]).value
        self.non_controlling_interests = dash_sheet.range(model_pos["non_controlling_interests"]).value
        self.pre_tax_profit = dash_sheet.range(model_pos["pre_tax_profit"]).value
