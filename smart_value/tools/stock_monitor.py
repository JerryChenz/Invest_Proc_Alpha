import xlwings
import re
from smart_value.tools import *
from smart_value.tools.find_docs import models_folder_path as models_folder_path
from smart_value.tools.find_docs import macro_monitor_file_path as macro_monitor_file_path
from smart_value.tools.find_docs import stock_monitor_file_path as stock_monitor_file_path


def update_monitor(quick=False):
    """Update the Monitor file

    :param quick: the option to update the monitor sheet without updating the valuations
    """

    opportunities = []

    # Step 1: Update the marco Monitor
    marco_monitor.update_marco(macro_monitor_file_path, "Free")

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
    :param quick: the option to skip updating the valuations
    :return: an Asset object
    """

    r_stock = re.compile(".*_Stock_Valuation")
    # get the formula results using xlwings because openpyxl doesn't evaluate formula
    with xlwings.App(visible=False) as app:
        xl_book = app.books.open(opportunities_path)
        dash_sheet = xl_book.sheets('Dashboard')
        asset_sheet = xl_book.sheets('Asset_Model')
        if r_stock.match(str(opportunities_path)):
            company = model.StockModel(dash_sheet.range('C3').value, "yq_quote")
            if quick is False:
                model.update_dashboard(dash_sheet, company)  # Update
            xl_book.save(opportunities_path)  # xls must be saved to update the values
            op = MonitorStock(dash_sheet, asset_sheet)  # the MonitorStock object representing an opportunity
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
        monitor_sheet.range((r, 4)).value = op.price
        monitor_sheet.range((r, 5)).value = op.price_currency
        monitor_sheet.range((r, 6)).value = op.current_excess_return
        monitor_sheet.range((r, 7)).value = op.frd_dividend
        monitor_sheet.range((r, 8)).value = f'=D{r}/I{r}-1'
        monitor_sheet.range((r, 9)).value = op.next_buy_price
        monitor_sheet.range((r, 10)).value = op.next_buy_shares
        monitor_sheet.range((r, 11)).value = op.exp_exit_price
        monitor_sheet.range((r, 12)).value = op.worst_case_price
        monitor_sheet.range((r, 13)).value = op.exp_value
        monitor_sheet.range((r, 14)).value = op.breakeven_price
        monitor_sheet.range((r, 15)).value = op.lfy_date
        monitor_sheet.range((r, 16)).value = op.next_review
        monitor_sheet.range((r, 17)).value = op.exchange
        monitor_sheet.range((r, 18)).value = op.inv_category
        r += 1
    print(f"Total {len(op_list)} opportunities Updated")


class MonitorStock:
    """Monitor class

    Defines what data can be extracted from the valuation model and used in the Monitor.
    """

    def __init__(self, dash_sheet):
        # Left side of the company info
        self.symbol = dash_sheet.range('C3').value
        self.name = dash_sheet.range('C4').value
        self.last_revision = dash_sheet.range('C5').value
        self.next_review = dash_sheet.range('D6').value
        self.watchlist = dash_sheet.range('C7').value
        # Right side of the company info
        self.price = dash_sheet.range('G3').value
        self.price_currency = dash_sheet.range('H3').value
        # US market yields
        self.us_riskfree = dash_sheet.range('C10').value
        self.us_bbb_yield = dash_sheet.range('C11').value
        self.us_required_return = dash_sheet.range('C12').value
        # China market yields
        self.cn_riskfree = dash_sheet.range('C14').value
        self.cn_bbb_yield = dash_sheet.range('C15').value
        self.cn_required_return = dash_sheet.range('C16').value
        # HK and other market yields
        self.hk_required_return = dash_sheet.range('D12').value
        self.other_required_return = dash_sheet.range('D17').value
        # Normalized Cost Structure
        self.cogs = dash_sheet.range('C20').value
        self.op_exp_less_da = dash_sheet.range('C21').value
        self.interest = dash_sheet.range('C22').value
        self.change_of_wc = dash_sheet.range('C23').value
        self.non_controlling_interests = dash_sheet.range('C24').value
        self.pre_tax_profit = dash_sheet.range('C25').value
        # Valuation
        self.lower_value = dash_sheet.range('C29').value
        self.upper_value = dash_sheet.range('D29').value
        self.value = dash_sheet.range('F29').value
