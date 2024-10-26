import xlwings
from smart_value.data.fred_data import risk_free_rate, inflation
from smart_value.data.hkma_data import get_hk_riskfree
from smart_value.tools.find_docs import macro_monitor_file_path as macro_monitor_file_path
from openpyxl import load_workbook

# Define the positions of the Marco data
us_riskfree_position = 'D6'


def update_marco(monitor_path, source):
    """Update the marco data in the Macro_Monitor.xlsx file.

    :param monitor_path: path of the Macro_Monitor.xlsx file
    :param source: the API option
    """

    print("Updating Marco data...")
    us_riskfree = 0.08
    cn_riskfree = 0.06
    hk_riskfree = us_riskfree
    us_inflation = 0.02

    if source == "Free":
        us_riskfree = risk_free_rate("us")
        # print(us_riskfree)
        cn_riskfree = risk_free_rate("cn")
        # print(cn_riskfree)
        hk_riskfree = get_hk_riskfree()
        # print(hk_riskfree)
        us_inflation = inflation("us")

    with xlwings.App(visible=False) as app:
        marco_book = app.books.open(monitor_path)
        macro_sheet = marco_book.sheets('Macro')
        macro_sheet.range('D6').value = us_riskfree
        macro_sheet.range('F6').value = cn_riskfree
        macro_sheet.range('H6').value = hk_riskfree
        macro_sheet.range('D7').value = us_inflation
        marco_book.save(monitor_path)
        marco_book.close()
    print("Finished Marco data Update")


def read_marco():

    # Marco data
    monitor_wb = load_workbook(macro_monitor_file_path, read_only=True, data_only=True)
    macro_sheet = monitor_wb["Macro"]
    us_riskfree = macro_sheet['D6'].value
    us_bbb_yield = macro_sheet['D3'].value
    us_required_return = macro_sheet['D3'].value
    cn_mos = macro_sheet['F3'].value
    hk_mos = macro_sheet['H3'].value
    us_target = macro_sheet['D10'].value
    cn_target = macro_sheet['F10'].value
    hk_target = macro_sheet['H10'].value

    cn_riskfree = macro_sheet['F6'].value
    hk_riskfree = macro_sheet['H6'].value
    monitor_wb.close()
    return MonitorMarco(us_riskfree)


class MonitorMarco:
    """Monitor class for Macro

    Defines what data can be extracted from the valuation model and used in the Monitor.
    """

    def __init__(self, us_riskfree):
        self.us_riskfree = us_riskfree
