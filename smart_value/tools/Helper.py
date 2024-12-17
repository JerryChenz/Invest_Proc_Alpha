import xlwings
from smart_value.tools.find_docs import get_model_paths


def copy_value():
    """Update the models in the opportunities folder with the latest template

    :param quick: update the price and forex if False
    """

    opportunities_path_list = get_model_paths()

    for p in opportunities_path_list:
        print(f"Updating {p.name}")
        with xlwings.App(visible=False) as app:
            xl_model = app.books.open(p)
            input_sheet = xl_model.sheets('Inputs')
            fin_sheet = xl_model.sheets('Fin_Analysis')
            fin_sheet.range("D3:D4").copy()
            input_sheet.range("C41:C42").paste(paste="value")
            fin_sheet.range("I49").copy()
            input_sheet.range("C37").paste(paste="value")
            xl_model.save(p)
            xl_model.close()
